"""SerpApi 搜索、缓存和免费额度控制。"""

import json
import time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from data import Data


class Serp:
    """只通过直连访问 SerpApi，并在请求前执行额度检查。"""

    def __init__(
        self,
        data,
        config,
        apiKey,
        checkpoint=None,
        log=None,
        opener=None,
    ):
        """初始化免费额度策略、回调和直连请求器。"""
        self.data = data
        self.config = config
        self.apiKey = str(apiKey or "")
        self.checkpoint = checkpoint or (lambda: None)
        self.log = log or (lambda message: None)
        self.opener = opener or urlopen
        self.searchUrl = str(config.get("serpapiUrl") or "https://serpapi.com/search.json")
        self.accountUrl = "https://serpapi.com/account.json"
        self.remoteAccount = None
        self.remoteAt = 0.0
        self.remoteLocalUsed = 0

    def buildQuery(self, mode, candidate):
        """按公司或个人模式生成一条紧凑搜索词。"""
        name = str(candidate.get("name") or "").strip()
        county = str(candidate.get("county") or "").strip()
        area = f"{county} Texas" if county else "Texas"
        if mode == "company":
            return f'"{name}" real estate contact email phone Facebook {area}'
        return f'"{name}" "real estate agent" contact email phone Facebook {area}'

    def refreshAccount(self, force = False):
        """直连读取 SerpApi 账户额度，五分钟内复用结果。"""
        if not self.apiKey:
            return {}
        if self.remoteAccount and not force and time.time() - self.remoteAt < 300:
            return dict(self.remoteAccount)
        request = Request(
            self.accountUrl + "?" + urlencode({"api_key": self.apiKey}),
            headers={"Accept": "application/json", "User-Agent": "TREC-Automation/3.0"},
        )
        response = self.opener(request, timeout=30)
        payload = json.loads(response.read().decode("utf-8"))
        self.remoteAccount = {
            "planName": str(payload.get("plan_name") or ""),
            "allowance": int(payload.get("searches_per_month") or 0),
            "remaining": int(payload.get("total_searches_left") or 0),
            "used": int(payload.get("this_month_usage") or 0),
            "renewalDate": str(payload.get("plan_renewal_date") or ""),
            "rateLimit": int(payload.get("account_rate_limit_per_hour") or 0),
        }
        # 记录官方快照生成时的本地次数，后续请求从官方剩余额度中实时扣减。
        self.remoteLocalUsed = self.data.quotaCount()
        self.remoteAt = time.time()
        return dict(self.remoteAccount)

    def snapshot(self, force = False):
        """合并本地记账和 SerpApi 官方账户额度。"""
        localUsed = self.data.quotaCount()
        allowance = int(self.config.get("serpMonthlyAllowance", 250))
        snapshot = {
            "allowance": allowance,
            "used": localUsed,
            "remaining": max(0, allowance - localUsed),
            "renewalDate": "",
            "source": "本地预算",
            "routineUsed": self.data.quotaCount("routine"),
            "reviewUsed": self.data.quotaCount("manual_review"),
            "todayUsed": self.data.quotaCount(dayOnly=True),
        }
        remote = {}
        if force:
            try:
                remote = self.refreshAccount(force=True)
            except Exception as error:
                self.log(f"SerpApi 官方额度刷新失败，继续使用本地保守记账：{str(error)[:120]}")
        elif self.remoteAccount:
            remote = dict(self.remoteAccount)
        if remote and int(remote.get("allowance") or 0) > 0:
            addedAfterRefresh = max(0, localUsed - self.remoteLocalUsed)
            snapshot.update({
                "allowance": int(remote["allowance"]),
                "used": int(remote.get("used") or 0) + addedAfterRefresh,
                "remaining": max(0, int(remote.get("remaining") or 0) - addedAfterRefresh),
                "renewalDate": str(remote.get("renewalDate") or ""),
                "source": "SerpApi",
            })
        return snapshot

    def canSearch(self, category, mode, combinedMode = True):
        """检查硬预留、可调整日上限、模式分配和分类月预算。"""
        snapshot = self.snapshot()
        hardReserve = max(0, int(self.config.get("serpHardReserve", 20)))
        dailyCap = max(1, int(self.config.get("dailySerpCap", 6)))
        if int(snapshot["remaining"]) <= hardReserve:
            return False, f"剩余额度 {snapshot['remaining']} 已到预留线 {hardReserve}"
        if int(snapshot["todayUsed"]) >= dailyCap:
            return False, f"今日 SerpApi 上限 {dailyCap} 次已用完"
        if combinedMode:
            modeCap = (dailyCap + 1) // 2 if mode == "company" else dailyCap // 2
            modeUsed = self.data.quotaCount(dayOnly=True, mode=mode)
            if modeUsed >= modeCap:
                modeName = "公司" if mode == "company" else "个人"
                return False, f"{modeName}模式今日分配额度 {modeCap} 次已用完"
        budgets = {
            "routine": int(self.config.get("serpRoutineBudget", 180)),
            "manual_review": int(self.config.get("serpReviewBudget", 20)),
        }
        budget = budgets.get(category, 0)
        if budget and self.data.quotaCount(category) >= budget:
            return False, f"{category} 月度预算 {budget} 次已用完"
        return True, "允许搜索"

    def find(
        self,
        mode,
        candidate,
        combinedMode = True,
    ):
        """优先返回最终缓存，其次恢复响应，最后发起一次新搜索。"""
        objectKey = str(candidate.get("objectKey") or candidate.get("object_key") or "")
        query = self.buildQuery(mode, candidate)
        fingerprint = self.data.queryHash(query, mode, objectKey)
        cached = self.data.searchCache(fingerprint)
        if cached and cached.get("result_json"):
            self.log(f"缓存命中，不消耗额度：{candidate.get('name', '')}")
            result = json.loads(str(cached["result_json"]))
            result["cacheHit"] = True
            return result
        if cached and cached.get("response_json"):
            payload = json.loads(str(cached["response_json"]))
            self.log(f"恢复已保存的 SerpApi 响应：{candidate.get('name', '')}")
        else:
            allowed, reason = self.canSearch("routine", mode, combinedMode)
            if not allowed:
                raise PermissionError(reason)
            if not self.apiKey:
                raise RuntimeError("SerpApi Key 未配置")
            payload = self.requestSearch(query, mode, objectKey, fingerprint)
        return {
            "payload": payload,
            "query": query,
            "queryHash": fingerprint,
            "cacheHit": bool(cached),
        }

    def requestSearch(
        self,
        query,
        mode,
        objectKey,
        fingerprint,
    ):
        """请求 Google 第一页；确定已响应的异常不自动重试。"""
        # 不传 num，沿用 SerpApi Google 第一页默认结果，减少不必要的参数差异。
        params = {
            "engine": "google",
            "q": query,
            "api_key": self.apiKey,
            "hl": "en",
            "gl": "us",
            "location": "Texas, United States",
            "start": 0,
        }
        request = Request(
            self.searchUrl + "?" + urlencode(params),
            headers={"Accept": "application/json", "User-Agent": "TREC-Automation/3.0"},
        )
        lastError = ""
        for attempt in range(1, 4):
            self.checkpoint()
            try:
                response = self.opener(request, timeout=60)
                raw = response.read().decode("utf-8")
                try:
                    payload = json.loads(raw)
                except json.JSONDecodeError as error:
                    self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, True)
                    raise RuntimeError(
                        "SerpApi 已响应但返回内容无法解析；为保护额度，本次不自动重试"
                    ) from error
                if payload.get("error"):
                    self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, True)
                    raise RuntimeError(
                        f"SerpApi 返回错误；为保护额度，本次不自动重试：{payload['error']}"
                    )
                searchId = str((payload.get("search_metadata") or {}).get("id") or "")
                self.data.saveSearchPayload(fingerprint, mode, objectKey, query, payload)
                self.data.recordQuota(
                    fingerprint, mode, objectKey, "routine", True, True, searchId
                )
                self.log(f"SerpApi 搜索完成：{query}")
                return payload
            except HTTPError as error:
                lastError = str(error)
                self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, False)
                if 500 <= int(error.code) < 600 and attempt < 3:
                    self.log(f"SerpApi 服务异常，第 {attempt}/3 次重试：{lastError[:120]}")
                    time.sleep(attempt * 2)
                    continue
                raise RuntimeError(f"SerpApi HTTP 错误：{lastError}") from error
            except (URLError, TimeoutError) as error:
                lastError = str(error)
                self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, False)
                if attempt < 3:
                    self.log(f"SerpApi 网络失败，第 {attempt}/3 次重试：{lastError[:120]}")
                    time.sleep(attempt * 2)
                    continue
        raise RuntimeError(lastError or "SerpApi 请求失败")
