"""SerpApi 搜索、缓存和免费额度控制。"""

import json
import time

import certifi
import requests
import ssl

del ssl


class Serp:
    """只通过直连访问 SerpApi，并在请求前执行额度检查。"""

    def __init__(self, data, config, apiKey, checkpoint=None, log=None):
        """初始化免费额度策略、回调和直连请求器。"""
        self.data = data
        self.config = config
        self.apiKey = str(apiKey or "")
        self.checkpoint = checkpoint or (lambda: None)
        self.log = log or (lambda message: None)
        self.searchUrl = str(config.get("serpapiUrl") or "https://serpapi.com/search.json")
        self.accountUrl = "https://serpapi.com/account.json"
        self.remoteAccount = None
        self.remoteAt = 0.0
        self.remoteLocalUsed = 0
        self.headers = {"Accept": "application/json", "User-Agent": "TDI-Automation/1.0"}

    def buildQuery(self, mode, candidate):
        """按公司或个人生成一条紧凑的搜索词。"""
        name = str(candidate.get("name") or "").strip()
        if mode == "person":
            return f'"{name}" insurance'
        city = str(candidate.get("city") or "").strip()
        state = str(candidate.get("state") or "").strip()
        area = f"{city} {state}".strip() if city else (state or "Texas")
        return f'"{name}" insurance agency contact email phone {area}'

    def refreshAccount(self, force=False):
        """直连读取 SerpApi 账户额度，五分钟内复用结果。"""
        if not self.apiKey:
            return {}
        if self.remoteAccount and not force and time.time() - self.remoteAt < 300:
            return dict(self.remoteAccount)
        payload = self.requestJson(self.accountUrl, {"api_key": self.apiKey}, timeout=30)
        self.remoteAccount = {
            "planName": str(payload.get("plan_name") or ""),
            "allowance": int(payload.get("searches_per_month") or 0),
            "remaining": int(payload.get("total_searches_left") or 0),
            "used": int(payload.get("this_month_usage") or 0),
            "renewalDate": str(payload.get("plan_renewal_date") or ""),
            "rateLimit": int(payload.get("account_rate_limit_per_hour") or 0),
        }
        self.remoteLocalUsed = self.data.quotaCount()
        self.remoteAt = time.time()
        return dict(self.remoteAccount)

    def requestJson(self, url, params, timeout=60):
        """使用 certifi 证书包直连 SerpApi，避免打包环境 HTTPS 证书缺失。"""
        requestTimeout = (8, min(int(timeout), 30)) if isinstance(timeout, (int, float)) else timeout
        session = requests.Session()
        session.trust_env = False
        response = session.get(
            url,
            params=params,
            headers=self.headers,
            timeout=requestTimeout,
            verify=certifi.where(),
        )
        response.raise_for_status()
        return response.json()

    def snapshot(self, force=False):
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
            "todayUsed": self.data.quotaCount(dayOnly=True),
        }
        remote = {}
        if force:
            try:
                remote = self.refreshAccount(force=True)
            except Exception as error:
                message = str(error)[:160]
                snapshot["source"] = "本地预算（官方刷新失败）"
                snapshot["remoteError"] = message
                self.log(f"SerpApi 官方额度刷新失败，继续使用本地保守记账：{message[:120]}")
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

    def canSearch(self, category="routine"):
        """检查硬预留、日上限和分类月预算。"""
        snapshot = self.snapshot()
        hardReserve = max(0, int(self.config.get("serpHardReserve", 20)))
        dailyCap = max(1, int(self.config.get("dailySerpCap", 6)))
        if int(snapshot["remaining"]) <= hardReserve:
            return False, f"剩余额度 {snapshot['remaining']} 已到预留线 {hardReserve}"
        if int(snapshot["todayUsed"]) >= dailyCap:
            return False, f"今日 SerpApi 上限 {dailyCap} 次已用完"
        budgets = {"routine": int(self.config.get("serpRoutineBudget", 180))}
        budget = budgets.get(category, 0)
        if budget and self.data.quotaCount(category) >= budget:
            return False, f"{category} 月度预算 {budget} 次已用完"
        return True, "允许搜索"

    def find(self, mode, candidate):
        """优先返回缓存，其次恢复响应，最后发起一次新搜索。"""
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
            allowed, reason = self.canSearch("routine")
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

    def requestSearch(self, query, mode, objectKey, fingerprint):
        """请求 Google 第一页；确定已响应的异常不自动重试。"""
        params = {
            "engine": "google",
            "q": query,
            "api_key": self.apiKey,
            "hl": "en",
            "gl": "us",
            "location": "Texas, United States",
            "start": 0,
        }
        lastError = ""
        for attempt in range(1, 4):
            self.checkpoint()
            try:
                try:
                    payload = self.requestJson(self.searchUrl, params, timeout=60)
                except ValueError as error:
                    self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, True)
                    raise RuntimeError("SerpApi 已响应但返回内容无法解析；为保护额度，本次不自动重试") from error
                if payload.get("error"):
                    self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, True)
                    raise RuntimeError(f"SerpApi 返回错误；为保护额度，本次不自动重试：{payload['error']}")
                searchId = str((payload.get("search_metadata") or {}).get("id") or "")
                self.data.saveSearchPayload(fingerprint, mode, objectKey, query, payload)
                self.data.recordQuota(fingerprint, mode, objectKey, "routine", True, True, searchId)
                self.log(f"SerpApi 搜索完成：{query}")
                return payload
            except requests.HTTPError as error:
                lastError = str(error)
                self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, False)
                statusCode = int(getattr(error.response, "status_code", 0) or 0)
                if 500 <= statusCode < 600 and attempt < 3:
                    self.log(f"SerpApi 服务异常，第 {attempt}/3 次重试：{lastError[:120]}")
                    time.sleep(attempt * 2)
                    continue
                raise RuntimeError(f"SerpApi HTTP 错误：{lastError}") from error
            except (requests.RequestException, TimeoutError) as error:
                lastError = str(error)
                self.data.recordQuota(fingerprint, mode, objectKey, "routine", False, False)
                if attempt < 3:
                    self.log(f"SerpApi 网络失败，第 {attempt}/3 次重试：{lastError[:120]}")
                    time.sleep(attempt * 2)
                    continue
        raise RuntimeError(lastError or "SerpApi 请求失败")
