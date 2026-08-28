"""企业微信内部 Offer 制单应用入口。"""

import json
import logging
import os
import re
import secrets
import shutil
import tempfile
import time
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import config
from flask import Flask, abort, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

from draft import MailDraft, parseAddressList
from jobs import OfferJobs
from offer import OfferPdf
from resume import ResumeExtractor
from settings import Settings


class RunGui:
    """提供企业成员登录、简历识别、Offer PDF和邮件草稿生成。"""

    def __init__(self):
        self.settings = Settings()
        self.jobs = OfferJobs(self.settings.jobDir)
        self.extractor = ResumeExtractor()
        self.offerPdf = OfferPdf(self.settings)
        self.mailDraft = MailDraft(self.settings)
        self.publicUrl = getattr(config, "PUBLIC_URL", "").rstrip("/")
        self.corpId = getattr(config, "CORP_ID", "")
        self.agentId = str(getattr(config, "AGENT_ID", ""))
        self.agentSecret = getattr(config, "AGENT_SECRET", "")
        self.localUserId = getattr(config, "LOCAL_DEBUG_USER_ID", "")
        notifyUsers = getattr(config, "OFFER_NOTIFY_USER_IDS", [])
        if isinstance(notifyUsers, str):
            notifyUsers = [item.strip() for item in re.split(r"[,|]", notifyUsers) if item.strip()]
        self.offerNotifyUserIds = list(dict.fromkeys(notifyUsers))
        self.accessToken = ""
        self.accessTokenExpiresAt = 0
        self.app = Flask(__name__, static_url_path="/offer/static")
        flaskSecretKey = str(getattr(config, "FLASK_SECRET_KEY", "")).strip()
        if not flaskSecretKey:
            raise RuntimeError("缺少 FLASK_SECRET_KEY，请在 config.py 中配置固定随机字符串")
        self.app.secret_key = flaskSecretKey
        self.app.config["MAX_CONTENT_LENGTH"] = self.settings.maxResumeMb * 1024 * 1024
        self.app.context_processor(self.templateContext)
        self.registerRoutes()

    def templateContext(self):
        """向所有页面提供统一公司品牌信息。"""
        return {"companyName": self.settings.companyName}

    def registerRoutes(self):
        """注册企业微信应用页面。"""
        self.app.add_url_rule("/", "home", self.home)
        self.app.add_url_rule("/offer/", "offerHome", self.home)
        self.app.add_url_rule("/offer/auth", "auth", self.auth)
        self.app.add_url_rule("/offer/extract", "extract", self.extract, methods=["POST"])
        self.app.add_url_rule("/offer/review/<jobId>", "review", self.review)
        self.app.add_url_rule("/offer/generate/<jobId>", "generate", self.generate, methods=["POST"])
        self.app.add_url_rule("/offer/done/<jobId>", "done", self.done)
        self.app.add_url_rule("/offer/download/<jobId>/<kind>", "download", self.download)
        self.app.add_url_rule("/healthz", "health", self.health)

    def requireUser(self):
        """仅允许企业微信成员使用，保留本机调试入口。"""
        if session.get("userId"):
            return None
        if request.remote_addr in {"127.0.0.1", "::1"}:
            session["userId"] = self.localUserId or "local-hr"
            return None
        if not all([self.publicUrl, self.corpId, self.agentId, self.agentSecret]):
            abort(503, "企业微信网页授权配置不完整")
        callback = quote(f"{self.publicUrl}/auth", safe="")
        return redirect(
            "https://open.weixin.qq.com/connect/oauth2/authorize?"
            f"appid={quote(self.corpId)}&redirect_uri={callback}&response_type=code&"
            f"scope=snsapi_base&agentid={quote(self.agentId)}#wechat_redirect"
        )

    def auth(self):
        """用企业微信授权码识别当前操作员工。"""
        code = request.args.get("code", "")
        if not code:
            abort(401, "未取得企业微信授权码")
        result = self.wecomGet(
            "https://qyapi.weixin.qq.com/cgi-bin/auth/getuserinfo",
            {"access_token": self.getAccessToken(), "code": code},
        )
        userId = result.get("userid") or result.get("UserId")
        if not userId:
            abort(403, "当前账号不是应用可见范围内的企业成员")
        session["userId"] = userId
        return redirect(url_for("offerHome"))

    def getAccessToken(self):
        """获取并缓存企业微信应用令牌。"""
        if self.accessToken and time.time() < self.accessTokenExpiresAt:
            return self.accessToken
        result = self.wecomGet(
            "https://qyapi.weixin.qq.com/cgi-bin/gettoken",
            {"corpid": self.corpId, "corpsecret": self.agentSecret},
        )
        self.accessToken = result.get("access_token", "")
        if not self.accessToken:
            raise RuntimeError(result.get("errmsg", "企业微信凭证无效"))
        self.accessTokenExpiresAt = time.time() + int(result.get("expires_in", 7200)) - 300
        return self.accessToken

    def wecomGet(self, endpoint, params):
        """调用企业微信只读接口。"""
        with urlopen(f"{endpoint}?{urlencode(params)}", timeout=10) as response:
            result = json.loads(response.read().decode("utf-8"))
        if result.get("errcode", 0) != 0:
            raise RuntimeError(result.get("errmsg", "企业微信接口调用失败"))
        return result

    def sendWecomOfferNotice(self, values, jobId):
        """邮件发送成功后向指定负责人发送企业微信应用通知。"""
        if not self.offerNotifyUserIds:
            return {"notified": False, "error": "未配置企业微信通知负责人UserID"}
        detailUrl = f"{self.publicUrl}/done/{jobId}" if self.publicUrl else ""
        content = (
            "Offer发送完成\n"
            f"候选人：{values['name']}\n"
            f"岗位：{values['position']}\n"
            f"入职部门：{values['department']}\n"
            f"入职日期：{values['entryDate']}\n"
            f"候选人邮箱：{values['email']}\n"
            f"抄送人：{values.get('ccNames') or '未设置'}\n"
            f"操作人：{session.get('userId') or '未知'}"
        )
        payload = {
            "touser": "|".join(self.offerNotifyUserIds),
            "msgtype": "textcard" if detailUrl else "text",
            "agentid": int(self.agentId),
            "safe": 0,
        }
        if detailUrl:
            payload["textcard"] = {
                "title": "Offer发送完成",
                "description": content.replace("\n", "<br>"),
                "url": detailUrl,
                "btntxt": "查看结果",
            }
        else:
            payload["text"] = {"content": content}
        endpoint = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={quote(self.getAccessToken())}"
        requestObject = Request(
            endpoint,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )
        with urlopen(requestObject, timeout=10) as response:
            result = json.loads(response.read().decode("utf-8"))
        if result.get("errcode", 0) != 0:
            raise RuntimeError(result.get("errmsg", "企业微信通知发送失败"))
        invalidUsers = str(result.get("invaliduser", "")).strip()
        invalidParties = str(result.get("invalidparty", "")).strip()
        invalidTags = str(result.get("invalidtag", "")).strip()
        invalidDetails = []
        if invalidUsers:
            invalidDetails.append(f"无效或不在应用可见范围的UserID：{invalidUsers}")
        if invalidParties:
            invalidDetails.append(f"无效部门：{invalidParties}")
        if invalidTags:
            invalidDetails.append(f"无效标签：{invalidTags}")
        if invalidDetails:
            raise RuntimeError("；".join(invalidDetails))
        return {"notified": True, "error": "", "messageId": str(result.get("msgid", ""))}

    def home(self):
        """显示简历上传步骤。"""
        authResult = self.requireUser()
        if authResult:
            return authResult
        return render_template(
            "index.html", stage="upload", maxResumeMb=self.settings.maxResumeMb,
            userId=session.get("userId"),
        )

    def extract(self):
        """识别简历并进入信息复核步骤。"""
        authResult = self.requireUser()
        if authResult:
            return authResult
        uploaded = request.files.get("resume")
        if not uploaded or not uploaded.filename:
            return render_template("index.html", stage="upload", error="请选择简历文件。", maxResumeMb=self.settings.maxResumeMb), 400
        filename = secure_filename(uploaded.filename) or "resume"
        suffix = Path(filename).suffix.lower()
        if suffix not in self.extractor.allowedSuffixes:
            return render_template("index.html", stage="upload", error="文件格式不支持。", maxResumeMb=self.settings.maxResumeMb), 400
        uploadDir = self.settings.dataDir / "uploads"
        uploadDir.mkdir(parents=True, exist_ok=True)
        descriptor, temporaryName = tempfile.mkstemp(suffix=suffix, dir=uploadDir)
        os.close(descriptor)
        Path(temporaryName).unlink(missing_ok=True)
        try:
            uploaded.save(temporaryName)
            extracted = self.extractor.extract(temporaryName)
            jobId = self.jobs.create(filename, extracted)
            shutil.move(temporaryName, self.jobs.path(jobId) / f"resume{suffix}")
        except Exception as exc:
            Path(temporaryName).unlink(missing_ok=True)
            logging.exception("简历识别失败")
            return render_template("index.html", stage="upload", error=str(exc), maxResumeMb=self.settings.maxResumeMb), 400
        return redirect(url_for("review", jobId=jobId))

    def review(self, jobId):
        """显示 Offer 信息复核步骤。"""
        authResult = self.requireUser()
        if authResult:
            return authResult
        job = self.jobs.get(jobId)
        if not job:
            abort(404)
        reportTimes = re.findall(r"\d{1,2}:\d{2}", self.settings.reportTime)
        extracted = dict(job["extracted"])
        extracted["position"] = self.cleanPosition(extracted.get("position", ""))
        values = {
            **extracted, "department": "", "reportPosition": "部门负责人",
            "salaryGrade": "", "probationMonths": "3",
            "probationBase": "", "probationConfidential": "", "probationPerformance": "",
            "regularBase": "", "regularConfidential": "", "regularPerformance": "",
            "entryDate": "", "trialEndDate": "",
            "salaryBank": "中国建设银行", "reportLocation": self.settings.reportLocation,
            "reportStartTime": reportTimes[0] if reportTimes else "09:00",
            "reportEndTime": reportTimes[1] if len(reportTimes) > 1 else "09:30",
            "hrName": self.settings.hrName,
            "hrPhone": self.settings.hrPhone, "responseDays": "1",
            "ccEmails": ", ".join(self.settings.offerCcEmails),
            "ccNames": self.settings.offerCcDisplay,
        }
        values.update(job.get("reviewed") or {})
        values["position"] = self.cleanPosition(values.get("position", ""))
        session.setdefault("csrfToken", secrets.token_urlsafe(32))
        return render_template(
            "index.html", stage="review", job=job, values=values, userId=session.get("userId"),
            departments=self.settings.departments,
            supervisorDepartments=self.settings.supervisorDepartments,
        )

    def renderReview(self, job, values, error, status=400):
        """带完整下拉配置重新显示复核表单。"""
        return render_template(
            "index.html", stage="review", job=job, values=values, error=error,
            userId=session.get("userId"), departments=self.settings.departments,
            supervisorDepartments=self.settings.supervisorDepartments,
        ), status

    def generate(self, jobId):
        """生成留存草稿，并按二次确认结果真实发送 Offer。"""
        authResult = self.requireUser()
        if authResult:
            return authResult
        job = self.jobs.get(jobId)
        if not job:
            abort(404)
        if job.get("status") == "SENT":
            return redirect(url_for("done", jobId=jobId))
        csrfToken = request.form.get("csrfToken", "")
        if not session.get("csrfToken") or not secrets.compare_digest(csrfToken, session["csrfToken"]):
            return redirect(url_for("offerHome"), code=303)
        if request.form.get("sendNow") != "yes":
            return redirect(url_for("offerHome"), code=303)
        fieldNames = [
            "name", "email", "phone", "education", "school", "major", "city",
            "position", "department", "reportPosition", "salaryGrade", "probationMonths",
            "probationBase", "probationConfidential", "probationPerformance",
            "regularBase", "regularConfidential", "regularPerformance",
            "entryDate", "trialEndDate", "salaryBank",
            "reportLocation", "reportStartTime", "reportEndTime",
            "hrName", "hrPhone", "responseDays",
        ]
        values = {name: request.form.get(name, "").strip() for name in fieldNames}
        values["ccEmails"] = ", ".join(self.settings.offerCcEmails)
        values["ccNames"] = self.settings.offerCcDisplay
        values["position"] = self.cleanPosition(values["position"])
        required = [
            "name", "email", "position", "department", "entryDate", "trialEndDate",
            "probationBase", "probationConfidential", "probationPerformance",
            "regularBase", "regularConfidential", "regularPerformance",
            "reportStartTime", "reportEndTime",
        ]
        if any(not values[name] for name in required) or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", values["email"]):
            job["reviewed"] = values
            self.jobs.save(jobId, job)
            return self.renderReview(job, values, "请补齐必填信息，并检查邮箱格式。")
        try:
            values["ccEmails"] = ", ".join(parseAddressList(values["ccEmails"]))
        except ValueError as exc:
            return self.renderReview(job, values, f"固定抄送{exc}", 500)
        try:
            date.fromisoformat(values["entryDate"])
            date.fromisoformat(values["trialEndDate"])
            time.strptime(values["reportStartTime"], "%H:%M")
            time.strptime(values["reportEndTime"], "%H:%M")
        except ValueError:
            return self.renderReview(job, values, "请选择有效的入职日期、试岗结束日期和报到时间。")
        if values["trialEndDate"] < values["entryDate"]:
            return self.renderReview(job, values, "试岗最后日期不能早于入职日期。")
        try:
            values["probationSalary"] = self.buildSalary(values, "probation", "试用期")
            values["regularSalary"] = self.buildSalary(values, "regular", "转正")
        except ValueError as exc:
            job["reviewed"] = values
            self.jobs.save(jobId, job)
            return self.renderReview(job, values, str(exc))
        if values["reportStartTime"] >= values["reportEndTime"]:
            return self.renderReview(job, values, "报到结束时间必须晚于开始时间。")
        values["reportTime"] = f"{values['reportStartTime']}-{values['reportEndTime']}"
        try:
            responseDays = max(1, int(values["responseDays"] or 1))
        except ValueError:
            responseDays = 1
        values["issueDate"] = date.today().strftime("%Y年%m月%d日")
        values["responseDeadline"] = (date.today() + timedelta(days=responseDays)).strftime("%Y年%m月%d日")
        safeName = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9_-]", "", values["name"]) or "candidate"
        outputDir = self.settings.outputDir / jobId
        pdfPath = outputDir / f"{safeName}-录用通知书.pdf"
        emlPath = outputDir / f"{safeName}-录用邮件草稿.eml"
        try:
            self.offerPdf.generate(values, pdfPath)
            draft = self.mailDraft.create(values, pdfPath, emlPath, saveServer=True, sendNow=True)
        except Exception as exc:
            logging.exception("Offer生成失败")
            return self.renderReview(job, values, str(exc), 500)
        wecomNotice = {"notified": False, "error": "邮件未发送，未通知负责人"}
        if draft.get("sent"):
            try:
                wecomNotice = self.sendWecomOfferNotice(values, jobId)
            except Exception as exc:
                logging.exception("企业微信负责人通知失败")
                wecomNotice = {"notified": False, "error": str(exc)}
        draft["wecomNotified"] = wecomNotice["notified"]
        draft["wecomError"] = wecomNotice["error"]
        job.update({
            "status": "SENT" if draft.get("sent") else "GENERATED", "reviewed": values, "pdfPath": str(pdfPath),
            "emlPath": str(emlPath), "draft": draft, "operator": session.get("userId"),
        })
        self.jobs.save(jobId, job)
        return redirect(url_for("done", jobId=jobId))

    def done(self, jobId):
        """显示生成结果。"""
        authResult = self.requireUser()
        if authResult:
            return authResult
        job = self.jobs.get(jobId)
        if not job or job.get("status") not in {"GENERATED", "SENT"}:
            abort(404)
        return render_template("index.html", stage="done", job=job, draft=job.get("draft") or {}, userId=session.get("userId"))

    def download(self, jobId, kind):
        """下载当前任务生成的文件。"""
        authResult = self.requireUser()
        if authResult:
            return authResult
        job = self.jobs.get(jobId)
        if not job or kind not in {"pdf", "eml"}:
            abort(404)
        path = Path(job.get("pdfPath" if kind == "pdf" else "emlPath", ""))
        if not path.exists() or path.parent != self.settings.outputDir / jobId:
            abort(404)
        return send_file(path, as_attachment=True, download_name=path.name)

    def buildSalary(self, values, prefix, label):
        """合计薪酬组成并生成录用通知中的薪酬描述。"""
        keys = [f"{prefix}Base", f"{prefix}Confidential", f"{prefix}Performance"]
        try:
            parts = [round(float(values[key]), 2) for key in keys]
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{label}薪酬必须填写有效金额。") from exc
        if any(amount < 0 for amount in parts):
            raise ValueError(f"{label}薪酬不能填写负数。")
        total = round(sum(parts), 2)
        amount = lambda value: f"{value:,.2f}".rstrip("0").rstrip(".")
        values[f"{prefix}Total"] = amount(total)
        return (
            f"总薪酬 {amount(total)} 元/月（基本工资 {amount(parts[0])} 元/月"
            f" + 保密费 {amount(parts[1])} 元/月 + 绩效 {amount(parts[2])} 元/月）"
        )

    def cleanPosition(self, value):
        """移除简历标签和城市片段，避免识别原文进入 Offer。"""
        position = re.sub(r"^(?:求职意向|意向岗位|应聘职位|目标岗位|期望职位)\s*[：:]?\s*", "", value or "")
        return re.split(r"\s*(?:意向城市|期望城市|所在城市|现居城市)\s*[：:]", position, maxsplit=1)[0].strip()

    def health(self):
        """返回服务状态。"""
        return {"ok": True, "service": "offer_app"}

    def run(self):
        """启动企业微信内部应用。"""
        self.app.run(host=self.settings.host, port=self.settings.port)


portal = RunGui()
app = portal.app


if __name__ == "__main__":
    portal.run()
