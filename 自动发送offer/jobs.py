"""人事制单任务的本地存储。"""

import json
import secrets
from datetime import datetime
from pathlib import Path


class OfferJobs:
    """按随机任务编号保存简历、识别结果和最终文件。"""

    def __init__(self, jobDir):
        self.jobDir = Path(jobDir)
        self.jobDir.mkdir(parents=True, exist_ok=True)

    def create(self, filename, extracted):
        """创建任务并返回任务编号。"""
        jobId = secrets.token_hex(8)
        data = {
            "jobId": jobId,
            "filename": filename,
            "status": "EXTRACTED",
            "extracted": extracted,
            "reviewed": {},
            "createdAt": datetime.now().isoformat(timespec="seconds"),
        }
        self.save(jobId, data)
        return jobId

    def path(self, jobId):
        """返回任务目录并阻止路径穿越。"""
        if not jobId or any(character not in "0123456789abcdef" for character in jobId):
            raise ValueError("无效任务编号")
        path = self.jobDir / jobId
        path.mkdir(parents=True, exist_ok=True)
        return path

    def save(self, jobId, data):
        """保存任务 JSON。"""
        path = self.path(jobId) / "job.json"
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def get(self, jobId):
        """读取任务，不存在时返回 None。"""
        path = self.path(jobId) / "job.json"
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))
