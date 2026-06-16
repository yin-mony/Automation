"""
企业微信审批申请模块
提供审批申请提交、查询等功能
"""

import json
import logging
from typing import Dict, Optional, List
from datetime import datetime
from wecom_api import WeComAPI
from config import APPROVAL_TEMPLATE_ID

logger = logging.getLogger(__name__)


class ApprovalManager:
    """审批管理类"""
    
    def __init__(self, wecom_api: WeComAPI, template_id: str = APPROVAL_TEMPLATE_ID):
        self.wecom_api = wecom_api
        self.template_id = template_id
    
    def submit_onboarding_approval(self, applicant_user_id: str, onboarding_data: Dict) -> Dict:
        """提交入职审批申请
        
        Args:
            applicant_user_id: 申请人用户ID
            onboarding_data: 入职信息字典，包含：
                - name: 姓名
                - department: 部门
                - position: 职位
                - entry_date: 入职日期
                - phone: 联系电话
                - email: 邮箱
                - education: 学历
                - salary: 薪资
                - recruiter: 招聘负责人
                - notes: 备注
                
        Returns:
            审批申请结果
        """
        token = self.wecom_api.get_access_token()
        url = f"{self.wecom_api.base_url}/oa/applyevent?access_token={token}"
        
        # 构建审批申请数据
        approval_data = {
            "creator_userid": applicant_user_id,
            "template_id": self.template_id,
            "use_template_approver": 1,  # 使用模板中的审批人
            "apply_data": {
                "contents": self._build_approval_contents(onboarding_data)
            },
            "summary_list": self._build_summary_list(onboarding_data)
        }
        
        try:
            response = requests.post(url, json=approval_data, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info(f"成功提交入职审批申请，申请ID: {result.get('sp_no')}")
                return {
                    "success": True,
                    "sp_no": result.get("sp_no"),
                    "sp_name": result.get("sp_name"),
                    "apply_time": datetime.now().isoformat()
                }
            else:
                logger.error(f"提交入职审批申请失败: {result}")
                return {
                    "success": False,
                    "error": result.get("errmsg"),
                    "errcode": result.get("errcode")
                }
        except Exception as e:
            logger.error(f"提交入职审批申请异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _build_approval_contents(self, onboarding_data: Dict) -> List[Dict]:
        """构建审批申请内容"""
        contents = []
        
        # 姓名字段
        if onboarding_data.get("name"):
            contents.append({
                "control": "text",
                "id": "name",
                "title": "姓名",
                "value": {
                    "text": onboarding_data["name"]
                }
            })
        
        # 部门字段
        if onboarding_data.get("department"):
            contents.append({
                "control": "text",
                "id": "department",
                "title": "部门",
                "value": {
                    "text": onboarding_data["department"]
                }
            })
        
        # 职位字段
        if onboarding_data.get("position"):
            contents.append({
                "control": "text",
                "id": "position",
                "title": "职位",
                "value": {
                    "text": onboarding_data["position"]
                }
            })
        
        # 入职日期字段
        if onboarding_data.get("entry_date"):
            contents.append({
                "control": "date",
                "id": "entry_date",
                "title": "入职日期",
                "value": {
                    "timestamp": int(datetime.strptime(onboarding_data["entry_date"], "%Y-%m-%d").timestamp())
                }
            })
        
        # 联系电话字段
        if onboarding_data.get("phone"):
            contents.append({
                "control": "text",
                "id": "phone",
                "title": "联系电话",
                "value": {
                    "text": onboarding_data["phone"]
                }
            })
        
        # 邮箱字段
        if onboarding_data.get("email"):
            contents.append({
                "control": "text",
                "id": "email",
                "title": "邮箱",
                "value": {
                    "text": onboarding_data["email"]
                }
            })
        
        # 学历字段
        if onboarding_data.get("education"):
            contents.append({
                "control": "text",
                "id": "education",
                "title": "学历",
                "value": {
                    "text": onboarding_data["education"]
                }
            })
        
        # 薪资字段
        if onboarding_data.get("salary"):
            contents.append({
                "control": "text",
                "id": "salary",
                "title": "薪资",
                "value": {
                    "text": str(onboarding_data["salary"])
                }
            })
        
        # 招聘负责人字段
        if onboarding_data.get("recruiter"):
            contents.append({
                "control": "member",
                "id": "recruiter",
                "title": "招聘负责人",
                "value": {
                    "members": [
                        {
                            "userid": onboarding_data["recruiter"]
                        }
                    ]
                }
            })
        
        # 备注字段
        if onboarding_data.get("notes"):
            contents.append({
                "control": "text",
                "id": "notes",
                "title": "备注",
                "value": {
                    "text": onboarding_data["notes"]
                }
            })
        
        return contents
    
    def _build_summary_list(self, onboarding_data: Dict) -> List[Dict]:
        """构建审批摘要列表"""
        summary_list = []
        
        if onboarding_data.get("name"):
            summary_list.append({
                "summary_info": [
                    {
                        "text": {
                            "content": f"姓名: {onboarding_data['name']}",
                            "lang": "zh_CN"
                        }
                    }
                ]
            })
        
        if onboarding_data.get("position"):
            summary_list.append({
                "summary_info": [
                    {
                        "text": {
                            "content": f"职位: {onboarding_data['position']}",
                            "lang": "zh_CN"
                        }
                    }
                ]
            })
        
        if onboarding_data.get("department"):
            summary_list.append({
                "summary_info": [
                    {
                        "text": {
                            "content": f"部门: {onboarding_data['department']}",
                            "lang": "zh_CN"
                        }
                    }
                ]
            })
        
        return summary_list
    
    def get_approval_detail(self, sp_no: str) -> Dict:
        """获取审批申请详情
        
        Args:
            sp_no: 审批单号
            
        Returns:
            审批详情
        """
        token = self.wecom_api.get_access_token()
        url = f"{self.wecom_api.base_url}/oa/getapprovaldetail?access_token={token}"
        
        data = {
            "sp_no": sp_no
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info(f"成功获取审批单 {sp_no} 的详情")
                return {
                    "success": True,
                    "data": result
                }
            else:
                logger.error(f"获取审批详情失败: {result}")
                return {
                    "success": False,
                    "error": result.get("errmsg"),
                    "errcode": result.get("errcode")
                }
        except Exception as e:
            logger.error(f"获取审批详情异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def get_approval_status(self, sp_no: str) -> str:
        """获取审批状态
        
        Args:
            sp_no: 审批单号
            
        Returns:
            审批状态：1-审批中，2-已通过，3-已驳回，4-已撤销
        """
        detail_result = self.get_approval_detail(sp_no)
        
        if detail_result.get("success"):
            approval_info = detail_result["data"].get("info", {})
            return approval_info.get("sp_status", "0")
        
        return "0"


import requests
