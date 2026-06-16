"""
Offer自动发送模块
审批通过后自动发送Offer给候选人
"""

import os
import logging
from typing import Dict, Optional
from datetime import datetime
from wecom_api import WeComAPI
from config import OFFER_TEMPLATE_PATH

logger = logging.getLogger(__name__)


class OfferSender:
    """Offer发送类"""
    
    def __init__(self, wecom_api: WeComAPI, template_path: str = OFFER_TEMPLATE_PATH):
        self.wecom_api = wecom_api
        self.template_path = template_path
    
    def generate_offer_content(self, onboarding_data: Dict) -> str:
        """生成Offer内容
        
        Args:
            onboarding_data: 入职信息字典
            
        Returns:
            Markdown格式的Offer内容
        """
        name = onboarding_data.get("name", "候选人")
        department = onboarding_data.get("department", "")
        position = onboarding_data.get("position", "")
        entry_date = onboarding_data.get("entry_date", "")
        salary = onboarding_data.get("salary", "")
        recruiter = onboarding_data.get("recruiter", "")
        
        offer_content = f"""# 🎉 恭喜您通过面试！

尊敬的 **{name}**：

我们很高兴地通知您，经过公司面试评估，您已成功获得 **{department}** 的 **{position}** 职位录用机会！

## 📋 入职信息

- **入职部门**：{department}
- **入职职位**：{position}
- **入职日期**：{entry_date}
- **薪资待遇**：{salary}
- **招聘负责人**：{recruiter}

## 📝 入职须知

1. 请在入职当天携带以下材料：
   - 身份证原件及复印件
   - 学历学位证书原件及复印件
   - 离职证明（如有）
   - 银行卡（用于工资发放）
   - 一寸免冠照片2张

2. 入职时间：{entry_date} 上午9:00
   - 入职地点：公司前台
   - 联系人：HR部门

3. 如有疑问，请联系招聘负责人或HR部门。

## 🎊 欢迎加入我们的团队！

我们期待您的加入，共同创造更美好的未来！

---
此邮件由系统自动发送，请勿直接回复。
如有疑问请联系HR部门。
"""
        return offer_content
    
    def send_offer(self, user_id: str, onboarding_data: Dict, send_file: bool = False) -> Dict:
        """发送Offer给用户
        
        Args:
            user_id: 用户ID
            onboarding_data: 入职信息字典
            send_file: 是否发送文件形式的Offer
            
        Returns:
            发送结果
        """
        try:
            # 生成Offer内容
            offer_content = self.generate_offer_content(onboarding_data)
            
            # 发送Markdown消息
            result = self.wecom_api.send_markdown_message(user_id, offer_content)
            
            if result.get("errcode") == 0:
                logger.info(f"成功发送Offer给用户 {user_id}")
                
                # 如果需要发送文件
                if send_file and os.path.exists(self.template_path):
                    try:
                        media_id = self.wecom_api.upload_file(self.template_path)
                        file_result = self.wecom_api.send_file_message(user_id, media_id)
                        
                        if file_result.get("errcode") == 0:
                            logger.info(f"成功发送Offer文件给用户 {user_id}")
                        else:
                            logger.warning(f"发送Offer文件失败: {file_result}")
                    except Exception as e:
                        logger.warning(f"发送Offer文件异常: {e}")
                
                return {
                    "success": True,
                    "message": "Offer发送成功",
                    "user_id": user_id,
                    "send_time": datetime.now().isoformat()
                }
            else:
                logger.error(f"发送Offer失败: {result}")
                return {
                    "success": False,
                    "error": result.get("errmsg"),
                    "errcode": result.get("errcode")
                }
        except Exception as e:
            logger.error(f"发送Offer异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def send_approval_notification(self, user_id: str, approval_status: str, onboarding_data: Dict) -> Dict:
        """发送审批状态通知
        
        Args:
            user_id: 用户ID
            approval_status: 审批状态（approved/rejected）
            onboarding_data: 入职信息字典
            
        Returns:
            发送结果
        """
        name = onboarding_data.get("name", "")
        position = onboarding_data.get("position", "")
        
        if approval_status == "approved":
            content = f"""# ✅ 入职审批已通过

尊敬的 **{name}**：

您的入职申请（职位：{position}）已通过审批！

我们将尽快为您发送正式Offer，请留意查收。

如有疑问，请联系HR部门。
"""
        else:
            content = f"""# ❌ 入职审批未通过

尊敬的 **{name}**：

很遗憾通知您，您的入职申请（职位：{position}）未通过审批。

如有疑问，请联系HR部门了解详情。
"""
        
        try:
            result = self.wecom_api.send_markdown_message(user_id, content)
            
            if result.get("errcode") == 0:
                logger.info(f"成功发送审批状态通知给用户 {user_id}")
                return {
                    "success": True,
                    "message": "审批状态通知发送成功"
                }
            else:
                logger.error(f"发送审批状态通知失败: {result}")
                return {
                    "success": False,
                    "error": result.get("errmsg")
                }
        except Exception as e:
            logger.error(f"发送审批状态通知异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
