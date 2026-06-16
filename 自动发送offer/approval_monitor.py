"""
审批流程监听模块
监听审批状态变化，审批通过后自动发送Offer
"""

import time
import logging
import threading
from typing import Dict, Callable, Optional
from datetime import datetime
from approval import ApprovalManager
from offer_sender import OfferSender

logger = logging.getLogger(__name__)


class ApprovalMonitor:
    """审批监听类"""
    
    def __init__(self, approval_manager: ApprovalManager, offer_sender: OfferSender):
        self.approval_manager = approval_manager
        self.offer_sender = offer_sender
        self.monitoring = False
        self.monitor_thread = None
        self.approval_records = {}  # 存储审批记录 {sp_no: onboarding_data}
    
    def add_approval_record(self, sp_no: str, onboarding_data: Dict, applicant_user_id: str):
        """添加审批记录进行监听
        
        Args:
            sp_no: 审批单号
            onboarding_data: 入职信息
            applicant_user_id: 申请人用户ID
        """
        self.approval_records[sp_no] = {
            "onboarding_data": onboarding_data,
            "applicant_user_id": applicant_user_id,
            "create_time": datetime.now(),
            "status": "1",  # 1-审批中
            "notified": False
        }
        logger.info(f"添加审批记录监听: {sp_no}")
    
    def start_monitoring(self, interval: int = 60):
        """启动审批监听
        
        Args:
            interval: 轮询间隔（秒）
        """
        if self.monitoring:
            logger.warning("审批监听已在运行中")
            return
        
        self.monitoring = True
        self.monitor_thread = threading.Thread(
            target=self._monitor_loop,
            args=(interval,),
            daemon=True
        )
        self.monitor_thread.start()
        logger.info(f"审批监听已启动，轮询间隔: {interval}秒")
    
    def stop_monitoring(self):
        """停止审批监听"""
        self.monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=5)
        logger.info("审批监听已停止")
    
    def _monitor_loop(self, interval: int):
        """监听循环"""
        while self.monitoring:
            try:
                self._check_approval_status()
            except Exception as e:
                logger.error(f"审批状态检查异常: {e}")
            
            time.sleep(interval)
    
    def _check_approval_status(self):
        """检查所有审批记录的状态"""
        sp_nos_to_remove = []
        
        for sp_no, record in self.approval_records.items():
            try:
                # 获取审批状态
                status = self.approval_manager.get_approval_status(sp_no)
                
                if status == "0":
                    # 获取状态失败，跳过
                    continue
                
                # 更新状态
                old_status = record["status"]
                record["status"] = status
                
                # 审批通过（状态2）
                if status == "2" and not record["notified"]:
                    logger.info(f"审批单 {sp_no} 已通过，准备发送Offer")
                    self._handle_approval_passed(sp_no, record)
                    record["notified"] = True
                    sp_nos_to_remove.append(sp_no)
                
                # 审批驳回（状态3）
                elif status == "3" and not record["notified"]:
                    logger.info(f"审批单 {sp_no} 已驳回")
                    self._handle_approval_rejected(sp_no, record)
                    record["notified"] = True
                    sp_nos_to_remove.append(sp_no)
                
                # 审批撤销（状态4）
                elif status == "4":
                    logger.info(f"审批单 {sp_no} 已撤销")
                    sp_nos_to_remove.append(sp_no)
                
            except Exception as e:
                logger.error(f"检查审批单 {sp_no} 状态异常: {e}")
        
        # 移除已处理的审批记录
        for sp_no in sp_nos_to_remove:
            del self.approval_records[sp_no]
            logger.info(f"移除已处理的审批记录: {sp_no}")
    
    def _handle_approval_passed(self, sp_no: str, record: Dict):
        """处理审批通过事件
        
        Args:
            sp_no: 审批单号
            record: 审批记录
        """
        try:
            onboarding_data = record["onboarding_data"]
            applicant_user_id = record["applicant_user_id"]
            
            # 发送审批通过通知
            self.offer_sender.send_approval_notification(
                applicant_user_id,
                "approved",
                onboarding_data
            )
            
            # 发送Offer
            offer_result = self.offer_sender.send_offer(
                applicant_user_id,
                onboarding_data,
                send_file=False
            )
            
            if offer_result.get("success"):
                logger.info(f"审批单 {sp_no} Offer发送成功")
            else:
                logger.error(f"审批单 {sp_no} Offer发送失败: {offer_result.get('error')}")
                
        except Exception as e:
            logger.error(f"处理审批通过事件异常: {e}")
    
    def _handle_approval_rejected(self, sp_no: str, record: Dict):
        """处理审批驳回事件
        
        Args:
            sp_no: 审批单号
            record: 审批记录
        """
        try:
            onboarding_data = record["onboarding_data"]
            applicant_user_id = record["applicant_user_id"]
            
            # 发送审批驳回通知
            self.offer_sender.send_approval_notification(
                applicant_user_id,
                "rejected",
                onboarding_data
            )
            
            logger.info(f"审批单 {sp_no} 驳回通知已发送")
            
        except Exception as e:
            logger.error(f"处理审批驳回事件异常: {e}")
    
    def get_monitoring_status(self) -> Dict:
        """获取监听状态"""
        return {
            "monitoring": self.monitoring,
            "approval_count": len(self.approval_records),
            "approval_records": self.approval_records
        }


class ApprovalCallbackHandler:
    """审批回调处理器（用于企业微信回调通知）"""
    
    def __init__(self, approval_manager: ApprovalManager, offer_sender: OfferSender):
        self.approval_manager = approval_manager
        self.offer_sender = offer_sender
    
    def handle_callback(self, callback_data: Dict) -> Dict:
        """处理审批回调
        
        Args:
            callback_data: 回调数据
            
        Returns:
            处理结果
        """
        try:
            # 解析回调数据
            approval_info = callback_data.get("ApprovalInfo", {})
            sp_no = approval_info.get("ThirdNo", "")
            status = approval_info.get("OpenSpStatus", "")
            applicant_user_id = approval_info.get("ApplyUserId", "")
            
            logger.info(f"收到审批回调: sp_no={sp_no}, status={status}, user={applicant_user_id}")
            
            # 审批通过
            if status == "2":
                # 获取审批详情
                detail_result = self.approval_manager.get_approval_detail(sp_no)
                
                if detail_result.get("success"):
                    # 提取入职信息
                    onboarding_data = self._extract_onboarding_data(detail_result["data"])
                    
                    # 发送Offer
                    offer_result = self.offer_sender.send_offer(
                        applicant_user_id,
                        onboarding_data
                    )
                    
                    return {
                        "success": True,
                        "message": "审批通过，Offer已发送"
                    }
            
            # 审批驳回
            elif status == "3":
                onboarding_data = {"name": "候选人"}
                self.offer_sender.send_approval_notification(
                    applicant_user_id,
                    "rejected",
                    onboarding_data
                )
                
                return {
                    "success": True,
                    "message": "审批驳回通知已发送"
                }
            
            return {
                "success": True,
                "message": "回调已处理"
            }
            
        except Exception as e:
            logger.error(f"处理审批回调异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _extract_onboarding_data(self, approval_detail: Dict) -> Dict:
        """从审批详情中提取入职信息"""
        onboarding_data = {}
        
        try:
            apply_data = approval_detail.get("info", {}).get("apply_data", {})
            contents = apply_data.get("contents", [])
            
            for item in contents:
                control = item.get("control", "")
                title = item.get("title", "")
                value = item.get("value", {})
                
                if control == "text":
                    onboarding_data[title] = value.get("text", "")
                elif control == "date":
                    timestamp = value.get("timestamp", 0)
                    onboarding_data[title] = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d")
                elif control == "member":
                    members = value.get("members", [])
                    if members:
                        onboarding_data[title] = members[0].get("userid", "")
        
        except Exception as e:
            logger.error(f"提取入职信息异常: {e}")
        
        return onboarding_data
