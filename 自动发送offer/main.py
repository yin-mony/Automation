"""
企业微信入职申请-审批通过-自动发送offer系统
主程序入口
"""

import sys
import logging
import argparse
from datetime import datetime
from QiYeVxLogin import QiYeVxLogin
from wecom_api import WeComAPI
from approval import ApprovalManager
from offer_sender import OfferSender
from approval_monitor import ApprovalMonitor, ApprovalCallbackHandler
from config import CORP_ID, AGENT_SECRET, AGENT_ID, APPROVAL_TEMPLATE_ID

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('./logs/app.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class OnboardingSystem:
    """入职管理系统"""
    
    def __init__(self):
        self.wecom_api = None
        self.approval_manager = None
        self.offer_sender = None
        self.approval_monitor = None
        self.qiye_login = None
    
    def initialize(self):
        """初始化系统"""
        try:
            # 初始化企业微信API
            self.wecom_api = WeComAPI(CORP_ID, AGENT_SECRET, AGENT_ID)
            logger.info("企业微信API初始化成功")
            
            # 初始化审批管理器
            self.approval_manager = ApprovalManager(self.wecom_api, APPROVAL_TEMPLATE_ID)
            logger.info("审批管理器初始化成功")
            
            # 初始化Offer发送器
            self.offer_sender = OfferSender(self.wecom_api)
            logger.info("Offer发送器初始化成功")
            
            # 初始化审批监听器
            self.approval_monitor = ApprovalMonitor(self.approval_manager, self.offer_sender)
            logger.info("审批监听器初始化成功")
            
            # 初始化企业微信登录
            self.qiye_login = QiYeVxLogin()
            logger.info("企业微信登录管理器初始化成功")
            
            return True
        except Exception as e:
            logger.error(f"系统初始化失败: {e}")
            return False
    
    def submit_onboarding_application(self, applicant_user_id: str, onboarding_data: dict) -> dict:
        """提交入职申请
        
        Args:
            applicant_user_id: 申请人用户ID
            onboarding_data: 入职信息字典
            
        Returns:
            提交结果
        """
        try:
            logger.info(f"开始提交入职申请，申请人: {applicant_user_id}")
            
            # 提交审批申请
            result = self.approval_manager.submit_onboarding_approval(
                applicant_user_id,
                onboarding_data
            )
            
            if result.get("success"):
                sp_no = result.get("sp_no")
                logger.info(f"入职申请提交成功，审批单号: {sp_no}")
                
                # 添加到监听列表
                self.approval_monitor.add_approval_record(
                    sp_no,
                    onboarding_data,
                    applicant_user_id
                )
                
                return {
                    "success": True,
                    "sp_no": sp_no,
                    "message": "入职申请提交成功，等待审批"
                }
            else:
                logger.error(f"入职申请提交失败: {result.get('error')}")
                return {
                    "success": False,
                    "error": result.get("error")
                }
        except Exception as e:
            logger.error(f"提交入职申请异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def start_monitoring(self, interval: int = 60):
        """启动审批监听
        
        Args:
            interval: 轮询间隔（秒）
        """
        try:
            self.approval_monitor.start_monitoring(interval)
            logger.info("审批监听已启动")
        except Exception as e:
            logger.error(f"启动审批监听失败: {e}")
    
    def stop_monitoring(self):
        """停止审批监听"""
        try:
            self.approval_monitor.stop_monitoring()
            logger.info("审批监听已停止")
        except Exception as e:
            logger.error(f"停止审批监听失败: {e}")
    
    def send_offer_directly(self, user_id: str, onboarding_data: dict) -> dict:
        """直接发送Offer（不经过审批流程）
        
        Args:
            user_id: 用户ID
            onboarding_data: 入职信息字典
            
        Returns:
            发送结果
        """
        try:
            logger.info(f"直接发送Offer给用户: {user_id}")
            result = self.offer_sender.send_offer(user_id, onboarding_data)
            return result
        except Exception as e:
            logger.error(f"直接发送Offer异常: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def ensure_qiye_login(self, timeout: int = 300) -> bool:
        """确保企业微信已登录
        
        Args:
            timeout: 超时时间（秒）
            
        Returns:
            是否登录成功
        """
        try:
            logger.info("检查企业微信登录状态...")
            result = self.qiye_login.ensure_login(timeout=timeout)
            
            if result:
                logger.info("企业微信登录成功")
            else:
                logger.error("企业微信登录失败")
            
            return result
        except Exception as e:
            logger.error(f"企业微信登录异常: {e}")
            return False


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='企业微信入职申请-审批通过-自动发送offer系统')
    parser.add_argument('--action', type=str, required=True,
                       choices=['submit', 'monitor', 'send', 'login'],
                       help='执行的操作：submit-提交申请，monitor-启动监听，send-直接发送offer，login-确保登录')
    parser.add_argument('--user-id', type=str, help='用户ID')
    parser.add_argument('--name', type=str, help='姓名')
    parser.add_argument('--department', type=str, help='部门')
    parser.add_argument('--position', type=str, help='职位')
    parser.add_argument('--entry-date', type=str, help='入职日期(YYYY-MM-DD)')
    parser.add_argument('--phone', type=str, help='联系电话')
    parser.add_argument('--email', type=str, help='邮箱')
    parser.add_argument('--education', type=str, help='学历')
    parser.add_argument('--salary', type=str, help='薪资')
    parser.add_argument('--recruiter', type=str, help='招聘负责人用户ID')
    parser.add_argument('--notes', type=str, help='备注')
    parser.add_argument('--interval', type=int, default=60, help='监听轮询间隔（秒）')
    parser.add_argument('--timeout', type=int, default=300, help='登录超时时间（秒）')
    
    args = parser.parse_args()
    
    # 初始化系统
    system = OnboardingSystem()
    if not system.initialize():
        logger.error("系统初始化失败，程序退出")
        sys.exit(1)
    
    # 根据操作类型执行相应功能
    if args.action == 'submit':
        # 提交入职申请
        if not args.user_id:
            logger.error("提交申请需要指定 --user-id 参数")
            sys.exit(1)
        
        onboarding_data = {
            "name": args.name or "",
            "department": args.department or "",
            "position": args.position or "",
            "entry_date": args.entry_date or "",
            "phone": args.phone or "",
            "email": args.email or "",
            "education": args.education or "",
            "salary": args.salary or "",
            "recruiter": args.recruiter or "",
            "notes": args.notes or ""
        }
        
        result = system.submit_onboarding_application(args.user_id, onboarding_data)
        
        if result.get("success"):
            logger.info(f"入职申请提交成功，审批单号: {result.get('sp_no')}")
            # 自动启动监听
            system.start_monitoring(args.interval)
            logger.info("审批监听已启动，按Ctrl+C停止")
            try:
                import time
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                logger.info("用户中断，停止监听")
                system.stop_monitoring()
        else:
            logger.error(f"入职申请提交失败: {result.get('error')}")
            sys.exit(1)
    
    elif args.action == 'monitor':
        # 启动审批监听
        logger.info("启动审批监听...")
        system.start_monitoring(args.interval)
        logger.info("审批监听已启动，按Ctrl+C停止")
        try:
            import time
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            logger.info("用户中断，停止监听")
            system.stop_monitoring()
    
    elif args.action == 'send':
        # 直接发送Offer
        if not args.user_id:
            logger.error("发送Offer需要指定 --user-id 参数")
            sys.exit(1)
        
        onboarding_data = {
            "name": args.name or "候选人",
            "department": args.department or "",
            "position": args.position or "",
            "entry_date": args.entry_date or "",
            "phone": args.phone or "",
            "email": args.email or "",
            "education": args.education or "",
            "salary": args.salary or "",
            "recruiter": args.recruiter or "",
            "notes": args.notes or ""
        }
        
        result = system.send_offer_directly(args.user_id, onboarding_data)
        
        if result.get("success"):
            logger.info("Offer发送成功")
        else:
            logger.error(f"Offer发送失败: {result.get('error')}")
            sys.exit(1)
    
    elif args.action == 'login':
        # 确保企业微信登录
        result = system.ensure_qiye_login(args.timeout)
        
        if result:
            logger.info("企业微信登录成功")
        else:
            logger.error("企业微信登录失败")
            sys.exit(1)


if __name__ == '__main__':
    main()
