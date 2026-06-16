"""
企业微信API基础模块
提供认证、消息发送等基础功能
"""

import requests
import json
import time
from typing import Dict, Optional, List
import logging
from config import CORP_ID, AGENT_SECRET, AGENT_ID

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class WeComAPI:
    """企业微信API基础类"""
    
    def __init__(self, corp_id: str = CORP_ID, agent_secret: str = AGENT_SECRET, agent_id: int = AGENT_ID):
        self.corp_id = corp_id
        self.agent_secret = agent_secret
        self.agent_id = agent_id
        self.access_token = None
        self.token_expire_time = 0
        self.base_url = "https://qyapi.weixin.qq.com/cgi-bin"
    
    def get_access_token(self) -> str:
        """获取access_token，自动处理过期"""
        # 检查token是否有效（提前5分钟刷新）
        if self.access_token and time.time() < self.token_expire_time - 300:
            return self.access_token
        
        url = f"{self.base_url}/gettoken"
        params = {
            "corpid": self.corp_id,
            "corpsecret": self.agent_secret
        }
        
        try:
            response = requests.get(url, params=params, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                self.access_token = result.get("access_token")
                self.token_expire_time = time.time() + result.get("expires_in", 7200)
                logger.info("成功获取access_token")
                return self.access_token
            else:
                logger.error(f"获取access_token失败: {result}")
                raise Exception(f"获取access_token失败: {result.get('errmsg')}")
        except Exception as e:
            logger.error(f"获取access_token异常: {e}")
            raise
    
    def send_text_message(self, user_id: str, content: str, safe: int = 0) -> Dict:
        """发送文本消息给指定用户
        
        Args:
            user_id: 用户ID
            content: 消息内容
            safe: 是否保密消息，0表示可对外分享，1表示不能分享且内容显示水印
            
        Returns:
            API响应结果
        """
        token = self.get_access_token()
        url = f"{self.base_url}/message/send?access_token={token}"
        
        data = {
            "touser": user_id,
            "msgtype": "text",
            "agentid": self.agent_id,
            "text": {
                "content": content
            },
            "safe": safe
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info(f"成功发送文本消息给用户 {user_id}")
            else:
                logger.error(f"发送文本消息失败: {result}")
            
            return result
        except Exception as e:
            logger.error(f"发送文本消息异常: {e}")
            raise
    
    def send_file_message(self, user_id: str, media_id: str, safe: int = 0) -> Dict:
        """发送文件消息给指定用户
        
        Args:
            user_id: 用户ID
            media_id: 媒体文件ID（需要先上传文件获取）
            safe: 是否保密消息
            
        Returns:
            API响应结果
        """
        token = self.get_access_token()
        url = f"{self.base_url}/message/send?access_token={token}"
        
        data = {
            "touser": user_id,
            "msgtype": "file",
            "agentid": self.agent_id,
            "file": {
                "media_id": media_id
            },
            "safe": safe
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info(f"成功发送文件消息给用户 {user_id}")
            else:
                logger.error(f"发送文件消息失败: {result}")
            
            return result
        except Exception as e:
            logger.error(f"发送文件消息异常: {e}")
            raise
    
    def upload_file(self, file_path: str, file_type: str = "file") -> str:
        """上传文件获取media_id
        
        Args:
            file_path: 文件路径
            file_type: 文件类型（image/voice/video/file）
            
        Returns:
            media_id
        """
        token = self.get_access_token()
        url = f"{self.base_url}/media/upload?access_token={token}&type={file_type}"
        
        try:
            with open(file_path, 'rb') as f:
                files = {'media': f}
                response = requests.post(url, files=files, timeout=30)
                result = response.json()
                
                if result.get("errcode") == 0:
                    media_id = result.get("media_id")
                    logger.info(f"成功上传文件，media_id: {media_id}")
                    return media_id
                else:
                    logger.error(f"上传文件失败: {result}")
                    raise Exception(f"上传文件失败: {result.get('errmsg')}")
        except Exception as e:
            logger.error(f"上传文件异常: {e}")
            raise
    
    def send_markdown_message(self, user_id: str, content: str) -> Dict:
        """发送markdown消息给指定用户
        
        Args:
            user_id: 用户ID
            content: markdown格式内容
            
        Returns:
            API响应结果
        """
        token = self.get_access_token()
        url = f"{self.base_url}/message/send?access_token={token}"
        
        data = {
            "touser": user_id,
            "msgtype": "markdown",
            "agentid": self.agent_id,
            "markdown": {
                "content": content
            }
        }
        
        try:
            response = requests.post(url, json=data, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info(f"成功发送markdown消息给用户 {user_id}")
            else:
                logger.error(f"发送markdown消息失败: {result}")
            
            return result
        except Exception as e:
            logger.error(f"发送markdown消息异常: {e}")
            raise
    
    def get_user_info(self, user_id: str) -> Dict:
        """获取用户信息
        
        Args:
            user_id: 用户ID
            
        Returns:
            用户信息
        """
        token = self.get_access_token()
        url = f"{self.base_url}/user/get?access_token={token}&userid={user_id}"
        
        try:
            response = requests.get(url, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info(f"成功获取用户 {user_id} 的信息")
                return result
            else:
                logger.error(f"获取用户信息失败: {result}")
                raise Exception(f"获取用户信息失败: {result.get('errmsg')}")
        except Exception as e:
            logger.error(f"获取用户信息异常: {e}")
            raise
    
    def get_department_list(self, id: Optional[int] = None) -> Dict:
        """获取部门列表
        
        Args:
            id: 部门ID，不传则获取全量
            
        Returns:
            部门列表
        """
        token = self.get_access_token()
        url = f"{self.base_url}/department/list?access_token={token}"
        if id:
            url += f"&id={id}"
        
        try:
            response = requests.get(url, timeout=10)
            result = response.json()
            
            if result.get("errcode") == 0:
                logger.info("成功获取部门列表")
                return result
            else:
                logger.error(f"获取部门列表失败: {result}")
                raise Exception(f"获取部门列表失败: {result.get('errmsg')}")
        except Exception as e:
            logger.error(f"获取部门列表异常: {e}")
            raise
