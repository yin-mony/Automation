import os
import socket
import subprocess
import time
from pathlib import Path

import pandas as pd
import psutil
from DrissionPage import Chromium, ChromiumPage

from YidekeLogin import YidekeLogin


class Auto:
    """易得客进店、Amazon 后台登录与评论下载主流程"""

    def __init__(self, config=None):
        # 外部传入配置，为空时使用可调试默认值
        config = config or {}
        # 原始运行配置
        self.config = config
        # 易得客账号密码
        self.yidekeUsername = config.get("yidekeUsername") or config.get("username") or config.get("yideke_username") or ""
        self.yidekePassword = config.get("yidekePassword") or config.get("password") or config.get("yideke_password") or ""
        # 店铺站点，中文用于易得客区域，英文用于 Amazon Seller 账号切换
        self.siteName = str(config.get("autoSiteName") or config.get("siteName") or "美国").strip()
        self.siteEnglishMap = {
            "美国": "United States",
            "加拿大": "Canada",
            "墨西哥": "Mexico",
            "巴西": "Brazil",
            "英国": "United Kingdom",
            "法国": "France",
            "德国": "Germany",
            "意大利": "Italy",
            "西班牙": "Spain",
            "荷兰": "Netherlands",
            "瑞典": "Sweden",
            "波兰": "Poland",
            "比利时": "Belgium",
            "爱尔兰": "Ireland",
            "日本": "Japan",
            "新加坡": "Singapore",
            "澳大利亚": "Australia",
            "印度": "India",
            "阿联酋": "United Arab Emirates",
            "沙特阿拉伯": "Saudi Arabia",
            "土耳其": "Turkey",
            "埃及": "Egypt",
            "南非": "South Africa",
        }
        self.siteEnglishName = self.siteEnglishMap.get(self.siteName, self.siteName)
        self.amazonHomeMap = {
            "美国": "https://www.amazon.com",
            "加拿大": "https://www.amazon.ca",
            "墨西哥": "https://www.amazon.com.mx",
            "巴西": "https://www.amazon.com.br",
            "英国": "https://www.amazon.co.uk",
            "法国": "https://www.amazon.fr",
            "德国": "https://www.amazon.de",
            "意大利": "https://www.amazon.it",
            "西班牙": "https://www.amazon.es",
            "荷兰": "https://www.amazon.nl",
            "瑞典": "https://www.amazon.se",
            "波兰": "https://www.amazon.pl",
            "比利时": "https://www.amazon.com.be",
            "爱尔兰": "https://www.amazon.ie",
            "日本": "https://www.amazon.co.jp",
            "新加坡": "https://www.amazon.sg",
            "澳大利亚": "https://www.amazon.com.au",
            "印度": "https://www.amazon.in",
            "阿联酋": "https://www.amazon.ae",
            "沙特阿拉伯": "https://www.amazon.sa",
            "土耳其": "https://www.amazon.com.tr",
            "埃及": "https://www.amazon.eg",
            "南非": "https://www.amazon.co.za",
        }
        self.amazonHome = self.amazonHomeMap.get(self.siteName, "https://www.amazon.com")
        # 店铺 IP 与调试端口列表
        shopIp = config.get("shopIp") or config.get("shop_ip") or config.get("ip") or []
        self.shopIp = shopIp if isinstance(shopIp, list) else [shopIp]
        self.shopIp = [str(item).strip() for item in self.shopIp if str(item or "").strip()]
        shopPort = config.get("shopPort") or config.get("shop_port") or config.get("port") or []
        self.shopPort = shopPort if isinstance(shopPort, list) else [shopPort]
        self.shopPort = [int(item) for item in self.shopPort if str(item or "").strip()]
        # 易得客管理窗口调试端口
        self.controlPort = int(config.get("controlPort") or 9222)
        # Amazon Seller 登录信息
        self.amazonEmail = config.get("amazonEmail") or config.get("amazon_email") or ""
        self.amazonPassword = config.get("amazonPassword") or config.get("amazon_password") or ""
        # 商品 ASIN 列表
        asinList = config.get("asinList") or config.get("experts") or config.get("asin") or []
        self.asinList = asinList if isinstance(asinList, list) else [asinList]
        self.asinList = [str(item).strip() for item in self.asinList if str(item or "").strip()]
        # 评论保存目录
        self.filePath = config.get("filePath") or config.get("file_path") or ""
        # 易得客安装根目录
        self.edeckerBase = Path.home() / "AppData/Local/eDecker6"
        # 易得客登录实例
        self.login = None
        # 当前店铺页面
        self.page = None

    def stopProgram(self):
        """强制结束浏览器并退出程序"""
        # 遍历浏览器相关进程并强制关闭
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                name = (proc.info["name"] or "").lower()
                if name in {"chrome.exe", "edecker.exe"}:
                    proc.kill()
                    print(f"已终止进程: {proc.info['name']} (PID: {proc.info['pid']})")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                # 进程已退出或权限不足时跳过
                pass
        os._exit(0)

    def killEdecker(self, excludePid=None):
        """结束除指定 PID 外的所有 edecker 进程"""
        # 遍历并关闭不需要保留的易得客进程
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                pid = proc.info["pid"]
                name = proc.info["name"]
                if name and name.lower() == "edecker.exe" and pid != excludePid:
                    proc.kill()
            except Exception:
                # 进程状态变化时跳过
                pass

    def killEdeckerPort(self, port):
        """启动店铺浏览器前结束占用该调试端口的 edecker 进程"""
        # 拼接调试端口参数
        flag = f"--remote-debugging-port={port}"
        # 遍历易得客进程并按命令行匹配端口
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                name = proc.info["name"]
                cmdline = proc.info.get("cmdline") or []
                if name and name.lower() == "edecker.exe" and any(flag in str(arg) for arg in cmdline):
                    proc.kill()
            except Exception:
                # 进程信息读取失败时跳过
                pass

    def waitPort(self, port, timeout=60):
        """等待店铺浏览器调试端口就绪"""
        # 按超时时间持续探测端口
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=2):
                    return
            except OSError:
                time.sleep(1)

        raise RuntimeError(f"等待 127.0.0.1:{port} 超时 ({timeout}s)")

    def resolveEdecker(self):
        """复用易得客登录模块的查找策略定位 edecker.exe"""
        # 创建临时登录对象用于查找可执行文件
        login = YidekeLogin(self.yidekeUsername, self.yidekePassword)
        # 返回当前机器上的易得客可执行文件
        return login.resolveEdecker()

    def visitShop(self, ipList, port=9222):
        """在易得客管理窗口中按店铺 IP 点击访问"""
        # 接管易得客管理浏览器
        browser = Chromium(port)
        tab = browser.latest_tab
        try:
            for tabId in browser.tab_ids:
                candidate = browser.get_tab(tabId)
                url = (candidate.url or "").lower()
                if (
                    "selleros.cn" in url
                    or "work-station" in url
                    or "shops.edecker.cn" in url
                    or "workbench" in url
                ):
                    tab = candidate
                    break
        except Exception:
            # 标签读取失败时保留 latest_tab 继续尝试
            pass
        time.sleep(2)

        for ip in ipList:
            # 在指定站点店铺卡片中按 IP 点击访问
            print(f"正在易得客店铺列表中访问 {self.siteName} 店铺 IP: {ip}")
            tab.ele(
                f'x://div[contains(@class,"platform-region")]//span[normalize-space()="{self.siteName}"]'
                f'/ancestor::div[contains(@class,"shop-item")]'
                f'[.//div[contains(@class,"text") and normalize-space()="{ip}"]]'
                f'//button[normalize-space()="访问"]',
                timeout=30,
            ).click()
            time.sleep(3)

        # 关闭管理窗口，只保留已打开的店铺浏览器
        self.killEdecker(browser.process_id)
        time.sleep(1)
        tab.refresh()
        time.sleep(3)

    def startEdecker(self, ip, port):
        """按店铺 IP 匹配 profile 并用指定端口启动浏览器"""
        # 定位易得客可执行文件和 profile 目录
        exePath = Path(self.resolveEdecker())
        profilesPath = self.edeckerBase / "Profiles"

        # 校验易得客可执行文件
        if not exePath.exists():
            raise FileNotFoundError(f"找不到 exe: {exePath}")
        # 校验 profile 目录
        if not profilesPath.exists():
            raise FileNotFoundError(f"找不到 profiles 目录: {profilesPath}")

        # 兼容点号 IP 和下划线 IP 两种 profile 命名
        ipDot = ip
        ipUnderline = ip.replace(".", "_")
        candidates = [
            path for path in profilesPath.iterdir()
            if path.is_dir() and (ipDot in path.name or ipUnderline in path.name)
        ]
        if not candidates:
            raise RuntimeError(f"未找到 IP={ip} 的 profile")

        # 使用最近更新的 profile
        latest = max(candidates, key=lambda path: path.stat().st_mtime)
        cmd = [
            str(exePath),
            f"--user-data-dir={latest}",
            "--no-sandbox",
            f"--remote-debugging-port={port}",
        ]
        print(f"启动店铺浏览器: IP={ip}, port={port}, profile={latest}")
        subprocess.Popen(cmd, cwd=str(self.edeckerBase))

    def openShop(self):
        """登录易得客并触发目标店铺访问"""
        # 创建登录对象并执行易得客登录
        self.login = YidekeLogin(self.yidekeUsername, self.yidekePassword)
        self.login.run()
        time.sleep(3)

        # 在管理窗口中访问全部目标店铺
        self.visitShop(self.shopIp, self.controlPort)
        time.sleep(4)

    def loginSeller(self, page):
        """按需执行 Amazon Seller 后台登录"""
        # 未填写 Amazon 后台账号时，仅复用易得客 profile 的既有登录态
        if not self.amazonEmail and not self.amazonPassword:
            print("未填写 Amazon 后台账号，跳过 Seller Central 登录步骤。")
            return page

        # 委托易得客工具中的 Amazon Seller 登录逻辑
        print(f"开始登录 Amazon Seller Central，站点: {self.siteEnglishName}")
        sellerPage = self.login.loginSeller(
            page,
            email=self.amazonEmail,
            password=self.amazonPassword,
            siteName=self.siteEnglishName,
        )
        print(f"Amazon Seller Central 登录后 URL: {sellerPage.url}")
        return sellerPage

    def clickFresh(self, page, locator, timeout=30, scroll=True):
        """页面局部刷新时重新定位元素后点击"""
        # 先定位元素并按需滚动到可视区域
        element = page.ele(locator, timeout=timeout)
        if scroll:
            page.scroll.to_see(element)
            time.sleep(1)

        try:
            element.click()
            return
        except Exception:
            # Amazon 页面常在滚动后替换节点，失效时重新定位并用 JS 点击
            element = page.ele(locator, timeout=timeout)
            element.click(by_js=True)

    def scrapeProfile(self, ip, port):
        """接管单个店铺 profile 浏览器并抓取全部 ASIN 评论"""
        # 启动前清理同端口旧店铺浏览器
        self.killEdeckerPort(port)
        time.sleep(1)
        # 启动指定店铺 profile
        self.startEdecker(ip, port)
        # 等待 DrissionPage 可接管
        self.waitPort(port)
        page = ChromiumPage(f"127.0.0.1:{port}")

        try:
            # 最大化窗口，方便页面元素加载
            page.set.window.max()
        except RuntimeError:
            # 易得客部分版本不支持最大化，不影响采集
            pass

        # Amazon 后台登录辅助，确保店铺 profile 已进入正确账号
        self.page = self.loginSeller(page)

        # Amazon 评论筛选项
        stars = ["5 star only", "4 star only", "3 star only", "2 star only", "1 star only"]
        # 导出用中文星级名
        starNames = ["5星", "4星", "3星", "2星", "1星"]
        # 当前店铺评论结果
        profileComments = {}

        # 新建 Amazon 前台首页标签，后续按 ASIN 抓取评论
        page = page.new_tab(self.amazonHome)
        time.sleep(3)
        page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=25).click()

        for asin in self.asinList:
            # 初始化当前 ASIN 评论容器
            print(f"\n正在处理店铺 {ip} 商品: {asin}")
            profileComments[asin] = {}

            # 搜索并进入商品详情页
            page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=30).input(f"{asin}\n", clear=True)
            time.sleep(3)
            page.ele(f'x://div[@data-asin="{asin}"]//a', timeout=30).click()
            time.sleep(3)

            # 进入更多评论页面
            self.clickFresh(page, 'x://div[text()="See more reviews"]', timeout=30)
            time.sleep(3)
            page.ele('x://span[text() = "All stars"]', timeout=30).click()
            time.sleep(3)

            for star, starName in zip(stars, starNames):
                # 切换到指定星级评论
                print(f"\n正在处理 {starName}...")
                page.ele(f'x://ul[@role="listbox"]/li/a[text()="{star}"]', timeout=30).click()
                time.sleep(3)

                # 持续点击加载更多评论
                while True:
                    commentBtn = page.ele('x://span/a[text()="Show 10 more reviews"]', timeout=0)
                    if not commentBtn:
                        print("没有更多评论了")
                        break
                    commentBtn.click()
                    time.sleep(2)
                    print(f"点击加载更多{starName}评论...")

                # 提取当前星级全部评论正文
                reviews = []
                for elem in page.eles('x://span[@data-hook="review-body"]/span'):
                    text = elem.text.strip()
                    if text:
                        reviews.append(text)
                profileComments[asin][starName] = reviews

                # 重新打开星级下拉，为下一档做准备
                page.ele(f'x://span[text() = "{star}"]', timeout=30).click()
                time.sleep(3)

                print(f"{starName} 共提取 {len(reviews)} 条评论")
                print("-" * 50)

            # 回到搜索框，准备下一个 ASIN
            page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=25).click()
            # 打印当前 ASIN 评论汇总，便于日志排查
            self.printComments(profileComments[asin])

        return profileComments

    def printComments(self, starComments):
        """打印当前 ASIN 各星级评论日志"""
        # 遍历各星级评论并输出明细
        for starName, reviews in starComments.items():
            print(f"\n{'=' * 50}")
            print(f"{starName} 评论 (共 {len(reviews)} 条)")
            print(f"{'=' * 50}")
            for index, review in enumerate(reviews, 1):
                print(f"{index}. {review}")

    def mergeComments(self, allComments, profileComments):
        """将单店铺抓取结果合并到总结果中"""
        # 按 ASIN 与星级逐层合并评论列表
        for asin, starComments in profileComments.items():
            targetAsin = allComments.setdefault(asin, {})
            for starName, reviews in starComments.items():
                targetAsin.setdefault(starName, []).extend(reviews)

    def exportExcel(self, allComments):
        """将评论结果导出为亚马逊评论.xlsx"""
        # 展平评论数据为 Excel 行
        data = []
        for asin, starComments in allComments.items():
            for star, reviews in starComments.items():
                for review in reviews:
                    data.append({
                        "ASIN": asin,
                        "星级": star,
                        "评论内容": review,
                        "评论长度": len(review),
                    })

        # 创建保存目录
        outputDir = Path(self.filePath)
        outputDir.mkdir(parents=True, exist_ok=True)
        outputPath = outputDir / "亚马逊评论.xlsx"

        # 写入 Excel 文件
        df = pd.DataFrame(data)
        df.to_excel(outputPath, index=False)
        print(f"已保存 {len(data)} 条评论到 {outputPath}")
        return outputPath

    def run(self):
        """执行完整评论下载流程"""
        # 校验店铺 IP 与端口数量一致
        if len(self.shopIp) != len(self.shopPort):
            raise ValueError(f"IP 数量 ({len(self.shopIp)}) 与端口数量 ({len(self.shopPort)}) 不一致")

        # 登录易得客并打开店铺浏览器
        self.openShop()
        # 汇总所有店铺评论
        allComments = {}
        for ip, port in zip(self.shopIp, self.shopPort):
            profileComments = self.scrapeProfile(ip, port)
            self.mergeComments(allComments, profileComments)

        # 导出最终 Excel
        self.exportExcel(allComments)
        return allComments


if __name__ == "__main__":
    # 本文件独立调试配置
    config = {
        "yidekeUsername": "",
        "yidekePassword": "",
        "autoSiteName": "美国",
        "shopIp": [],
        "shopPort": [],
        "amazonEmail": "",
        "amazonPassword": "",
        "experts": [],
        "file_path": r"C:\RPA流程\亚马逊评论分析\flie",
    }

    # 未配置关键参数时只提示调试方式，避免误触发真实浏览器流程
    if not config["yidekeUsername"] or not config["yidekePassword"] or not config["shopIp"]:
        print("请在 main.py 的 main 配置中填写易得客账号、密码、店铺 IP、端口和 ASIN 后再调试。")
    else:
        service = Auto(config)
        service.run()


# =============================================================================
# 旧版 main.py 主流程代码归档
# =============================================================================
#
# 以下注释保留旧版评论下载主流程的关键代码，后续如需对照旧逻辑，可直接查看本段。
# 当前正式实现位于 main.py：Auto.run()
#
# class Comment:
#     """易得客 + Amazon 评论抓取与 Excel 导出。"""
#
#     def __init__(self, config):
#         self.username = config["username"]
#         self.password = config["password"]
#         ips = config["ip"]
#         self.ip = ips if isinstance(ips, list) else [ips]
#         ports = config["port"]
#         self.port = ports if isinstance(ports, list) else [ports]
#         self.experts = config["experts"]
#         self.file_path = config["file_path"]
#
#     def main(self):
#         """逐个 ASIN 抓取 1～5 星评论，汇总后调用 excel_files 导出。"""
#         sp = Specification(self.username, self.password)
#         time.sleep(5)
#         sp.YidekeLogin()
#         time.sleep(3)
#         self.run_edecker_automation(self.ip)
#         time.sleep(4)
#
#         for index, ip in enumerate(self.ip):
#             self.kill_edecker_on_port(self.port[index])
#             time.sleep(1)
#             self.start_edecker(self.ip[index], self.port[index])
#             self.wait_for_port(self.port[index])
#             page = ChromiumPage("127.0.0.1:" + str(self.port[index]))
#
#             try:
#                 page.set.window.max()
#             except RuntimeError:
#                 pass
#
#             stars = ["5 star only", "4 star only", "3 star only", "2 star only", "1 star only"]
#             star_names = ["5星", "4星", "3星", "2星", "1星"]
#             all_comments = {}
#
#             page = page.new_tab("https://www.amazon.com")
#             time.sleep(3)
#             page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=25).click()
#
#             for experts in self.experts:
#                 print(f"\n正在处理商品: {experts}")
#                 all_comments[experts] = {}
#                 page.ele(
#                     'x://div/input[@placeholder="Search Amazon"]',
#                     timeout=30
#                 ).input(f'{experts}\n', clear=True)
#                 time.sleep(3)
#                 page.ele(f'x://div[@data-asin="{experts}"]//a', timeout=30).click()
#                 time.sleep(3)
#
#                 ele = page.ele('x://div[text()="See more reviews"]')
#                 page.scroll.to_see(ele)
#                 time.sleep(3)
#                 page.ele('x://div[text()="See more reviews"]').click()
#                 time.sleep(3)
#                 page.ele('x://span[text() = "All stars"]').click()
#                 time.sleep(3)
#
#                 for star, star_name in zip(stars, star_names):
#                     print(f"\n正在处理 {star_name}...")
#                     page.ele(f'x://ul[@role="listbox"]/li/a[text()="{star}"]').click()
#                     time.sleep(3)
#
#                     while True:
#                         comment_button = page.ele(
#                             'x://span/a[text()="Show 10 more reviews"]',
#                             timeout=0
#                         )
#                         if not comment_button:
#                             print("没有更多评论了")
#                             break
#                         comment_button.click()
#                         time.sleep(2)
#                         print(f"点击加载更多{star_name}评论...")
#
#                     reviews = []
#                     for elem in page.eles('x://span[@data-hook="review-body"]/span'):
#                         text = elem.text.strip()
#                         if text:
#                             reviews.append(text)
#                     all_comments[experts][star_name] = reviews
#
#                     page.ele(f'x://span[text() = "{star}"]').click()
#                     time.sleep(3)
#                     print(f"{star_name} 共提取 {len(reviews)} 条评论")
#                     print("-" * 50)
#
#                 page.ele('x://div/input[@placeholder="Search Amazon"]', timeout=25).click()
#
#             self.excel_files(all_comments)
#             return all_comments
#
#     def excel_files(self, all_comments):
#         """将 all_comments 展平为行，写入 {file_path}/亚马逊评论.xlsx。"""
#         data = []
#         for asin, star_comments in all_comments.items():
#             for star, reviews in star_comments.items():
#                 for review in reviews:
#                     data.append({
#                         'ASIN': asin,
#                         '星级': star,
#                         '评论内容': review,
#                         '评论长度': len(review)
#                     })
#
#         df = pd.DataFrame(data)
#         output_dir = Path(self.file_path)
#         output_dir.mkdir(parents=True, exist_ok=True)
#         output_path = output_dir / "亚马逊评论.xlsx"
#         df.to_excel(output_path, index=False)
#         print(f"已保存 {len(data)} 条评论到 {output_path}")
#         return output_path
