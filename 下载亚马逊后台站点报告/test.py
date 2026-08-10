"""下载亚马逊后台站点报告测试专用入口。"""

from main import AmazonReport, reportTypes, siteMap, siteModes


class Test:
    """测试入口：展示配置支持项，必要时执行真实浏览器流程。"""

    def buildConfig(self):
        """集中维护测试配置，正式运行前按本机店铺环境填写。"""
        return {
            "username": "",
            "password": "",
            "yidekeUsername": "",
            "yidekePassword": "",
            "autoSiteName": "美国",
            "ip": [""],
            "port": [8888],
            "shopIp": [""],
            "shopPort": [8888],
            "amazonEmail": "",
            "amazonPassword": "",
            "amazonSiteName": "美国",
            "amazonSiteNames": list(siteMap.keys()),
            "siteMode": "single",
            "reportType": reportTypes["summary"]["label"],
            "isOnline": False,
            "runBrowser": False,
        }

    def showSupportedOptions(self):
        """打印当前测试入口支持的站点、模式和报告类型。"""
        print("支持的亚马逊后台站点：")
        for siteName, englishName in siteMap.items():
            print(f"- {siteName}: {englishName}")
        print("易得客店铺站点与 Amazon 后台站点使用同一套站点映射。")
        print("支持的站点切换模式：")
        for key, label in siteModes.items():
            print(f"- {label} ({key})")
        print("支持的报告类型：")
        for report in reportTypes.values():
            print(f"- {report['label']} ({report['optionValue']})")

    def main(self):
        """默认只做配置展示，开启 runBrowser 后执行真实浏览器流程。"""
        config = self.buildConfig()
        self.showSupportedOptions()
        if not config.get("runBrowser"):
            print("当前 runBrowser=False，仅展示测试配置，不启动易得客或浏览器。")
            print("需要真实测试时，请填写账号、密码、店铺 IP，并把 runBrowser 改为 True。")
            return
        AmazonReport(config).run()


if __name__ == "__main__":
    Test().main()
