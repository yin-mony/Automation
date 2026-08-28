from main import Factory1688


def main():
    """使用默认配置启动1688筛选工厂流程。"""
    config = {
        # 待搜索图片目录，目录名会作为材质名称。
        "folderPath": "./file/不锈钢镀金",
        # 结果表输出目录；留空时输出到图片目录的上一级目录。
        "outputDir": "",
        # 最终导出的Excel文件名。
        "outputFileName": "1688筛选工厂结果.xlsx",
        # 等待人工完成主体框选并确认的最长秒数。
        "cropTimeout": 300
    }

    factory = Factory1688(config)
    factory.run()


if __name__ == "__main__":
    main()
