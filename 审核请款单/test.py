import os
import time
from pathlib import Path
from DrissionPage import ChromiumPage
from DrissionPage import ChromiumOptions
from pywinauto import Application

from SaihuERPLogin import SaihuERPLogin


def connect_edge_with_ap():
    debug_addr = os.getenv("EDGE_DEBUG_ADDR", "127.0.0.1:9222").strip()
    co = ChromiumOptions().set_address(debug_addr)
    page = ChromiumPage(co)
    page.set.window.max()
    return page, debug_addr


def main(mode="recent_7_days", start_date="", end_date="", username=None, password=None, export_dir=""):
    page, debug_addr = connect_edge_with_ap()
    print(f"已连接 Edge 实例: {debug_addr}", flush=True)




    # host_down = page.ele('x://wujie-app', timeout=5)
    # shadow_down = host_down.shadow_root
    #
    # down_btn = shadow_down.ele('x://span[text()="下载请款单及附件"]', timeout=5)
    # print(down_btn.text,99999)
    # # print(down_btn.text)
    # # print(down_btn.attr('class'))
    # time.sleep(19999)


    login = SaihuERPLogin(page, username=username, password=password)
    login.login(prefer_entry_check=True)
    print("赛狐页面登录流程完成，当前登录态已保持。", flush=True)

    # 关闭通知广告弹窗
    if page.ele('x://button/span[contains(text(),"下一条")]', timeout=1.2):
        print("关闭通知广告弹窗", flush=True)
        for _ in range(3):
            next_btn = page.ele('x://button/span[contains(text(),"下一条")]', timeout=1.2)
            if not next_btn:
                break
            try:
                next_btn.click()
                time.sleep(0.5)
            except Exception:
                break
        close_btn = page.ele('x://button[2]/span[contains(text(),"关闭")]', timeout=1.2)
        if close_btn:
            try:
                close_btn.click()
                time.sleep(0.5)
            except Exception:
                pass

    # 进入请款单审核页面
    page.ele('x://div/ul/li/span[text()="财务"]', timeout=5).click()
    time.sleep(1)
    page.ele('x://a[text()="请款单"]', timeout=5).click()
    time.sleep(1)

    host = page.ele('x://wujie-app', timeout=5)
    shadow = host.shadow_root
    shadow.ele('x://input[@placeholder="开始日期"]', timeout=5).click()
    # time.sleep(1)

    # 根据 mode 判断点击哪个时间按钮
    if mode == "this_month":
        shadow.ele('x://div/button[text()="本月"]', timeout=5).click()
    elif mode == "last_month":
        shadow.ele('x://div/button[text()="上月"]', timeout=5).click()
    elif mode == "recent_7_days":
        shadow.ele('x://div/button[text()="最近7天"]', timeout=5).click()
    elif mode == "recent_30_days":
        shadow.ele('x://div/button[text()="最近30天"]', timeout=5).click()
    else:
        shadow.ele('x://div/button[text()="最近7天"]', timeout=5).click()
    time.sleep(1)
    # shadow.ele('x://div/button[text()="最近7天"]', timeout=5).click()

    # 设置遍历页数
    Page = 30
    Total = shadow.ele('x://span[contains(@class, "sf_pagination_total_num")]').text
    total_pages = (int(Total) + int(Page) - 1) // int(Page)  # 向上取整
    print(total_pages)
    # 设置第一页默认操作，进行点击
    header = shadow.ele('x://span[@title="全选/取消"]',timeout=10)
    header.scroll.to_see()
    rect = header.rect
    x = rect.location[0] + rect.size[0] / 2
    y = rect.location[1] + rect.size[1] / 2
    # 点击全选/取消
    page.actions.move_to((x, y)).click()

    # 点击下载
    shadow.ele('x://span[text()="打印/下载"]', timeout=5).click()
    shadow.ele('x://span[@class="el-cascader-node__label"]/span[text()="下载请款单及附件"]', timeout=5).click()
    time.sleep(1)
    # 连接到已打开的Chrome下载窗口
    app = Application(backend="uia").connect(title_re=".*赛狐ERP.*")

    # 通过title参数找到并点击
    dlg = app.window(title_re=".*赛狐ERP.*")
    radio_button = dlg.child_window(title="下载请款单 ", control_type="RadioButton")
    radio_button.click()
    time.sleep(1)
    download_button = dlg.child_window(title="下载", control_type="Button")
    download_button.click()
    # 保持浏览器自身下载路径，不使用 downloads_done()（该方法依赖显式下载目录）
    time.sleep(100)
    # time.sleep(10)


    # html = shadow.html
    # with open('header_text.txt', 'w', encoding='utf-8') as f:
    #     f.write(html)
    # 从第二页开始遍历
    for page_num in range(2, total_pages + 1):
        shadow.ele('x://button[@aria-label = "下一页"]', timeout=5).click()
        time.sleep(1)
        # 换页后默认操作，进行点击
        header = shadow.ele('x://span[@title="全选/取消"]', timeout=10)
        header.scroll.to_see()
        rect = header.rect
        x = rect.location[0] + rect.size[0] / 2
        y = rect.location[1] + rect.size[1] / 2
        # 点击全选/取消
        page.actions.move_to((x, y)).click()

        # 点击下载
        shadow.ele('x://span[text()="打印/下载"]', timeout=5).click()
        shadow.ele('x://span[@class="el-cascader-node__label"]/span[text()="下载请款单及附件"]', timeout=5).click()
        time.sleep(1)
        # 连接到已打开的Chrome下载窗口
        app = Application(backend="uia").connect(title_re=".*赛狐ERP.*")

        # 通过title参数找到并点击
        dlg = app.window(title_re=".*赛狐ERP.*")
        radio_button = dlg.child_window(title="下载请款单 ", control_type="RadioButton")
        radio_button.click()
        time.sleep(1)
        download_button = dlg.child_window(title="下载", control_type="Button")
        download_button.click()
        time.sleep(100)
        # time.sleep(10)
    time.sleep(1)
    # shadow.ele('x://div/button/span[text()="导出"]', timeout=10).click(by_js=True)
    # time.sleep(1)



    # 下载操作
    # if mode == "this_month":
    #     mode_text = "本月"
    # elif mode == "last_month":
    #     mode_text = "上个月"
    # elif mode == "recent_30_days":
    #     mode_text = "最近30天"
    # else:
    #     mode_text = "最近7天"
    # file_name = f"请款单-{mode_text}"

    # save_path = export_dir or str(Path.home() / "Desktop")

    # download = page.ele('x://span[contains(text(), "立即下载")]', timeout=10).click.to_download(
    #     save_path=save_path,
    #     rename=file_name,
    #     timeout=120,
    # )
    # downloaded_file = download.wait()
    # print(f"已完成下载文件路径: {downloaded_file}", flush=True)

    # shadow.ele('x://div/button[text()="上月"]', timeout=5).click()
    # shadow.ele('x://div/button[text()="最近7天"]', timeout=5).click()
    # shadow.ele('x://div/button[text()="最近30天"]', timeout=5).click()


    # 保持脚本常驻，避免主进程退出后影响自动化后续步骤。
    # while True:
    #     time.sleep(60)





if __name__ == "__main__":
    main()
