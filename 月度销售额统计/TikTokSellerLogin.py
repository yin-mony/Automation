import json
import time
from pathlib import Path

_DBG_LOG = Path(__file__).resolve().parent.parent / "debug-38661f.log"


def _agent_dbg(hypothesis_id, location, message, data=None):
    # #region agent log
    try:
        entry = {
            "sessionId": "38661f",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data or {},
            "timestamp": int(time.time() * 1000),
        }
        with _DBG_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass
    # #endregion


class TikTokSellerLogin:
    """TikTok Shop 卖家中心登录；滑块验证码需人工完成。"""

    def __init__(self, page, on_captcha=None):
        self.page = page
        self.on_captcha = on_captcha
        self._captcha_notified = False

    def is_login_page(self, page=None):
        page = page or self.page
        url = (page.url or '').lower()
        if '/account/login' in url:
            return True
        try:
            return bool(page.ele('x://button[contains(., "Continue")]', timeout=2))
        except Exception:
            return False

    def _is_backend_url(self, url):
        url = (url or '').lower()
        if '/account/login' in url:
            return False
        return any(
            k in url
            for k in ('seller-us.tiktok.com', 'seller.tiktok', 'tiktokglobalshop')
        )

    def is_seller_backend(self, page=None):
        page = page or self.page
        return self._is_backend_url(page.url)

    def is_logged_in(self, page=None):
        page = page or self.page
        url = (page.url or '').lower()
        if '/account/login' in url:
            return False
        if self.is_seller_backend(page):
            return True
        try:
            return bool(page.ele('x://div[text()="Analytics"]', timeout=2))
        except Exception:
            return False

    def _pick_seller_tab(self, page):
        tab_urls = []
        backend_tab = None
        login_tab = None
        fallback_tab = None
        try:
            browser = page.browser
            for tab_id in browser.tab_ids:
                tab = browser.get_tab(tab_id)
                url = (tab.url or '').lower()
                tab_urls.append(url[:120])
                if url.startswith('chrome-extension://'):
                    continue
                if self._is_backend_url(url):
                    backend_tab = tab
                    break
                if '/account/login' in url:
                    login_tab = tab
                    continue
                if any(k in url for k in ('seller', 'tiktokglobalshop', 'tiktok.com')):
                    fallback_tab = tab
            picked = backend_tab or login_tab or fallback_tab
            if picked:
                self.page = picked
                # #region agent log
                _agent_dbg("B", "TikTokSellerLogin._pick_seller_tab", "picked seller tab", {
                    "tab_urls": tab_urls,
                    "picked_url": picked.url,
                    "pick_type": "backend" if picked is backend_tab else ("login" if picked is login_tab else "fallback"),
                })
                # #endregion
                return picked
        except Exception as e:
            # #region agent log
            _agent_dbg("B", "TikTokSellerLogin._pick_seller_tab", "exception", {"error": str(e), "tab_urls": tab_urls})
            # #endregion
            pass
        # #region agent log
        _agent_dbg("B", "TikTokSellerLogin._pick_seller_tab", "no seller tab found", {"tab_urls": tab_urls, "current_url": page.url})
        # #endregion
        return page

    def _resolve_seller_page(self, timeout=60):
        page = self.page
        deadline = time.time() + timeout
        while time.time() < deadline:
            page = self._pick_seller_tab(page)
            if self.is_login_page(page) or self.is_logged_in(page) or self.is_seller_backend(page):
                return page
            url = (page.url or '').lower()
            if not url.startswith('chrome-extension://'):
                return page
            time.sleep(2)
        return self._pick_seller_tab(page)

    def get_active_page(self):
        """返回已切换到卖家后台的标签页。"""
        self.page = self._pick_seller_tab(self.page)
        return self.page

    def login(self, email, password, captcha_timeout=300, timeout=300):
        # #region agent log
        _agent_dbg("C", "TikTokSellerLogin.login", "entry", {"initial_url": self.page.url})
        # #endregion
        page = self._resolve_seller_page(timeout=30)
        # #region agent log
        _agent_dbg("E", "TikTokSellerLogin.login", "after resolve", {"resolved_url": page.url, "is_logged_in": self.is_logged_in(page), "is_login_page": self.is_login_page(page)})
        # #endregion

        if self.has_img_code(page):
            self.handle_captcha(page, timeout=captcha_timeout)
            page = self._pick_seller_tab(page)

        if self.is_logged_in(page) and not self.is_login_page(page):
            print('未检测到登录页，判断为已登录')
            if self.has_img_code(page):
                self.handle_captcha(page, timeout=captcha_timeout)
            if self.is_logged_in(page):
                print('TikTok 卖家后台已就绪')
                # #region agent log
                _agent_dbg("C", "TikTokSellerLogin.login", "early return ready", {"self_page_url": self.page.url})
                # #endregion
                return self.get_active_page()

        if self.is_logged_in(page) and not self.has_img_code(page):
            print('TikTok 已登录，跳过登录步骤')
            # #region agent log
            _agent_dbg("C", "TikTokSellerLogin.login", "early return skip", {"self_page_url": self.page.url})
            # #endregion
            return self.get_active_page()

        if not self.is_login_page(page):
            deadline = time.time() + 30
            while time.time() < deadline:
                page = self._pick_seller_tab(page)
                if self.is_login_page(page) or self.is_logged_in(page):
                    break
                time.sleep(2)

        if self.is_logged_in(page) and not self.is_login_page(page):
            print('未检测到登录页，判断为已登录')
            if self.has_img_code(page):
                self.handle_captcha(page, timeout=captcha_timeout)
            return self._wait_logged_in(page, timeout=timeout, captcha_timeout=captcha_timeout) or self.get_active_page()

        if not self.is_login_page(page):
            page = self._resolve_seller_page(timeout=60)
            if self.is_logged_in(page) and not self.is_login_page(page):
                print('未检测到登录页，判断为已登录')
                if self.has_img_code(page):
                    self.handle_captcha(page, timeout=captcha_timeout)
                return self._wait_logged_in(page, timeout=timeout, captcha_timeout=captcha_timeout) or self.get_active_page()
            if not self.is_login_page(page):
                raise RuntimeError(f'未进入 TikTok 卖家后台或登录页，当前 URL: {page.url}')

        print('检测到 TikTok 登录页，自动填写账号...')
        page = self._pick_seller_tab(page)
        if self.is_logged_in(page) and not self.is_login_page(page):
            print('已切换到卖家后台，跳过登录填写')
            if self.has_img_code(page):
                self.handle_captcha(page, timeout=captcha_timeout)
            return self.get_active_page()

        if self.has_img_code(page):
            self.handle_captcha(page, timeout=captcha_timeout)
            page = self._pick_seller_tab(page)
            if self.is_logged_in(page) and not self.is_login_page(page):
                return self.get_active_page()

        email_ele = self._find_login_email(page, timeout=30)
        if not email_ele:
            page = self._pick_seller_tab(page)
            if self.is_logged_in(page) and not self.is_login_page(page):
                print('未找到登录框，但卖家后台已就绪')
                return self.get_active_page()
            # #region agent log
            _agent_dbg("F", "TikTokSellerLogin.login", "email input not found", {"url": page.url})
            # #endregion
            raise RuntimeError('未找到 TikTok 邮箱输入框')
        email_ele.input(email, clear=True)

        pwd_ele = page.ele('x://input[@type="password"]', timeout=30)
        if not pwd_ele:
            raise RuntimeError('未找到 TikTok 密码输入框')
        pwd_ele.input(password, clear=True)

        continue_btn = page.ele('x://button[contains(., "Continue")]', timeout=30)
        if not continue_btn:
            raise RuntimeError('未找到 TikTok Continue 按钮')
        continue_btn.click()
        time.sleep(3)

        if self.has_img_code(page):
            self.handle_captcha(page, timeout=captcha_timeout)

        self._wait_logged_in(page, timeout=timeout, captcha_timeout=captcha_timeout)
        return self.get_active_page()

    def _find_login_email(self, page, timeout=30):
        selectors = (
            'x://input[@type="email"]',
            'x://input[@name="email"]',
            'x://input[@autocomplete="email"]',
            'x://input[contains(@placeholder,"mail")]',
            'x://input[contains(@placeholder,"Email")]',
            'x://input[contains(@placeholder,"phone")]',
        )
        deadline = time.time() + timeout
        while time.time() < deadline:
            for selector in selectors:
                try:
                    ele = page.ele(selector, timeout=1)
                    if ele:
                        return ele
                except Exception:
                    pass
            time.sleep(1)
        return None

    def handle_captcha(self, page=None, timeout=300):
        """处理验证码：可关闭则尝试关闭，不可关闭则等待人工滑动。"""
        page = page or self.page
        t0 = time.time()
        if not self.has_img_code(page):
            return True

        if not self._captcha_notified:
            quick_msg = '检测到 TikTok 验证码，请在浏览器中完成验证（不可关闭时需手动滑动）。'
            print('=' * 50)
            print(quick_msg)
            print('=' * 50)
            if self.on_captcha:
                try:
                    self.on_captcha(quick_msg)
                except Exception as e:
                    print(f'弹窗提醒失败: {e}')
            self._captcha_notified = True

        # #region agent log
        _agent_dbg("G", "TikTokSellerLogin.handle_captcha", "captcha detected", {
            "detect_ms": int((time.time() - t0) * 1000), "url": page.url,
        })
        # #endregion

        closable = self.can_close_captcha(page)
        if closable:
            print('验证码弹窗可关闭，尝试点击关闭...')
            if self.try_close_captcha(page):
                print('验证码已关闭，继续执行')
                self._captcha_notified = False
                return True
            print('验证码关闭失败，请继续手动完成滑块验证')
        else:
            print('验证码不可关闭，请手动完成滑块验证')

        return self._wait_manual_slide(page, timeout=timeout)

    def _find_captcha_close_button(self, page):
        container_selectors = (
            'css:#captcha_container',
            'css:.captcha_verify_container',
            'css:.captcha-disable-scroll',
        )
        close_selectors = (
            '[aria-label="Close"]',
            '[aria-label="close"]',
            '[aria-label="关闭"]',
            '[data-testid="close"]',
            '.secsdk-captcha-close',
            '.secsdk-captcha-close-btn',
            '.captcha_verify_close',
            '.captcha-close',
            '[class*="close-btn"]',
            '[class*="closeBtn"]',
            '[class*="CloseIcon"]',
            '[class*="close-icon"]',
            '[class*="close_icon"]',
        )

        for container in container_selectors:
            for close_sel in close_selectors:
                try:
                    btn = page.ele(f'{container} {close_sel}', timeout=0.3)
                    if btn:
                        return btn
                except Exception:
                    pass

        xpath_selectors = (
            'x://div[@id="captcha_container"]//*[@aria-label="Close" or @aria-label="close" or @aria-label="关闭"]',
            'x://div[contains(@class,"captcha_verify_container")]//*[@aria-label="Close" or @aria-label="close" or @aria-label="关闭"]',
            'x://div[@id="captcha_container"]//*[contains(@class,"close") and not(contains(@class,"closed"))]',
            'x://div[contains(@class,"captcha_verify_container")]//*[contains(@class,"close") and not(contains(@class,"closed"))]',
        )
        for selector in xpath_selectors:
            try:
                btn = page.ele(selector, timeout=0.3)
                if btn:
                    return btn
            except Exception:
                pass

        try:
            clicked = page.run_js('''
                const root = document.querySelector('#captcha_container')
                    || document.querySelector('.captcha_verify_container')
                    || document.querySelector('.captcha-disable-scroll');
                if (!root) return false;

                const isVisible = (el) => {
                    const rect = el.getBoundingClientRect();
                    return rect.width > 0 && rect.height > 0;
                };

                const tryClick = (el) => {
                    if (!el || !isVisible(el)) return false;
                    el.click();
                    return true;
                };

                for (const el of root.querySelectorAll('[aria-label]')) {
                    const label = (el.getAttribute('aria-label') || '').toLowerCase();
                    if (label.includes('close') || label.includes('关闭')) {
                        if (tryClick(el)) return true;
                    }
                }

                for (const el of root.querySelectorAll('[class*="close" i], [class*="Close"]')) {
                    if (tryClick(el)) return true;
                }

                for (const svg of root.querySelectorAll('svg')) {
                    const parent = svg.closest('button, [role="button"], div, span');
                    if (parent && isVisible(parent) && parent.getBoundingClientRect().width <= 80) {
                        if (tryClick(parent)) return true;
                    }
                }
                return false;
            ''')
            if clicked:
                return 'js_clicked'
        except Exception:
            pass

        return None

    def can_close_captcha(self, page=None):
        """判断当前验证码弹窗是否存在可点击的关闭按钮。"""
        page = page or self.page
        if not self.has_img_code(page):
            return False
        return self._find_captcha_close_button(page) is not None

    def try_close_captcha(self, page=None):
        """存在关闭按钮时点击关闭，关闭成功返回 True。"""
        page = page or self.page
        if not self.has_img_code(page):
            return True

        close_btn = self._find_captcha_close_button(page)
        if not close_btn:
            return False

        if close_btn != 'js_clicked':
            try:
                close_btn.click()
            except Exception:
                try:
                    close_btn.click(by_js=True)
                except Exception:
                    return False

        time.sleep(1.5)
        return not self.has_img_code(page)

    def _notify_captcha(self, closable=False, close_failed=False):
        if close_failed:
            msg = (
                '检测到 TikTok 验证码可关闭但自动关闭失败。\n'
                '请在浏览器中手动完成滑块拖动，完成后脚本将自动继续。'
            )
        elif closable:
            msg = (
                '检测到 TikTok 验证码，正在尝试关闭...\n'
                '若关闭失败请手动完成滑块验证。'
            )
        else:
            msg = (
                '检测到 TikTok 验证码不可关闭，必须人工完成滑块拖动。\n'
                '请在浏览器中完成验证，完成后脚本将自动继续，请勿关闭浏览器。'
            )

        if not self._captcha_notified or close_failed or not closable:
            print('=' * 50)
            print(msg)
            print('=' * 50)
            if self.on_captcha and (not closable or close_failed):
                try:
                    self.on_captcha(msg)
                except Exception as e:
                    print(f'弹窗提醒失败: {e}')
            if not closable or close_failed:
                self._captcha_notified = True

    def _wait_manual_slide(self, page, timeout=300):
        """等待人工完成不可关闭的滑块验证码。"""
        print(f'等待人工完成滑块验证（最长 {timeout} 秒）...')
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not self.has_img_code(page):
                print('滑块验证已通过，继续执行')
                self._captcha_notified = False
                return True
            time.sleep(1)

        raise RuntimeError('等待人工完成滑块验证超时')

    def wait_manual_captcha(self, page=None, timeout=300):
        """兼容入口：统一走 handle_captcha。"""
        return self.handle_captcha(page=page, timeout=timeout)

    def _wait_logged_in(self, page, timeout=300, captcha_timeout=300):
        deadline = time.time() + timeout
        while time.time() < deadline:
            page = self._pick_seller_tab(page)
            if self.has_img_code(page):
                self.handle_captcha(page, timeout=captcha_timeout)
            if self.is_logged_in(page):
                print('TikTok 登录成功，已进入卖家后台')
                self.page = page
                return True
            time.sleep(2)
        raise RuntimeError(f'等待 TikTok 登录完成超时，当前 URL: {page.url}')

    def has_img_code(self, page=None):
        page = page or self.page
        try:
            if page.ele('css:#captcha_container', timeout=0.2):
                return True
            if page.ele('css:.captcha_verify_container', timeout=0.2):
                return True
            if page.ele('x://*[contains(text(),"Verify to continue")]', timeout=0.2):
                return True
            if page.ele('x://*[contains(text(),"Drag the puzzle piece")]', timeout=0.2):
                return True
            if page.ele('x://*[contains(text(),"Drag the slider")]', timeout=0.2):
                return True
        except Exception:
            pass
        return False
