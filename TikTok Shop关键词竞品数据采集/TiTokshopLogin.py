import base64
import random
import time
from pathlib import Path

import cv2
import numpy as np
import requests
from DrissionPage import ChromiumPage
from PIL import Image


class TikTokPage:

    def __init__(self, page):
        self.page = page
        self.img_dir = Path(__file__).resolve().parent / 'img_code'
        self.img_dir.mkdir(exist_ok=True)
        self.piece_path = self.img_dir / 'piece.png'
        self.bg_path = self.img_dir / 'bg.jpg'

    

    def main(self):
        self.page.get('https://shop.tiktok.com/us')
        self.img_code()

    # 运行
    def run(self):
        self.main()
        time.sleep(3)

    # 滑动验证码（检测并处理，无验证码则直接返回）
    def img_code(self, page=None, timeout=180):
        page = page or self.page
        if not self.has_img_code(page):
            return True
        print('检测到滑动验证码，开始自动识别...')
        if self._solve_img_code(page):
            return True
        return self._wait_manual_img_code(page, timeout)

    # 自动识别验证码
    def _solve_img_code(self, page, max_attempts=5):
        for attempt in range(1, max_attempts + 1):
            if not self.has_img_code(page):
                print('验证码已通过，继续执行')
                return True

            print(f'自动识别验证码，第 {attempt}/{max_attempts} 次...')
            self.save_img_code(page)
            distance = self.img_code_df()
            if distance is None:
                print('验证码图片匹配失败，2 秒后重试...')
                time.sleep(2)
                continue

            move_x = int(float(distance) * 0.58)
            print(f'图片匹配距离: {distance}px，滑块拖动: {move_x}px')

            slider = page.ele('x://div[@class="sc-hMqMXs eUBOsN"]', timeout=3)
            if not slider:
                print('未找到滑块元素，2 秒后重试...')
                time.sleep(2)
                continue

            slider.scroll.to_see(center=True)
            time.sleep(0.2)
            vx, vy = slider.rect.viewport_midpoint
            print(f'滑块视口坐标: ({vx}, {vy})，开始拖动...')

            tracks = self.get_tracks(move_x)
            ac = page.actions
            ac.move_to(slider, duration=random.uniform(0.3, 0.5))
            time.sleep(0.2)
            ac.hold()
            for track in tracks:
                ac.move(
                    offset_x=track,
                    offset_y=random.randint(-2, 2),
                    duration=random.uniform(0.01, 0.03),
                )
                if random.random() < 0.1:
                    time.sleep(random.uniform(0.02, 0.1))
            ac.release()
            time.sleep(2.5)

            if not self.has_img_code(page):
                print('自动滑动验证码成功')
                return True

        print(f'自动识别验证码 {max_attempts} 次均未通过')
        return False

    # 等待人工完成验证码
    def _wait_manual_img_code(self, page, timeout=180):
        # print('自动识别失败，请手动完成滑动验证码...')
        start_time = time.time()
        while time.time() - start_time < timeout:
            if not self.has_img_code(page):
                print('人工验证码已通过，继续执行')
                return True
            time.sleep(2)
        raise TimeoutError('等待人工完成验证码超时')

    # 保存验证码图片（从页面下载拼图块和背景图）
    def save_img_code(self, page=None):
        page = page or self.page
        imgs = page.eles('css:#captcha_container img')
        if not imgs:
            imgs = page.eles('css:.captcha_verify_container img')

        for img in imgs:
            src = img.attr('src') or ''
            if not src:
                continue

            if src.startswith('data:'):
                data = base64.b64decode(src.split('base64,', 1)[1])
            else:
                session = requests.Session()
                for c in page.cookies():
                    session.cookies.set(c['name'], c['value'])
                data = session.get(src, timeout=15).content

            if data[:8] == b'\x89PNG\r\n\x1a\n':
                self.piece_path.write_bytes(data)
            elif data[:2] == b'\xff\xd8':
                self.bg_path.write_bytes(data)

    # 检查页面是否存在验证码（页面上还有没有验证码）
    def has_img_code(self, page=None):
        page = page or self.page
        title = ''
        html = ''
        try:
            title = page.title.lower()
        except Exception:
            pass
        try:
            html = page.html.lower()
        except Exception:
            pass
        keywords = (
            'security check',
            'verify to continue',
            'drag the puzzle piece',
            'drag the puzzle piece into place',
        )
        return any(k in title or k in html for k in keywords)

    # 验证码图片匹配（计算拼图块在背景图中的位置）
    def img_code_df(self):
        if not self.piece_path.exists() or not self.bg_path.exists():
            print(f'验证码图片不存在: {self.piece_path} / {self.bg_path}')
            return None

        piece = np.array(Image.open(self.piece_path).convert('RGBA'))
        bg = np.array(Image.open(self.bg_path).convert('RGB'))

        template = piece[:, :, :3]
        alpha = piece[:, :, 3]
        rgb = template.astype(np.int16)

        white = (rgb[:, :, 0] > 235) & (rgb[:, :, 1] > 235) & (rgb[:, :, 2] > 235)
        dark = rgb.mean(axis=2) < 18
        mask = ((alpha > 200) & (~white) & (~dark)).astype(np.uint8) * 255

        result = cv2.matchTemplate(bg, template, cv2.TM_CCOEFF_NORMED, mask=mask)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        print(f'匹配置信度: {max_val:.3f}')
        return max_loc[0]

    # 生成滑动轨迹（模拟滑动过程）
    @staticmethod
    def get_tracks(distance):
        value = round(random.uniform(0.55, 0.75), 2)
        v, t, sum1 = 0, 0.3, 0
        plus = []
        mid = distance * value
        while sum1 < distance:
            if sum1 < mid:
                a = round(random.uniform(2.5, 3.5), 1)
            else:
                a = -round(random.uniform(2.0, 3.0), 1)
            s = v * t + 0.5 * a * (t ** 2)
            v = v + a * t
            sum1 += s
            plus.append(round(s))
        return plus


if __name__ == '__main__':
    page = ChromiumPage()
    login_config = TikTokPage(page=page)
    login_config.run()
