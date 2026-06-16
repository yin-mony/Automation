(function() {
    function parseCookie() {
        return document.cookie.split(';')
            .filter(function(item) {
                return !!item.trim();
            })
            .reduce(function(obj, item) {
                var pairs = item.split('=');
                var key = pairs[0] && pairs[0].trim();
                var val = pairs[1] && pairs[1].trim();
                obj[key] = val;
                return obj;
            }, {});
    }

    function setCookie(name, value) {
        try {
            if (window.sessionStorage) {
                if (!window.sessionStorage.getItem(name)) {
                    window.sessionStorage.setItem(name, value)
                }
            }
            if (window.localStorage) {
                if (!window.localStorage.getItem(name)) {
                    window.localStorage.setItem(name, value)
                }
            }
            if (!!!parseCookie()[name]) {
                document.cookie = name + "=" + value + ";";
            }
        } catch (e) { }
    }

    function initCaptcha() {
        // 设置 referrer cookie
        setCookie("__ac_referer", document.referrer || "__ac_blank");
        
        // 获取配置数据
        const configElement = document.getElementById('captcha-config');
        const config = JSON.parse(configElement.textContent);
        
        const options = {
            commonOptions: {
                aid: config.aid,
                iid: '0',
                did: '0',
            },
            captchaOptions: {
                hideCloseBtn: true,
                showMode: 'mask',
                successCb: function(){
                    setTimeout(() => {
                        location.reload();
                    }, 2000)
                },
            },
        };

        window.OECCaptcha.init(options);
        window.OECCaptcha.render({
            verify_data: config.verify_data
        });
    }

    // 确保 DOM 加载完成后初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initCaptcha);
    } else {
        initCaptcha();
    }
})();
