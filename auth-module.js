// 授权验证模块
var authModule = (function () {
    // 授权验证API
    const AUTH_API_URL = 'https://aiwechat-vercel-virid.vercel.app/api/check_timeout';

    // 本地存储键名
    const STORAGE_KEYS = {
        AUTH_CODE: 'zhixin_word_auth_code',
        AUTH_EXPIRY: 'zhixin_word_auth_expiry',
        LAST_VERIFIED: 'zhixin_word_last_verified'
    };

    // 授权状态枚举
    const AUTH_STATUS = {
        VALID: 'valid',
        EXPIRED: 'expired',
        INVALID: 'invalid',
        UNKNOWN: 'unknown'
    };

    // 当前授权状态
    let currentAuthStatus = AUTH_STATUS.UNKNOWN;
    let currentExpiryDate = null;

    // 初始化授权模块
    function init() {
        console.log("初始化授权模块...");

        // 创建并添加授权按钮到页脚
        createAuthButton();

        // 创建并添加授权面板
        createAuthPanel();

        // 添加事件处理程序
        attachEventHandlers();

        // 检查授权状态
        checkAuthOnStartup();

        // 延迟添加功能拦截，确保其他脚本先初始化
        setTimeout(function () {
            setupFeatureRestrictions();
        }, 1000);
    }

    // 创建授权按钮
    function createAuthButton() {
        // 查找页脚文本元素
        const footerText = document.querySelector('.footer .footer-text');
        if (footerText) {
            // 创建授权按钮容器
            const authBtnContainer = document.createElement('div');
            authBtnContainer.style.display = 'inline-block';
            authBtnContainer.style.marginLeft = '10px';
            authBtnContainer.style.verticalAlign = 'middle';

            // 创建授权按钮
            const authBtn = document.createElement('button');
            authBtn.id = 'auth-button';
            authBtn.className = 'ms-Button ms-Button--icon';
            authBtn.title = '授权管理';
            authBtn.style.background = 'transparent';
            authBtn.style.border = 'none';
            authBtn.style.padding = '0';
            authBtn.style.margin = '0';
            authBtn.style.color = 'white';
            authBtn.style.cursor = 'pointer';
            authBtn.style.opacity = '0.8';
            authBtn.style.transition = 'opacity 0.2s';
            authBtn.innerHTML = '<i class="ms-Icon ms-Icon--Permissions" style="font-size: 14px;"></i>';

            // 添加悬停效果
            authBtn.onmouseover = function () { this.style.opacity = '1'; };
            authBtn.onmouseout = function () { this.style.opacity = '0.8'; };

            // 添加到容器
            authBtnContainer.appendChild(authBtn);

            // 插入到页脚文本后面
            footerText.appendChild(authBtnContainer);
        }
    }

    // 创建授权面板
    function createAuthPanel() {
        // 检查是否已存在
        if (document.getElementById('auth-panel')) {
            return;
        }

        // 创建面板元素
        const authPanel = document.createElement('div');
        authPanel.id = 'auth-panel';
        authPanel.className = 'ms-Dialog';
        authPanel.style.display = 'none';
        authPanel.style.position = 'fixed';
        authPanel.style.top = '50%';
        authPanel.style.left = '50%';
        authPanel.style.transform = 'translate(-50%, -50%)';
        authPanel.style.background = 'white';
        authPanel.style.zIndex = '1001';
        authPanel.style.padding = '20px';
        authPanel.style.boxShadow = '0 0 10px rgba(0,0,0,0.2)';
        authPanel.style.minWidth = '300px';
        authPanel.style.maxWidth = '90%';
        authPanel.style.borderRadius = '2px';

        // 构建面板HTML
        authPanel.innerHTML = `
            <div class="ms-Dialog-title" style="font-weight:600; font-size:18px; margin-bottom:16px;">产品授权验证</div>
            <div class="ms-Dialog-content">
                <div id="auth-status-section" style="margin-bottom:16px; padding:10px; border-radius:3px;">
                    <!-- 授权状态将动态填充 -->
                </div>
                <div class="ms-TextField">
                    <label class="ms-Label">授权码</label>
                    <input id="auth-code-input" class="ms-TextField-field" type="text" placeholder="请输入授权码">
                </div>
                <div id="auth-error" class="ms-MessageBar ms-MessageBar--error" style="display:none; margin-top:12px;">
                    <div class="ms-MessageBar-content">
                        <div class="ms-MessageBar-icon">
                            <i class="ms-Icon ms-Icon--Error"></i>
                        </div>
                        <div class="ms-MessageBar-text" id="auth-error-text"></div>
                    </div>
                </div>
            </div>
            <div class="ms-Dialog-actions" style="display:flex; justify-content:flex-end; margin-top:20px;">
                <button id="verify-auth-btn" class="ms-Button ms-Button--primary">
                    <span class="ms-Button-label">验证授权</span>
                </button>
                <button id="close-auth-panel-btn" class="ms-Button" style="margin-left:8px;">
                    <span class="ms-Button-label">关闭</span>
                </button>
            </div>
        `;

        // 添加到文档
        document.body.appendChild(authPanel);
    }

    // 添加事件处理程序
    function attachEventHandlers() {
        // 授权按钮点击事件
        document.addEventListener('click', function (e) {
            if (e.target.id === 'auth-button' ||
                (e.target.parentElement && e.target.parentElement.id === 'auth-button')) {
                updateAuthStatusDisplay();
                showAuthPanel();
            }
        });

        // 验证按钮点击事件
        document.addEventListener('click', function (e) {
            if (e.target.id === 'verify-auth-btn' ||
                (e.target.parentElement && e.target.parentElement.id === 'verify-auth-btn')) {
                handleVerifyAuth();
            }
        });

        // 关闭按钮点击事件
        document.addEventListener('click', function (e) {
            if (e.target.id === 'close-auth-panel-btn' ||
                (e.target.parentElement && e.target.parentElement.id === 'close-auth-panel-btn')) {
                hideAuthPanel();
            }
        });

        // 点击遮罩层关闭面板
        document.addEventListener('click', function (e) {
            if (e.target.id === 'auth-panel-mask') {
                hideAuthPanel();
            }
        });
    }

    // 设置功能限制
    function setupFeatureRestrictions() {
        // 需要授权的功能按钮ID列表
        const restrictedFeatures = [
            'improve-clarity',
            'improve-conciseness',
            'improve-tone',
            'plagiarism-check',
            'readability-score',
            'send-message'
        ];

        // 拦截点击事件
        document.addEventListener('click', function (event) {
            // 查找点击的按钮或其父元素
            let targetElement = event.target;

            // 向上查找四层，检查是否是按钮
            for (let i = 0; i < 4; i++) {
                if (!targetElement) break;

                // 如果是按钮且在限制列表中
                if (targetElement.id && restrictedFeatures.includes(targetElement.id)) {
                    // 检查授权状态
                    if (currentAuthStatus !== AUTH_STATUS.VALID) {
                        // 授权无效，阻止默认行为
                        event.preventDefault();
                        event.stopPropagation();

                        // 显示授权面板
                        showNotification("需要授权", `功能"${targetElement.innerText || targetElement.id}"需要有效授权`);
                        showAuthPanel();
                        return false;
                    }
                    // 授权有效，允许继续执行
                    break;
                }

                // 向上查找父元素
                targetElement = targetElement.parentElement;
            }
        }, true); // 使用捕获阶段

        console.log("功能限制设置完成");
    }

    // 更新授权状态显示
    function updateAuthStatusDisplay() {
        var statusSection = document.getElementById('auth-status-section');
        var currentDate = new Date().toLocaleDateString();
        var statusHtml = '';

        switch (currentAuthStatus) {
            case AUTH_STATUS.VALID:
                var expiryDateStr = currentExpiryDate ? formatDate(currentExpiryDate) : '未知';
                statusHtml = `
                    <div style="background-color:#dff6dd; padding:10px; border-left:4px solid #107c10;">
                        <div style="font-weight:600; margin-bottom:5px;">授权有效</div>
                        <div>截止日期: ${expiryDateStr}</div>
                        <div style="font-size:12px; margin-top:5px;">今日日期: ${currentDate}</div>
                    </div>`;
                break;

            case AUTH_STATUS.EXPIRED:
                statusHtml = `
                    <div style="background-color:#fed9cc; padding:10px; border-left:4px solid #d83b01;">
                        <div style="font-weight:600; margin-bottom:5px;">授权已过期</div>
                        <div>请输入新的授权码继续使用</div>
                        <div style="font-size:12px; margin-top:5px;">今日日期: ${currentDate}</div>
                    </div>`;
                break;

            case AUTH_STATUS.INVALID:
                statusHtml = `
                    <div style="background-color:#fed9cc; padding:10px; border-left:4px solid #d83b01;">
                        <div style="font-weight:600; margin-bottom:5px;">授权无效</div>
                        <div>请输入有效的授权码</div>
                        <div style="font-size:12px; margin-top:5px;">今日日期: ${currentDate}</div>
                    </div>`;
                break;

            default:
                statusHtml = `
                    <div style="background-color:#f0f0f0; padding:10px; border-left:4px solid #666666;">
                        <div style="font-weight:600; margin-bottom:5px;">授权状态未知</div>
                        <div>请输入授权码进行验证</div>
                        <div style="font-size:12px; margin-top:5px;">今日日期: ${currentDate}</div>
                    </div>`;
                break;
        }

        statusSection.innerHTML = statusHtml;
    }

    // 显示授权面板
    function showAuthPanel() {
        var authPanel = document.getElementById('auth-panel');
        if (authPanel) {
            hideAuthError();
            authPanel.style.display = 'block';

            // 显示背景遮罩
            var mask = document.createElement('div');
            mask.id = 'auth-panel-mask';
            mask.style.position = 'fixed';
            mask.style.top = '0';
            mask.style.left = '0';
            mask.style.width = '100%';
            mask.style.height = '100%';
            mask.style.backgroundColor = 'rgba(0,0,0,0.3)';
            mask.style.zIndex = '1000';
            document.body.appendChild(mask);

            // 聚焦授权码输入框
            setTimeout(function () {
                var authCodeInput = document.getElementById('auth-code-input');
                if (authCodeInput) {
                    authCodeInput.focus();
                }
            }, 100);
        } else {
            console.error("找不到授权面板元素");
        }
    }

    // 隐藏授权面板
    function hideAuthPanel() {
        var authPanel = document.getElementById('auth-panel');
        if (authPanel) {
            authPanel.style.display = 'none';

            // 移除背景遮罩
            var mask = document.getElementById('auth-panel-mask');
            if (mask) {
                document.body.removeChild(mask);
            }
        }
    }

    // 显示授权错误
    function showAuthError(message) {
        var authError = document.getElementById('auth-error');
        var authErrorText = document.getElementById('auth-error-text');
        if (authError && authErrorText) {
            authErrorText.textContent = message;
            authError.style.display = 'block';
        }
    }

    // 隐藏授权错误
    function hideAuthError() {
        var authError = document.getElementById('auth-error');
        if (authError) {
            authError.style.display = 'none';
        }
    }

    // 验证授权按钮处理
    function handleVerifyAuth() {
        var authCodeInput = document.getElementById('auth-code-input');
        if (authCodeInput) {
            var authCode = authCodeInput.value.trim();
            if (!authCode) {
                showAuthError("请输入授权码");
                return;
            }

            // 显示验证中状态
            document.getElementById('verify-auth-btn').disabled = true;
            document.getElementById('verify-auth-btn').innerHTML = '<span class="ms-Button-label">验证中...</span>';

            verifyAuthCode(authCode)
                .then(result => {
                    // 恢复按钮状态
                    document.getElementById('verify-auth-btn').disabled = false;
                    document.getElementById('verify-auth-btn').innerHTML = '<span class="ms-Button-label">验证授权</span>';

                    if (result.status === AUTH_STATUS.VALID) {
                        saveAuthInfo(authCode, result.expiryDate);
                        updateAuthStatusDisplay();
                        hideAuthError();
                        setTimeout(() => {
                            hideAuthPanel();
                            showNotification("授权成功", "您的产品授权有效期至 " + formatDate(result.expiryDate));
                        }, 1000);
                    } else {
                        showAuthError(result.message || "授权码无效");
                    }
                })
                .catch(error => {
                    // 恢复按钮状态
                    document.getElementById('verify-auth-btn').disabled = false;
                    document.getElementById('verify-auth-btn').innerHTML = '<span class="ms-Button-label">验证授权</span>';

                    console.error("验证授权码时出错:", error);
                    showAuthError("验证过程中出错，请稍后重试");
                });
        }
    }

    // 启动时检查授权状态
    function checkAuthOnStartup() {
        console.log("启动时检查授权状态...");

        // 获取存储的授权信息
        const authCode = localStorage.getItem(STORAGE_KEYS.AUTH_CODE);
        const authExpiry = localStorage.getItem(STORAGE_KEYS.AUTH_EXPIRY);
        const lastVerified = localStorage.getItem(STORAGE_KEYS.LAST_VERIFIED);

        // 如果没有授权信息，但不强制要求授权
        if (!authCode || !authExpiry) {
            console.log("未找到授权信息");
            currentAuthStatus = AUTH_STATUS.UNKNOWN;
            return;
        }

        // 检查本地存储的过期时间
        const expiryDate = new Date(authExpiry);
        const now = new Date();

        currentExpiryDate = expiryDate;

        if (expiryDate <= now) {
            console.log("授权已过期");
            currentAuthStatus = AUTH_STATUS.EXPIRED;
            return;
        }

        // 检查是否需要每日验证
        const needDailyVerification = shouldPerformDailyVerification(lastVerified);

        // 如果需要每日验证
        if (needDailyVerification) {
            console.log("需要进行每日验证");
            verifyAuthCode(authCode)
                .then(result => {
                    if (result.status === AUTH_STATUS.VALID) {
                        console.log("每日验证成功");
                        // 更新上次验证时间
                        updateLastVerifiedTime();
                        // 更新授权过期时间（如果服务器返回了新的过期时间）
                        if (result.expiryDate) {
                            saveAuthInfo(authCode, result.expiryDate);
                        }
                    } else {
                        console.log("每日验证失败:", result.message);
                        currentAuthStatus = result.status;
                    }
                })
                .catch(error => {
                    console.error("每日验证出错:", error);
                    // 如果验证失败但本地授权未过期，仍然允许使用
                    updateLastVerifiedTime();
                });
        } else {
            console.log("无需进行每日验证");
            currentAuthStatus = AUTH_STATUS.VALID;
        }
    }

    // 验证授权码
    function verifyAuthCode(authCode) {
        return new Promise((resolve, reject) => {
            console.log("开始验证授权码:", authCode);

            // 对于测试，模拟几个有效的授权码
            const testAuthCodes = {
                'TEST-123-456': { valid: true, expiryDate: new Date(new Date().setMonth(new Date().getMonth() + 1)) },
                'DEMO-789-000': { valid: true, expiryDate: new Date(new Date().setDate(new Date().getDate() + 7)) },
                'TRIAL-VERSION': { valid: true, expiryDate: new Date(new Date().setHours(new Date().getHours() + 24)) }
            };

            // 如果是测试授权码，直接返回结果（无需真实API调用）
            if (testAuthCodes[authCode]) {
                setTimeout(() => {
                    resolve({
                        status: AUTH_STATUS.VALID,
                        message: "授权验证成功",
                        expiryDate: testAuthCodes[authCode].expiryDate
                    });
                }, 1000);
                return;
            }

            // 真实API调用
            fetch(AUTH_API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ code: authCode })
            })
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`服务器返回错误: ${response.status}`);
                    }
                    return response.json();
                })
                .then(data => {
                    console.log("验证响应:", data);

                    // 解析服务器响应
                    if (data.valid === true) {
                        // 授权有效
                        let expiryDate = new Date();
                        if (data.expiry_date) {
                            expiryDate = new Date(data.expiry_date);
                        } else if (data.days_remaining) {
                            // 如果API只返回剩余天数，计算过期日期
                            expiryDate.setDate(expiryDate.getDate() + data.days_remaining);
                        }

                        resolve({
                            status: AUTH_STATUS.VALID,
                            message: "授权验证成功",
                            expiryDate: expiryDate
                        });
                    } else {
                        // 授权无效
                        resolve({
                            status: AUTH_STATUS.INVALID,
                            message: data.message || "授权码无效"
                        });
                    }
                })
                .catch(error => {
                    console.error("验证请求失败:", error);
                    reject(error);
                });
        });
    }

    // 判断是否需要执行每日验证
    function shouldPerformDailyVerification(lastVerified) {
        if (!lastVerified) {
            return true;
        }

        const lastVerifiedDate = new Date(lastVerified);
        const currentDate = new Date();

        // 检查是否是同一天
        return lastVerifiedDate.getFullYear() !== currentDate.getFullYear() ||
            lastVerifiedDate.getMonth() !== currentDate.getMonth() ||
            lastVerifiedDate.getDate() !== currentDate.getDate();
    }

    // 保存授权信息
    function saveAuthInfo(authCode, expiryDate) {
        localStorage.setItem(STORAGE_KEYS.AUTH_CODE, authCode);
        localStorage.setItem(STORAGE_KEYS.AUTH_EXPIRY, expiryDate.toISOString());
        updateLastVerifiedTime();
        currentAuthStatus = AUTH_STATUS.VALID;
        currentExpiryDate = expiryDate;

        console.log("授权信息已保存，有效期至:", expiryDate);
    }

    // 更新上次验证时间
    function updateLastVerifiedTime() {
        const now = new Date();
        localStorage.setItem(STORAGE_KEYS.LAST_VERIFIED, now.toISOString());
        console.log("更新上次验证时间:", now);
    }

    // 检查当前授权状态
    function checkCurrentAuthStatus() {
        return currentAuthStatus;
    }

    // 格式化日期
    function formatDate(date) {
        if (!(date instanceof Date)) {
            date = new Date(date);
        }
        return `${date.getFullYear()}-${padZero(date.getMonth() + 1)}-${padZero(date.getDate())}`;
    }

    // 数字补零
    function padZero(num) {
        return num < 10 ? `0${num}` : num;
    }

    // 显示通知
    function showNotification(title, message) {
        if (typeof window.showNotification === 'function') {
            window.showNotification(title, message);
        } else {
            alert(`${title}: ${message}`);
        }
    }

    // 公开的API
    return {
        init: init,
        showAuthPanel: showAuthPanel,
        verifyAuthCode: verifyAuthCode,
        checkAuthStatus: checkCurrentAuthStatus
    };
})();

// 当DOM加载完成后初始化
document.addEventListener('DOMContentLoaded', function () {
    // 在其他脚本之后初始化授权模块
    setTimeout(function () {
        authModule.init();
    }, 500);
});