// 聊天功能集成脚本 - 与主脚本分离，避免冲突
(function () {
    // 等待文档加载完成
    function initChatFeature() {
        console.log("正在初始化聊天功能...");
    
  
        // 保存的Prompt模板
        var chatModule = {
            promptTemplates: {
                "default": "请基于以下内容回答问题:\n\n{selection}\n\n",
                "academic": "请帮我润色以下学术内容，提高其专业性和流畅度，保持学术风格:\n\n{selection}",
                "creative": "请帮我优化以下创意写作内容，提高其吸引力和表现力:\n\n{selection}",
                "technical": "请帮我优化以下技术文档，使其更清晰、准确，并易于理解:\n\n{selection}"
            },

            // 初始化功能
            init: function () {
                // 检查OpenAIService是否可用
                if (typeof OpenAIService === 'undefined') {
                    console.error("错误: OpenAIService未定义! 请确保先加载openai-service.js");
                    this.showNotification("错误", "AI服务未正确加载");
                    return;
                }
                console.log("OpenAIService已检测到");
                // 绑定事件处理程序
                this.bindEvents();
                // 加载模板
                this.loadTemplates();
                // 确保Prompt显示正常
                this.ensurePromptVisibility();
                console.log("聊天功能初始化完成");
            },

            // 绑定界面事件
            bindEvents: function () {
                var self = this;
                console.log("绑定事件...");

                // 选择器变更事件
                var selector = document.getElementById("prompt-selector");
                if (selector) {
                    selector.addEventListener("change", function () {
                        self.loadPromptTemplate();
                    });
                }

                // 绑定全局处理程序
                window.handleSavePromptClick = function () {
                    console.log("保存按钮点击");
                    self.showSavePromptDialog();
                };

                window.handleEditPromptClick = function () {
                    console.log("编辑按钮点击");
                    self.editPromptTemplate();
                };

                window.handleDeletePromptClick = function () {
                    console.log("删除按钮点击");
                    self.deletePromptTemplate();
                };

                window.handleSaveConfirmClick = function () {
                    console.log("保存确认点击");
                    self.savePromptTemplate();
                };

                window.handleSaveCancelClick = function () {
                    console.log("保存取消点击");
                    self.hideSavePromptDialog();
                };

                window.handleEditConfirmClick = function () {
                    console.log("编辑确认点击");
                    self.updatePromptTemplate();
                };

                window.handleEditCancelClick = function () {
                    console.log("编辑取消点击");
                    self.hideEditPromptDialog();
                };

                window.handleDeleteYesClick = function () {
                    console.log("删除确认点击");
                    self.confirmDeleteTemplate();
                };

                window.handleDeleteNoClick = function () {
                    console.log("删除取消点击");
                    self.cancelDeleteTemplate();
                };

                // 发送按钮
                var sendButton = document.getElementById("send-message");
                if (sendButton) {
                    sendButton.addEventListener("click", function () {
                        self.sendMessage();
                    });
                }

                // 备用方式 - 直接绑定到元素上
                this.bindBackupEvents();

                console.log("事件绑定完成");
            },

            // 备用绑定方式 - 直接给按钮绑定事件
            bindBackupEvents: function () {
                var self = this;

                var saveBtn = document.getElementById("save-prompt");
                if (saveBtn) {
                    saveBtn.addEventListener("click", function () {
                        self.showSavePromptDialog();
                    });
                }

                var editBtn = document.getElementById("edit-prompt");
                if (editBtn) {
                    editBtn.addEventListener("click", function () {
                        self.editPromptTemplate();
                    });
                }

                var deleteBtn = document.getElementById("delete-prompt");
                if (deleteBtn) {
                    deleteBtn.addEventListener("click", function () {
                        self.deletePromptTemplate();
                    });
                }

                // 保存对话框按钮
                var saveConfirmBtn = document.getElementById("save-prompt-confirm");
                if (saveConfirmBtn) {
                    saveConfirmBtn.addEventListener("click", function () {
                        self.savePromptTemplate();
                    });
                }

                var saveCancelBtn = document.getElementById("save-prompt-cancel");
                if (saveCancelBtn) {
                    saveCancelBtn.addEventListener("click", function () {
                        self.hideSavePromptDialog();
                    });
                }

                // 编辑对话框按钮
                var editConfirmBtn = document.getElementById("edit-prompt-confirm");
                if (editConfirmBtn) {
                    editConfirmBtn.addEventListener("click", function () {
                        self.updatePromptTemplate();
                    });
                }

                var editCancelBtn = document.getElementById("edit-prompt-cancel");
                if (editCancelBtn) {
                    editCancelBtn.addEventListener("click", function () {
                        self.hideEditPromptDialog();
                    });
                }
            },

            // 加载模板
            loadTemplates: function () {
                console.log("加载模板...");
                // 设置默认模板到文本框
                var promptContentElem = document.getElementById("prompt-content");
                if (promptContentElem) {
                    promptContentElem.value = this.promptTemplates["default"];
                    promptContentElem.style.color = "#333333";
                    promptContentElem.style.backgroundColor = "#ffffff";
                }

                // 尝试从本地存储加载自定义模板
                try {
                    var savedTemplates = localStorage.getItem('customPromptTemplates');
                    if (savedTemplates) {
                        var selector = document.getElementById("prompt-selector");
                        var customTemplates = JSON.parse(savedTemplates);

                        // 添加到选择器
                        for (var name in customTemplates) {
                            if (customTemplates.hasOwnProperty(name)) {
                                var option = document.createElement("option");
                                option.value = name;
                                option.textContent = name;
                                selector.appendChild(option);
                            }
                        }
                    }
                } catch (e) {
                    console.error("加载自定义模板失败:", e);
                }
            },

            // 确保Prompt内容正确显示
            ensurePromptVisibility: function () {
                var promptContent = document.getElementById("prompt-content");

                if (promptContent) {
                    // 强制设置样式确保可见
                    promptContent.style.color = "#333333";
                    promptContent.style.backgroundColor = "#ffffff";
                    promptContent.style.fontFamily = "inherit";
                    promptContent.style.fontSize = "14px";

                    // 检查内容
                    if (!promptContent.value) {
                        var selector = document.getElementById("prompt-selector");
                        if (selector && selector.value) {
                            var template = this.promptTemplates[selector.value];
                            if (template) {
                                promptContent.value = template;
                                console.log("设置了默认Prompt内容");
                            }
                        }
                    }
                }
            },

            // 加载选中的Prompt模板
            loadPromptTemplate: function () {
                console.log("加载模板内容...");
                var selector = document.getElementById("prompt-selector");
                var selectedTemplate = selector.value;

                // 首先从预设模板中查找
                if (this.promptTemplates[selectedTemplate]) {
                    var promptContent = document.getElementById("prompt-content");
                    promptContent.value = this.promptTemplates[selectedTemplate];
                    promptContent.style.color = "#333333";
                    return;
                }

                // 如果不是预设模板，从自定义模板中查找
                try {
                    var savedTemplates = localStorage.getItem('customPromptTemplates');
                    if (savedTemplates) {
                        var customTemplates = JSON.parse(savedTemplates);
                        if (customTemplates[selectedTemplate]) {
                            var promptContent = document.getElementById("prompt-content");
                            promptContent.value = customTemplates[selectedTemplate];
                            promptContent.style.color = "#333333";
                        }
                    }
                } catch (e) {
                    console.error("加载自定义模板失败:", e);
                    this.showNotification("错误", "加载模板失败");
                }
            },

            // 显示保存Prompt对话框
            showSavePromptDialog: function () {
                console.log("显示保存对话框...");
                var dialog = document.getElementById("save-prompt-dialog");

                if (dialog) {
                    var nameInput = document.getElementById("prompt-name");
                    if (nameInput) nameInput.value = "";

                    dialog.style.display = "block";
                    dialog.style.position = "fixed";
                    dialog.style.zIndex = "1001";

                    // 添加遮罩
                    this.addDialogMask();
                    console.log("保存对话框已显示");
                } else {
                    console.error("未找到保存对话框!");
                    this.showNotification("错误", "未找到保存对话框");
                }
            },

            // 添加对话框遮罩
            addDialogMask: function () {
                // 移除旧遮罩（如果存在）
                this.removeDialogMask();

                // 添加新遮罩
                var mask = document.createElement("div");
                mask.className = "dialog-mask";
                mask.style.position = "fixed";
                mask.style.top = "0";
                mask.style.left = "0";
                mask.style.width = "100%";
                mask.style.height = "100%";
                mask.style.background = "rgba(0,0,0,0.3)";
                mask.style.zIndex = "1000";
                document.body.appendChild(mask);
            },

            // 移除对话框遮罩
            removeDialogMask: function () {
                var mask = document.querySelector(".dialog-mask");
                if (mask) {
                    document.body.removeChild(mask);
                }
            },

            // 隐藏保存Prompt对话框
            hideSavePromptDialog: function () {
                var dialog = document.getElementById("save-prompt-dialog");
                if (dialog) {
                    dialog.style.display = "none";
                    this.removeDialogMask();
                }
            },

            // 保存Prompt模板
            savePromptTemplate: function () {
                var name = document.getElementById("prompt-name").value.trim();
                var content = document.getElementById("prompt-content").value.trim();

                if (!name) {
                    this.showNotification("错误", "请输入模板名称");
                    return;
                }

                if (!content) {
                    this.showNotification("错误", "模板内容不能为空");
                    return;
                }

                try {
                    // 保存到本地存储
                    var customTemplates = {};
                    var savedTemplates = localStorage.getItem('customPromptTemplates');
                    if (savedTemplates) {
                        customTemplates = JSON.parse(savedTemplates);
                    }

                    // 添加新模板
                    customTemplates[name] = content;
                    localStorage.setItem('customPromptTemplates', JSON.stringify(customTemplates));

                    // 更新选择器
                    var selector = document.getElementById("prompt-selector");
                    var optionExists = false;

                    // 检查是否已存在此名称的选项
                    for (var i = 0; i < selector.options.length; i++) {
                        if (selector.options[i].value === name) {
                            optionExists = true;
                            break;
                        }
                    }

                    // 如果不存在，添加新选项
                    if (!optionExists) {
                        var option = document.createElement("option");
                        option.value = name;
                        option.textContent = name;
                        selector.appendChild(option);
                    }

                    // 选中新保存的模板
                    selector.value = name;

                    this.showNotification("成功", "已保存模板: " + name);
                    this.hideSavePromptDialog();
                } catch (e) {
                    console.error("保存模板失败:", e);
                    this.showNotification("错误", "保存模板失败");
                }
            },

            // 删除Prompt模板功能 - 不使用confirm
            deletePromptTemplate: function () {
                console.log("执行删除模板...");
                var selector = document.getElementById("prompt-selector");
                var selectedValue = selector.value;

                // 检查是否是预设模板（不允许删除）
                var defaultTemplates = ["default", "academic", "creative", "technical"];
                if (defaultTemplates.includes(selectedValue)) {
                    this.showNotification("提示", "预设模板不能删除");
                    return;
                }

                // 显示自定义确认UI
                this.tempSelectedTemplate = selectedValue; // 存储当前选择的模板

                // 显示确认对话框
                var confirmHtml = `
                    <div style="background-color:#fff; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.1); margin-bottom:20px; overflow:hidden;">
                        <div style="background-color:#f0f2f5; padding:12px 16px; font-size:16px; font-weight:600; color:#0078d4; border-bottom:1px solid #e6e9ed;">确认删除</div>
                        <div style="padding:16px;">
                            <div style="margin-bottom:16px;">您确定要删除模板 "${selectedValue}" 吗？</div>
                            <div style="display:flex; justify-content:flex-end; gap:10px;">
                                <button id="delete-no" class="ms-Button" onclick="window.handleDeleteNoClick()">
                                    <span class="ms-Button-label">取消</span>
                                </button>
                                <button id="delete-yes" class="ms-Button ms-Button--primary" onclick="window.handleDeleteYesClick()">
                                    <span class="ms-Button-label">确认删除</span>
                                </button>
                            </div>
                        </div>
                    </div>
                `;

                this.showResults(confirmHtml);
            },

            // 确认删除模板
            confirmDeleteTemplate: function () {
                var selectedValue = this.tempSelectedTemplate;
                if (!selectedValue) {
                    this.showNotification("错误", "没有要删除的模板");
                    return;
                }

                var selector = document.getElementById("prompt-selector");

                try {
                    // 从下拉菜单中移除
                    for (var i = 0; i < selector.options.length; i++) {
                        if (selector.options[i].value === selectedValue) {
                            selector.remove(i);
                            break;
                        }
                    }

                    // 从存储中删除
                    var savedTemplates = localStorage.getItem('customPromptTemplates');
                    if (savedTemplates) {
                        var customTemplates = JSON.parse(savedTemplates);
                        if (customTemplates[selectedValue]) {
                            delete customTemplates[selectedValue];
                            localStorage.setItem('customPromptTemplates', JSON.stringify(customTemplates));
                        }
                    }

                    // 选择默认模板
                    selector.value = "default";
                    this.loadPromptTemplate();

                    // 清空临时存储
                    this.tempSelectedTemplate = null;

                    // 显示成功通知
                    this.showNotification("成功", "已删除模板: " + selectedValue);

                    // 清空结果区域
                    this.clearResults();
                } catch (e) {
                    console.error("删除模板失败:", e);
                    this.showNotification("错误", "删除模板失败");
                }
            },

            // 取消删除模板
            cancelDeleteTemplate: function () {
                // 清空临时存储
                this.tempSelectedTemplate = null;

                // 清空结果区域
                this.clearResults();
            },

            // 显示编辑Prompt对话框
            editPromptTemplate: function () {
                console.log("显示编辑对话框...");
                var selector = document.getElementById("prompt-selector");
                var selectedValue = selector.value;

                // 检查是否是预设模板
                var defaultTemplates = ["default", "academic", "creative", "technical"];
                if (defaultTemplates.includes(selectedValue)) {
                    this.showNotification("提示", "预设模板不能编辑，请另存为新模板");
                    return;
                }

                // 获取当前模板内容
                var currentContent = document.getElementById("prompt-content").value;

                // 查找编辑对话框
                var dialog = document.getElementById("edit-prompt-dialog");
                if (!dialog) {
                    console.error("未找到编辑对话框!");

                    // 如果找不到编辑对话框，尝试创建一个
                    this.createEditDialog();
                    dialog = document.getElementById("edit-prompt-dialog");

                    if (!dialog) {
                        this.showNotification("错误", "无法创建编辑对话框");
                        return;
                    }
                }

                // 设置到编辑对话框
                var nameField = document.getElementById("edit-prompt-name");
                var contentField = document.getElementById("edit-prompt-content");

                if (nameField && contentField) {
                    nameField.value = selectedValue;
                    contentField.value = currentContent;
                    contentField.style.color = "#333333";
                    contentField.style.backgroundColor = "#ffffff";

                    // 显示编辑对话框
                    dialog.style.display = "block";
                    dialog.style.position = "fixed";
                    dialog.style.zIndex = "1001";
                    dialog.style.top = "50%";
                    dialog.style.left = "50%";
                    dialog.style.transform = "translate(-50%, -50%)";

                    // 添加遮罩
                    this.addDialogMask();
                    console.log("编辑对话框已显示");
                } else {
                    console.error("未找到编辑对话框的必要字段!");
                    this.showNotification("错误", "无法编辑模板");
                }
            },

            // 创建编辑对话框（如果不存在）
            createEditDialog: function () {
                if (!document.getElementById("edit-prompt-dialog")) {
                    var dialogHtml = `
                        <div id="edit-prompt-dialog" class="ms-Dialog" style="display:none; position:fixed; top:50%; left:50%; transform:translate(-50%,-50%); background:white; z-index:1001; padding:20px; box-shadow:0 0 10px rgba(0,0,0,0.2); min-width:300px;">
                            <div class="ms-Dialog-title">编辑Prompt</div>
                            <div class="ms-Dialog-content">
                                <div class="ms-TextField">
                                    <label class="ms-Label">模板名称</label>
                                    <input id="edit-prompt-name" class="ms-TextField-field" type="text" placeholder="模板名称">
                                </div>
                                <div class="ms-TextField" style="margin-top:16px;">
                                    <label class="ms-Label">模板内容</label>
                                    <textarea id="edit-prompt-content" class="ms-TextField-field" 
                                            placeholder="输入自定义Prompt，可以使用{selection}表示选中的文本"
                                            rows="5" style="width:100%; color:#333333;"></textarea>
                                </div>
                            </div>
                            <div class="ms-Dialog-actions">
                                <button id="edit-prompt-confirm" class="ms-Button ms-Button--primary" onclick="window.handleEditConfirmClick()">
                                    <span class="ms-Button-label">保存更改</span>
                                </button>
                                <button id="edit-prompt-cancel" class="ms-Button" onclick="window.handleEditCancelClick()">
                                    <span class="ms-Button-label">取消</span>
                                </button>
                            </div>
                        </div>
                    `;

                    document.body.insertAdjacentHTML('beforeend', dialogHtml);
                    console.log("创建了编辑对话框");

                    // 绑定事件
                    var self = this;
                    document.getElementById("edit-prompt-confirm").addEventListener("click", function () {
                        self.updatePromptTemplate();
                    });

                    document.getElementById("edit-prompt-cancel").addEventListener("click", function () {
                        self.hideEditPromptDialog();
                    });
                }
            },

            // 隐藏编辑对话框
            hideEditPromptDialog: function () {
                var dialog = document.getElementById("edit-prompt-dialog");
                if (dialog) {
                    dialog.style.display = "none";
                    this.removeDialogMask();
                }
            },

            // 更新Prompt模板
            updatePromptTemplate: function () {
                var oldName = document.getElementById("prompt-selector").value;
                var newName = document.getElementById("edit-prompt-name").value.trim();
                var content = document.getElementById("edit-prompt-content").value.trim();

                if (!newName) {
                    this.showNotification("错误", "请输入模板名称");
                    return;
                }

                if (!content) {
                    this.showNotification("错误", "模板内容不能为空");
                    return;
                }

                try {
                    // 获取现有模板
                    var customTemplates = {};
                    var savedTemplates = localStorage.getItem('customPromptTemplates');
                    if (savedTemplates) {
                        customTemplates = JSON.parse(savedTemplates);
                    }

                    // 删除旧模板
                    if (oldName !== newName && customTemplates[oldName]) {
                        delete customTemplates[oldName];
                    }

                    // 添加新模板
                    customTemplates[newName] = content;
                    localStorage.setItem('customPromptTemplates', JSON.stringify(customTemplates));

                    // 更新选择器
                    var selector = document.getElementById("prompt-selector");

                    // 如果名称已更改，需要更新选择器
                    if (oldName !== newName) {
                        // 删除旧选项
                        for (var i = 0; i < selector.options.length; i++) {
                            if (selector.options[i].value === oldName) {
                                selector.remove(i);
                                break;
                            }
                        }

                        // 添加新选项
                        var option = document.createElement("option");
                        option.value = newName;
                        option.textContent = newName;
                        selector.appendChild(option);

                        // 选中新选项
                        selector.value = newName;
                    }

                    // 更新文本框
                    document.getElementById("prompt-content").value = content;

                    this.showNotification("成功", "已更新模板");
                    this.hideEditPromptDialog();
                } catch (e) {
                    console.error("更新模板失败:", e);
                    this.showNotification("错误", "更新模板失败");
                }
            },
            // 发送消息到AI - 替换为使用OpenAIService
            sendMessage: function () {
                var self = this;
                console.log("发送消息到AI...");

                // 检查OpenAIService是否可用
                if (typeof OpenAIService === 'undefined') {
                    console.error("错误: OpenAIService未定义!");
                    this.showNotification("错误", "AI服务未正确加载");
                    return;
                }

                // 获取当前日期时间
                var currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
                console.log(`[${currentTime}] 准备发送请求到OpenAI服务`);

                Word.run(function (context) {
                    // 获取用户选中的文本范围
                    var range = context.document.getSelection();
                    range.load('text');

                    return context.sync().then(function () {
                        var selectedText = range.text;
                        var promptTemplate = document.getElementById("prompt-content").value;

                        if (!promptTemplate) {
                            self.showNotification("错误", "请先选择或输入Prompt模板");
                            return;
                        }

                        // 检查是否需要选中文本
                        if (!selectedText && promptTemplate.includes("{selection}")) {
                            self.showNotification("提示", "请先在文档中选择文本");
                            return;
                        }

                        // 构建最终提示
                        var finalPrompt = promptTemplate.replace("{selection}", selectedText || "");
                        console.log(`[${currentTime}] 生成的提示，长度: ${finalPrompt.length}`);

                        // 显示加载状态到结果区域
                        self.updateResultsPanel(`
                            <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; padding:20px;">
                                <div style="width:40px; height:40px; border:4px solid #f3f3f3; border-top:4px solid #0078d4; border-radius:50%; animation:spin 1s linear infinite;"></div>
                                <div style="margin-top:16px; text-align:center;">AI思考中，请稍候...</div>
                            </div>
                            <style>
                                @keyframes spin {
                                    0% { transform: rotate(0deg); }
                                    100% { transform: rotate(360deg); }
                                }
                            </style>`);

                        // 使用OpenAIService直接调用OpenAI
                        console.log(`[${currentTime}] 调用OpenAIService.sendRequest`);

                        // 直接调用OpenAI服务
                        OpenAIService.sendRequest(finalPrompt)
                            .then(function (aiResponseText) {
                                // 展示响应
                                console.log("接收到的AI回复:", aiResponseText.substring(0, 100) + "...");
                                // 创建美观的HTML显示
                                var formattedResponse = '<p>' + self.formatAIResponse(aiResponseText) + '</p>';

                                var html = `
                                <style>
                                    .ai-response-content {
                                        white-space: pre-wrap;
                                        line-height: 1.5;
                                        background-color: #f9f9f9;
                                        border-left: 4px solid #0078d4;
                                        padding: 16px;
                                        border-radius: 0 4px 4px 0;
                                        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                                    }
                                    .ai-response-content h1, .ai-response-content h2, .ai-response-content h3,
                                    .ai-response-content h4, .ai-response-content h5 {
                                        margin-top: 20px;
                                        margin-bottom: 10px;
                                        font-weight: 600;
                                        line-height: 1.25;
                                    }
                                    .ai-response-content h1 { font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }
                                    .ai-response-content h2 { font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }
                                    .ai-response-content h3 { font-size: 1.25em; }
                                    .ai-response-content h4 { font-size: 1em; }
                                    .ai-response-content h5 { font-size: 0.875em; }
                                    .ai-response-content hr { height: 1px; background-color: #e1e4e8; border: 0; margin: 16px 0; }
                                    .ai-response-content ul, .ai-response-content ol { padding-left: 20px; margin: 16px 0; }
                                    .ai-response-content li { margin: 4px 0; }
                                    .ai-response-content code.inline-code {
                                        font-family: SFMono-Regular, Consolas, Liberation Mono, Menlo, monospace;
                                        padding: 2px 4px;
                                        background-color: rgba(27, 31, 35, 0.05);
                                        border-radius: 3px;
                                        font-size: 85%;
                                    }
                                    .ai-response-content .code-block {
                                        margin: 16px 0;
                                    }
                                    .ai-response-content .code-block pre {
                                        background-color: #f6f8fa;
                                        border-radius: 3px;
                                        padding: 16px;
                                        overflow: auto;
                                        margin: 0;
                                    }
                                    .ai-response-content .code-block code {
                                        font-family: SFMono-Regular, Consolas, Liberation Mono, Menlo, monospace;
                                        font-size: 85%;
                                        color: #24292e;
                                        line-height: 1.45;
                                        word-break: normal;
                                        white-space: pre;
                                    }
                                    .ai-response-content p {
                                        margin-top: 0;
                                        margin-bottom: 16px;
                                    }
                                </style>
                                <div class="ai-response">
                                    <div class="ai-response-header">
                                        <!-- 头部内容保持不变 -->
                                    </div>
            
                                    <div class="ai-response-content">
                                        ${formattedResponse}
                                    </div>
                
                                    <div style="margin-top:20px; display:flex; justify-content:flex-end;">
                                        <button id="btn-insert-response" style="background-color:#0078d4; color:white; border:none; padding:6px 12px; border-radius:2px; cursor:pointer; margin-right:8px;">
                                            插入到文档
                                        </button>
                                        <button id="btn-copy-response" style="background-color:#f3f3f3; color:#333; border:none; padding:6px 12px; border-radius:2px; cursor:pointer;">
                                            复制
                                        </button>
                                    </div>
                                </div>
                                `;

                                // 调用更新函数显示结果
                                self.updateResultsPanel(html);

                                // 存储结果供后续使用
                                self.currentAIResponse = aiResponseText;

                                // 添加按钮事件监听器 - 关键是在DOM更新后绑定事件
                                setTimeout(function () {
                                    var insertButton = document.getElementById("btn-insert-response");
                                    var copyButton = document.getElementById("btn-copy-response");

                                    if (insertButton) {
                                        insertButton.addEventListener("click", function () {
                                            self.insertTextToDocument(aiResponseText);
                                        });
                                    }

                                    if (copyButton) {
                                        copyButton.addEventListener("click", function () {
                                            self.copyToClipboard(aiResponseText);
                                        });
                                    }
                                }, 100); // 短暂延时确保DOM已更新
                            })
                            .catch(function (error) {
                                console.error("AI响应错误:", error);
                                self.updateResultsPanel(`
                                <div style="background-color:#fdeceb; border-radius:4px; padding:16px; display:flex; align-items:flex-start;">
                                    <div style="font-size:24px; color:#a80000; margin-right:16px;">
                                        <i class="ms-Icon ms-Icon--ErrorBadge"></i>
                                    </div>
                                    <div>
                                        <div style="font-weight:600; margin-bottom:4px;">AI服务调用失败</div>
                                        <div style="font-size:14px; opacity:0.9;">${error.message || "未知错误"}</div>
                                    </div>
                                </div>
                                `);
                            });
                    })
                }).catch(function (error) {
                    console.error("Word.run错误:", error);
                    self.showNotification("错误", "无法获取选定文本");
                });
            },

            // 发送消息到AI
            //sendMessage: function () {
            //    var self = this;
            //    console.log("发送消息到AI...");

            //    // 获取API基址
            //    var apiBaseUrl;
            //    if (typeof window.API_BASE_URL !== 'undefined') {
            //        apiBaseUrl = window.API_BASE_URL;
            //    } else {
            //        // 尝试从页面中找到API调用来推断基址
            //        apiBaseUrl = "http://localhost:54112/api"; // 默认值
            //    }

            //    // 获取当前日期时间
            //    var currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
            //    console.log(`[${currentTime}] 准备发送请求，API基址: ${apiBaseUrl}`);
             
            //    Word.run(function (context)
            //    {
            //        // 获取用户选中的文本范围
            //        var range = context.document.getSelection();
            //        range.load('text');

            //        return context.sync().then(function () {
            //                      var selectedText = range.text;
            //                var promptTemplate = document.getElementById("prompt-content").value;

            //                if (!promptTemplate) {
            //                    self.showNotification("错误", "请先选择或输入Prompt模板");
            //                    return;
            //                }

            //                // 检查是否需要选中文本
            //                if (!selectedText && promptTemplate.includes("{selection}")) {
            //                    self.showNotification("提示", "请先在文档中选择文本");
            //                    return;
            //                }

            //                // 构建最终提示
            //                var finalPrompt = promptTemplate.replace("{selection}", selectedText || "");
            //            console.log(`[${currentTime}] 生成的提示，长度: ${finalPrompt.length}`);

            //                    // 显示加载状态到结果区域
            //                self.updateResultsPanel(`
            //        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; padding:20px;">
            //            <div style="width:40px; height:40px; border:4px solid #f3f3f3; border-top:4px solid #0078d4; border-radius:50%; animation:spin 1s linear infinite;"></div>
            //            <div style="margin-top:16px; text-align:center;">AI思考中，请稍候...</div>
            //        </div>
            //        <style>
            //            @keyframes spin {
            //                0% { transform: rotate(0deg); }
            //                100% { transform: rotate(360deg); }
            //            }
            //        </style>`);
            //                  // 调用AI服务
            //            console.log(`[${currentTime}] 调用AI服务: ${apiBaseUrl}/TextAnalysis/aIreply`);
            //                                        // 创建请求
            //            return  fetch(`${apiBaseUrl}/TextAnalysis/aIreply`, {
            //                    method: "POST",
            //                    headers: {
            //                        "Content-Type": "application/json"
            //                    },
            //                    body: JSON.stringify({
            //                        text: finalPrompt,
            //                        mode: "custom" // 使用自定义模式
            //                    })
            //                })
            //        })
            //            .then(response => { // 处理fetch返回的响应
            //                if (!response.ok) {
            //                    throw new Error(`API请求失败: ${response.status}`);
            //                }
            //                return response.json();
            //            })
            //            .then(result => { // 处理json解析的结果
            //                console.log("API返回数据:", result);
            //                // 使用aireplaylatedText作为响应文本（根据您后端的属性名）
            //                var aiResponseText = result.aireplaylatedText;

            //                if (!aiResponseText) {
            //                    console.error("找不到期望的属性，完整响应:", result);
            //                    throw new Error("API返回的响应中未找到文本内容");
            //                }

            //                // 展示响应
            //                console.log("接收到的AI回复:", aiResponseText.substring(0, 100) + "...");
            //                // 创建美观的HTML显示
            //                var currentDate = new Date().toLocaleString();
            //                //var formattedResponse = self.formatAIResponse(aiResponseText)
            //                var formattedResponse = '<p>' + self.formatAIResponse(aiResponseText) + '</p>';
            //                //var formattedResponse = aiResponseText
            //                //    .replace(/\n/g, '<br>')  // 将换行符转换为HTML换行
            //                //    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')  // 处理Markdown加粗
            //                //    .replace(/\*(.*?)\*/g, '<em>$1</em>');  // 处理Markdown斜体

            //                var html = `
            //            <style>
            //                .ai-response-content {
            //                    white-space: pre-wrap;
            //                    line-height: 1.5;
            //                    background-color: #f9f9f9;
            //                    border-left: 4px solid #0078d4;
            //                    padding: 16px;
            //                    border-radius: 0 4px 4px 0;
            //                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            //                }
            //                .ai-response-content h1, .ai-response-content h2, .ai-response-content h3,
            //                .ai-response-content h4, .ai-response-content h5 {
            //                    margin-top: 20px;
            //                    margin-bottom: 10px;
            //                    font-weight: 600;
            //                    line-height: 1.25;
            //                }
            //                .ai-response-content h1 { font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }
            //                .ai-response-content h2 { font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }
            //                .ai-response-content h3 { font-size: 1.25em; }
            //                .ai-response-content h4 { font-size: 1em; }
            //                .ai-response-content h5 { font-size: 0.875em; }
            //                .ai-response-content hr { height: 1px; background-color: #e1e4e8; border: 0; margin: 16px 0; }
            //                .ai-response-content ul, .ai-response-content ol { padding-left: 20px; margin: 16px 0; }
            //                .ai-response-content li { margin: 4px 0; }
            //                .ai-response-content code.inline-code {
            //                    font-family: SFMono-Regular, Consolas, Liberation Mono, Menlo, monospace;
            //                    padding: 2px 4px;
            //                    background-color: rgba(27, 31, 35, 0.05);
            //                    border-radius: 3px;
            //                    font-size: 85%;
            //                }
            //                .ai-response-content .code-block {
            //                    margin: 16px 0;
            //                }
            //                .ai-response-content .code-block pre {
            //                    background-color: #f6f8fa;
            //                    border-radius: 3px;
            //                    padding: 16px;
            //                    overflow: auto;
            //                    margin: 0;
            //                }
            //                .ai-response-content .code-block code {
            //                    font-family: SFMono-Regular, Consolas, Liberation Mono, Menlo, monospace;
            //                    font-size: 85%;
            //                    color: #24292e;
            //                    line-height: 1.45;
            //                    word-break: normal;
            //                    white-space: pre;
            //                }
            //                .ai-response-content p {
            //                    margin-top: 0;
            //                    margin-bottom: 16px;
            //                }
            //            </style>
            //            <div class="ai-response">
            //                <div class="ai-response-header">
            //                    <!-- 头部内容保持不变 -->
            //                </div>
    
            //                <div class="ai-response-content">
            //                    ${formattedResponse}
            //                </div>
        
            //                <div style="margin-top:20px; display:flex; justify-content:flex-end;">
            //                    <button id="btn-insert-response"  onclick="chatModule.insertToDocument()" style="background-color:#0078d4; color:white; border:none; padding:6px 12px; border-radius:2px; cursor:pointer; margin-right:8px;">
            //                        插入到文档
            //                    </button>
            //                    <button id="btn-copy-response" onclick="chatModule.copyToClipboard()" style="background-color:#f3f3f3; color:#333; border:none; padding:6px 12px; border-radius:2px; cursor:pointer;">
            //                        复制
            //                    </button>
            //                </div>
            //            </div>
            //            `;

            //                // 调用更新函数显示结果
            //                self.updateResultsPanel(html);

            //                // 存储结果供后续使用
            //                self.currentAIResponse = aiResponseText;
            //                // 添加按钮事件监听器 - 关键是在DOM更新后绑定事件
            //                setTimeout(function () {
            //                    var insertButton = document.getElementById("btn-insert-response");
            //                    var copyButton = document.getElementById("btn-copy-response");

            //                    if (insertButton) {
            //                        insertButton.addEventListener("click", function () {
            //                            self.insertTextToDocument(aiResponseText);
            //                        });
            //                    }

            //                    if (copyButton) {
            //                        copyButton.addEventListener("click", function () {
            //                            self.copyToClipboard(aiResponseText);
            //                        });
            //                    }
            //                }, 100); // 短暂延时确保DOM已更新
            //            })
            //    });
               
            //},

            // 添加这个辅助函数来更新结果面板
            updateResultsPanel: function (html) {
                // 尝试查找结果面板元素
                var resultsPanel = document.getElementById("results-panel");
                var resultsContent = document.getElementById("results-content");

                if (resultsContent) {
                    // 如果找到了结果内容元素，更新它
                    resultsContent.innerHTML = html;

                    // 如果有结果面板，确保它可见
                    if (resultsPanel) {
                        resultsPanel.style.display = "block";
                    }
                }
               else {
                    // 找不到结果面板元素，尝试创建一个临时的显示区域
                    console.log("结果面板元素未找到，尝试创建临时显示区域");

                    // 查找一个可以附加内容的容器
                    var container = document.querySelector(".ms-Pivot-content") ||
                        document.querySelector(".content-container") ||
                        document.querySelector("main") ||
                        document.body;

                    // 如果已存在临时结果区，则更新它
                    var tempResults = document.getElementById("temp-results-container");
                    if (tempResults) {
                        tempResults.innerHTML = html;
                        return;
                    }

                    // 创建一个新的临时结果区域
                    tempResults = document.createElement("div");
                    tempResults.id = "temp-results-container";
                    tempResults.style.margin = "20px";
                    tempResults.style.padding = "15px";
                    tempResults.style.border = "1px solid #e0e0e0";
                    tempResults.style.borderRadius = "4px";
                    tempResults.style.backgroundColor = "#fff";
                    tempResults.style.boxShadow = "0 2px 4px rgba(0,0,0,0.1)";
                    tempResults.innerHTML = html;

                    // 添加到容器
                    container.appendChild(tempResults);
                }
            },
            
            // 格式化AI响应
            // 增强版的AI响应格式化函数
            formatAIResponse: function (text) {
                if (!text) return "";

                return text
                    // 处理标题
                    .replace(/^##### (.*$)/gm, '<h5>$1</h5>')
                    .replace(/^#### (.*$)/gm, '<h4>$1</h4>')
                    .replace(/^### (.*$)/gm, '<h3>$1</h3>')
                    .replace(/^## (.*$)/gm, '<h2>$1</h2>')
                    .replace(/^# (.*$)/gm, '<h1>$1</h1>')

                    // 处理水平线
                    .replace(/^\s*(?:[-*_]){3,}\s*$/gm, '<hr />')

                    // 处理粗体和斜体
                    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                    .replace(/\*(.*?)\*/g, '<em>$1</em>')

                    // 处理列表
                    .replace(/^\s*[\-\*]\s+(.*)/gm, '<li>$1</li>')
                    .replace(/(<li>.*<\/li>\n<li>.*<\/li>)/gs, '<ul>$1</ul>')
                    .replace(/^\s*(\d+)\.\s+(.*)/gm, '<li>$2</li>')
                    .replace(/(<li>.*<\/li>\n<li>.*<\/li>)/gs, '<ol>$1</ol>')

                    // 处理代码块和内联代码
                    .replace(/```([\s\S]*?)```/g, function (match, code) {
                        return '<div class="code-block"><pre><code>' +
                            code.replace(/</g, '&lt;').replace(/>/g, '&gt;') +
                            '</code></pre></div>';
                    })
                    .replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>')

                    // 处理换行和段落
                    .replace(/\n\s*\n/g, '</p><p>') // 空行分隔段落
                    .replace(/\n/g, '<br>'); // 其他换行
            },
            // 将AI回复插入到文档中
            insertTextToDocument: function (text) {
                var self = this;
                Word.run(function (context) {
                    var selection = context.document.getSelection();
                    selection.insertText(text, Word.InsertLocation.after);

                    return context.sync()
                        .then(function () {
                            self.showNotification("成功", "已将AI回复插入到文档中");
                        });
                }).catch(function (error) {
                    console.error("插入文本错误:", error);
                    self.showNotification("错误", "插入文本失败");
                });
            },

            // 复制文本到剪贴板
            copyToClipboard: function (text) {
                var self = this;
                try {
                    navigator.clipboard.writeText(text)
                        .then(function () {
                            self.showNotification("成功", "已复制AI回复到剪贴板");
                        })
                        .catch(function (err) {
                            console.error("复制到剪贴板错误:", err);
                            self.showNotification("错误", "复制失败: " + err.message);
                        });
                } catch (e) {
                    // 如果clipboard API不可用，尝试备用方法
                    console.error("剪贴板API不可用:", e);
                    self.showNotification("信息", "无法自动复制，请手动复制文本");

                    // 显示可选择的文本区域
                    self.showResults(`
                        <div style="background-color:#f5f5f5; padding:16px; border-radius:4px;">
                            <div style="font-weight:600; margin-bottom:10px;">请手动复制以下内容:</div>
                            <textarea style="width:100%; height:200px; padding:8px;">${text}</textarea>
                        </div>
                    `);
                }
            },

            // 显示结果
            showResults: function (html) {
                try {
                    // 尝试使用现有的showResults函数
                    if (typeof window.showResults === 'function') {
                        window.showResults(html);
                        return;
                    }

                    // 否则，使用自定义实现
                    var resultsPanel = document.getElementById("results-panel");
                    var resultsContent = document.getElementById("results-content");

                    if (resultsPanel && resultsContent) {
                        resultsContent.innerHTML = html;
                        resultsPanel.style.display = "block";
                    } else {
                        console.error("未找到结果面板元素");
                    }
                } catch (e) {
                    console.error("显示结果错误:", e);
                }
            },

            // 清空结果区域
            clearResults: function () {
                try {
                    var resultsPanel = document.getElementById("results-panel");
                    if (resultsPanel) {
                        resultsPanel.style.display = "none";
                    }

                    var resultsContent = document.getElementById("results-content");
                    if (resultsContent) {
                        resultsContent.innerHTML = "";
                    }
                } catch (e) {
                    console.error("清空结果错误:", e);
                }
            },

            // 显示通知 - 不使用alert
            showNotification: function (header, content) {
                try {
                    // 尝试使用现有的showNotification函数
                    if (typeof window.showNotification === 'function') {
                        window.showNotification(header, content);
                        return;
                    }

                    console.log("通知: " + header + " - " + content);

                    // 在结果区域显示通知
                    var notificationHtml = `
                        <div style="background-color:${header.toLowerCase() === 'error' ? '#fdeceb' : '#eaf6ff'}; padding:16px; border-radius:4px; margin-bottom:16px;">
                            <div style="font-weight:600; margin-bottom:8px;">${header}</div>
                            <div>${content}</div>
                        </div>
                    `;

                    // 检查结果区域是否空闲
                    var resultsPanel = document.getElementById("results-panel");
                    if (resultsPanel && resultsPanel.style.display !== "block") {
                        // 结果区域空闲，直接使用
                        this.showResults(notificationHtml);
                    } else {
                        // 结果区域忙，创建临时通知
                        var tempNotification = document.createElement('div');
                        tempNotification.className = 'temp-notification';
                        tempNotification.style.position = 'fixed';
                        tempNotification.style.top = '10px';
                        tempNotification.style.right = '10px';
                        tempNotification.style.zIndex = '2000';
                        tempNotification.style.maxWidth = '300px';
                        tempNotification.innerHTML = notificationHtml;

                        document.body.appendChild(tempNotification);

                        // 5秒后自动移除
                        setTimeout(function () {
                            if (tempNotification && tempNotification.parentNode) {
                                tempNotification.parentNode.removeChild(tempNotification);
                            }
                        }, 5000);
                    }
                } catch (e) {
                    console.error("显示通知错误:", e);
                }
            }
        };

        // 延迟初始化以确保DOM已完全加载
        setTimeout(function () {
            chatModule.init();
        }, 1000);

        return chatModule;
    }

    // 在Office初始化后或直接调用
    if (typeof Office !== 'undefined') {
        // 检查Office是否已加载
        if (Office.context) {
            console.log("Office已加载，初始化聊天功能");
            window.chatModule = initChatFeature();
        } else {
            // 等待Office加载
            console.log("等待Office加载...");
            Office.onReady(function () {
                console.log("Office已准备就绪，初始化聊天功能");
                window.chatModule = initChatFeature();
            });
        }
    } else {
        // 非Office环境，直接初始化
        console.log("非Office环境，直接初始化聊天功能");
        setTimeout(function () {
            window.chatModule = initChatFeature();
        }, 1000);
    }
})();