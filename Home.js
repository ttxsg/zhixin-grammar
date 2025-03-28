(function () {
    "use strict";

    // API基础URL - 替换为您的实际后端地址
    const API_BASE_URL = "http://localhost:54112/api";
    // 保存的Prompt模板
    const promptTemplates = {
        "default": " :\n\n{selection}\n\n",
        "academic": "请帮我润色以下学术内容，提高其专业性和流畅度，保持学术风格:\n\n{selection}",
        "creative": "请帮我优化以下创意写作内容，提高其吸引力和表现力:\n\n{selection}",
        "technical": "请帮我优化以下技术文档，使其更清晰、准确，并易于理解:\n\n{selection}"
    };
    // Office 初始化
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Word) {
            // 建立事件监听器
            document.getElementById("check-text").onclick = checkSelectedText;
            document.getElementById("improve-clarity").onclick = () => processText("improveClarity");
            document.getElementById("improve-conciseness").onclick = () => processText("improveConciseness");
            document.getElementById("improve-tone").onclick = () => processText("improveTone");
            document.getElementById("plagiarism-check").onclick = () => processText("checkPlagiarism");
            document.getElementById("readability-score").onclick = () => processText("checkReadability");
            // 已有的事件监听器...
            document.getElementById("translate-text").onclick = translateText;
            //// 绑定聊天相关事件
            //if (document.getElementById("prompt-selector")) {
            //    document.getElementById("prompt-selector").onchange = loadPromptTemplate;
            //}
            //if (document.getElementById("save-prompt")) {
            //    document.getElementById("save-prompt").onclick = showSavePromptDialog;
            //}
            //if (document.getElementById("save-prompt-confirm")) {
            //    document.getElementById("save-prompt-confirm").onclick = savePromptTemplate;
            //}
            //if (document.getElementById("save-prompt-cancel")) {
            //    document.getElementById("save-prompt-cancel").onclick = hideSavePromptDialog;
            //}
            //if (document.getElementById("send-message")) {
            //    document.getElementById("send-message").onclick = sendMessage;
            //}

            // 初始化默认Prompt
            //initPromptTemplate();
            // 初始化通知系统
            initializeNotification();
            // 初始化OpenAI服务
            console.log("准备初始化OpenAI服务...");
            initializeSecureOpenAIService().catch(error => {
                console.error("初始化OpenAI服务失败:", error);
            });
            //initializeOpenAIService();
            console.log("OpenAI服务初始化调用完成");
        }
    });
    // 在恰当的位置（如Office.onReady回调中）初始化服务
  

    // 错误处理函数
    function handleError(error) {
        console.error('发生错误:', error);

        // 如果错误有详细信息，则显示详细信息
        if (error.debugInfo) {
            console.error('调试信息:', error.debugInfo);
        }

        // 向用户显示友好的错误消息
        let errorMessage = "操作过程中发生错误";

        if (error.message) {
            errorMessage += ": " + error.message;
        }

        // 使用横幅通知用户
        showNotification("错误", errorMessage);

        // 隐藏加载中的结果面板
        const resultsPanel = document.getElementById("results-panel");
        if (resultsPanel && resultsPanel.style.display !== "none") {
            const resultsContent = document.getElementById("results-content");
            if (resultsContent) {
                resultsContent.innerHTML = `
                <div class="ms-MessageBar ms-MessageBar--error">
                    <div class="ms-MessageBar-content">
                        <div class="ms-MessageBar-icon">
                            <i class="ms-Icon ms-Icon--ErrorBadge"></i>
                        </div>
                        <div class="ms-MessageBar-text">
                            ${errorMessage}
                        </div>
                    </div>
                </div>
            `;
            }
        }
    }
    // 初始化通知
    // 初始化通知系统 - 添加错误处理
    function initializeNotification() {
        try {
            console.log("正在初始化通知系统...");

            // 检查必要对象是否存在
            if (typeof fabric === 'undefined') {
                console.error("错误：fabric 对象未定义。Fabric UI库可能未正确加载。");
                return; // 返回但不抛出错误，允许后续代码继续执行
            }

            var element = document.querySelector('.ms-MessageBanner');
            if (!element) {
                console.error("错误：未找到 .ms-MessageBanner 元素。");
                return; // 返回但不抛出错误
            }

            messageBanner = new fabric.MessageBanner(element);
            messageBanner.init();
            console.log("通知系统初始化成功");
        } catch (error) {
            console.error("通知系统初始化失败:", error);
            // 不要抛出错误，让后续代码继续执行
        }
    }

    // 通知横幅全局变量
    var messageBanner;

    // 显示通知
    function showNotification(header, content) {
        document.getElementById("notification-header").innerHTML = header;
        document.getElementById("notification-body").innerHTML = content;
        messageBanner.show();
    }

    // 隐藏结果面板
    function hideResults() {
        document.getElementById("results-panel").style.display = "none";
    }
    // 检查选定文本
    function checkSelectedText() {
        Word.run(function (context) {
            // 获取用户选中的文本范围
            var range = context.document.getSelection();
            range.load('text');

            return context.sync()
                .then(function () {
                    if (range.text.trim() === "") {
                        showNotification("提示", "请先选择要检查的文本");
                        hideResults();
                        return;
                    }

                    // 显示"正在分析"消息
                    showResults("<div class='loading-spinner'></div><div class='ms-font-m' style='text-align:center;'>正在分析文本，请稍候...</div>");

                    // 使用OpenAI服务而不是后端API
                    OpenAIService.analyzeText(range.text)
                        .then(analysis => {
                            displayAnalysisResults(analysis);
                        })
                        .catch(err => {
                            console.error(err);
                            showResults(`<div class="ms-MessageBar ms-MessageBar--error">
                                  <div class="ms-MessageBar-content">
                                    <div class="ms-MessageBar-icon">
                                      <i class="ms-Icon ms-Icon--ErrorBadge"></i>
                                    </div>
                                    <div class="ms-MessageBar-text">
                                      分析时发生错误: ${err.message}<br>
                                      <small>请检查网络连接和API密钥设置。</small>
                                    </div>
                                  </div>
                                </div>`);
                        });
                })
                .catch(handleError);
        });
    }
    // 检查选定文本
    //function checkSelectedText() {
    //    Word.run(function (context) {
    //        // 获取用户选中的文本范围
    //        var range = context.document.getSelection();
    //        range.load('text');

    //        return context.sync()
    //            .then(function () {
    //                if (range.text.trim() === "") {
    //                    showNotification("提示", "请先选择要检查的文本");
    //                    hideResults();
    //                    return;
    //                }

    //                // 显示"正在分析"消息
    //                showResults("<div class='loading-spinner'></div><div class='ms-font-m' style='text-align:center;'>正在分析文本，请稍候...</div>");

    //                // 调用C#后端API
    //                fetch(`${API_BASE_URL}/TextAnalysis/analyze`, {
    //                    method: "POST",
    //                    headers: {
    //                        "Content-Type": "application/json"
    //                    },
    //                    body: JSON.stringify({ text: range.text })
    //                })
    //                    .then(response => {
    //                        if (!response.ok) {
    //                            throw new Error(`API请求失败: ${response.status}`);
    //                        }
    //                        return response.json();
    //                    })
    //                    .then(analysis => {
    //                        displayAnalysisResults(analysis);
    //                    })
    //                    .catch(err => {
    //                        console.error(err);
    //                        showResults(`<div class="ms-MessageBar ms-MessageBar--error">
    //                                  <div class="ms-MessageBar-content">
    //                                    <div class="ms-MessageBar-icon">
    //                                      <i class="ms-Icon ms-Icon--ErrorBadge"></i>
    //                                    </div>
    //                                    <div class="ms-MessageBar-text">
    //                                      分析时发生错误: ${err.message}<br>
    //                                      <small>请确保后端服务正在运行并且可以访问。</small>
    //                                    </div>
    //                                  </div>
    //                                </div>`);
    //                    });
    //            })
    //            .catch(handleError);
    //    });
    //}

    // 显示分析结果
    //function displayAnalysisResults(analysis) {
    //    var html = "";

    //    // 添加统计信息
    //    html += "<div class='ms-font-m ms-fontWeight-semibold'>文本统计</div>";
    //    html += "<div class='ms-Grid'>";
    //    html += "<div class='ms-Grid-row'>";
    //    html += "<div class='ms-Grid-col ms-sm6 ms-md6'>";

    //    // 检查stats对象是否存在并含有预期的属性
    //    if (analysis.stats) {
    //        html += "<div class='ms-font-s'>可读性评分: " + (analysis.stats.readabilityScore || 0) + "/100</div>";
    //        html += "<div class='ms-font-s'>字数: " + (analysis.stats.wordCount || 0) + "</div>";
    //    } else {
    //        html += "<div class='ms-font-s'>统计数据不可用</div>";
    //    }

    //    html += "</div>";
    //    html += "<div class='ms-Grid-col ms-sm6 ms-md6'>";

    //    if (analysis.stats) {
    //        html += "<div class='ms-font-s'>句子数: " + (analysis.stats.sentenceCount || 0) + "</div>";
    //        html += "<div class='ms-font-s'>平均句长: " + (analysis.stats.averageSentenceLength || 0) + " 字/句</div>";
    //    }

    //    html += "</div>";
    //    html += "</div>";
    //    html += "</div>";
    //    html += "<hr style='margin:10px 0;border:0;border-top:1px solid #eaeaea'>";

    //    // 添加错误和建议
    //    if (analysis.errors.length > 0 || analysis.suggestions.length > 0) {
    //        html += "<div class='ms-font-m ms-fontWeight-semibold'>发现的问题</div>";

    //        analysis.errors.forEach(function (error) {
    //            var cssClass = "suggestion-item " + (error.type === "grammar" ? "error-grammar" : "error-spelling");
    //            html += "<div class='" + cssClass + "'>";
    //            html += "<div class='ms-font-s ms-fontWeight-semibold'>" +
    //                (error.type === "grammar" ? "语法错误" : "拼写错误") + "</div>";
    //            html += "<div class='ms-font-s'>\"..." + error.text + "...\"</div>";
    //            html += "<div class='ms-font-s'>建议: " + error.suggestion + "</div>";
    //            html += "<div class='ms-font-xs ms-fontColor-neutralSecondary'>" + error.explanation + "</div>";
    //            html += "<button class='ms-Button ms-Button--small' onclick='applyCorrection(\"" +
    //                encodeURIComponent(error.text) + "\", \"" +
    //                encodeURIComponent(error.suggestion) + "\")'>" +
    //                "<span class='ms-Button-label'>应用修改</span></button>";
    //            html += "</div>";
    //        });

    //        analysis.suggestions.forEach(function (suggestion) {
    //            html += "<div class='suggestion-item suggestion-clarity'>";
    //            html += "<div class='ms-font-s ms-fontWeight-semibold'>表达建议</div>";
    //            html += "<div class='ms-font-s'>\"..." + suggestion.text + "...\"</div>";
    //            html += "<div class='ms-font-s'>建议: " + suggestion.suggestion + "</div>";
    //            html += "<div class='ms-font-xs ms-fontColor-neutralSecondary'>" + suggestion.explanation + "</div>";
    //            html += "<button class='ms-Button ms-Button--small' onclick='applyCorrection(\"" +
    //                encodeURIComponent(suggestion.text) + "\", \"" +
    //                encodeURIComponent(suggestion.suggestion) + "\")'>" +
    //                "<span class='ms-Button-label'>应用修改</span></button>";
    //            html += "</div>";
    //        });
    //    } else {
    //        html += "<div class='ms-font-m ms-fontColor-neutralSecondary' style='text-align:center;padding:20px;'>" +
    //            "<i class='ms-Icon ms-Icon--Emoji2' style='font-size:24px;margin-right:10px;'></i> " +
    //            "未发现问题，您的文本看起来很好！</div>";
    //    }

    //    showResults(html);

    //    // 添加全局函数以应用修改
    //    window.applyCorrection = function (originalText, newText) {
    //        applyTextCorrection(decodeURIComponent(originalText), decodeURIComponent(newText));
    //    };
    //}
    // 显示分析结果
    // 显示分析结果
    function displayAnalysisResults(analysis) {
        var html = "";

        // 添加统计信息卡片 - 使用内联样式
        html += `
        <div style="background-color:#fff; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.1); margin-bottom:20px; overflow:hidden;">
            <div style="background-color:#f0f2f5; padding:12px 16px; font-size:16px; font-weight:600; color:#0078d4; border-bottom:1px solid #e6e9ed;">文本统计</div>
            <div style="padding:16px;">
                <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(120px, 1fr)); gap:16px; margin-bottom:8px;">
    `;

        // 统计数据
        if (analysis.stats) {
            html += `
            <div style="padding:12px; text-align:center; background-color:#f8f9fa; border-radius:4px;">
                <div style="font-size:24px; font-weight:600; color:#0078d4; margin-bottom:4px;">${analysis.stats.readabilityScore || 0}/100</div>
                <div style="font-size:13px; color:#555;">可读性评分</div>
            </div>
            <div style="padding:12px; text-align:center; background-color:#f8f9fa; border-radius:4px;">
                <div style="font-size:24px; font-weight:600; color:#0078d4; margin-bottom:4px;">${analysis.stats.wordCount || 0}</div>
                <div style="font-size:13px; color:#555;">字数</div>
            </div>
            <div style="padding:12px; text-align:center; background-color:#f8f9fa; border-radius:4px;">
                <div style="font-size:24px; font-weight:600; color:#0078d4; margin-bottom:4px;">${analysis.stats.sentenceCount || 0}</div>
                <div style="font-size:13px; color:#555;">句子数</div>
            </div>
            <div style="padding:12px; text-align:center; background-color:#f8f9fa; border-radius:4px;">
                <div style="font-size:24px; font-weight:600; color:#0078d4; margin-bottom:4px;">${analysis.stats.averageSentenceLength || 0}</div>
                <div style="font-size:13px; color:#555;">平均句长</div>
            </div>
        `;
        } else {
            html += `
            <div style="grid-column:1/-1; padding:12px; text-align:center;">
                <div style="font-size:13px; color:#555;">统计数据不可用</div>
            </div>
        `;
        }

        html += `
                </div>
            </div>
        </div>
    `;

        // 添加错误卡片
        if (analysis.errors && analysis.errors.length > 0) {
            html += `
            <div style="background-color:#fff; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.1); margin-bottom:20px; overflow:hidden;">
                <div style="background-color:#f0f2f5; padding:12px 16px; font-size:16px; font-weight:600; color:#0078d4; border-bottom:1px solid #e6e9ed;">语法和拼写问题</div>
                <div style="padding:16px;">
        `;

            analysis.errors.forEach(function (error, index) {
                const errorType = error.type === "grammar" ? "语法错误" : "拼写错误";
                const typeBgColor = error.type === "grammar" ? "#fff4e5" : "#ffece8";
                const typeColor = error.type === "grammar" ? "#d9730d" : "#e03e2d";

                // 每个错误项的样式
                html += `
                <div style="position:relative; margin-bottom:${index < analysis.errors.length - 1 ? '20px' : '0'}; padding-bottom:${index < analysis.errors.length - 1 ? '20px' : '0'}; border-bottom:${index < analysis.errors.length - 1 ? '1px solid #eee' : 'none'};">
                    <div style="display:inline-block; font-size:12px; font-weight:600; padding:4px 8px; margin-bottom:8px; border-radius:3px; background-color:${typeBgColor}; color:${typeColor};">${errorType}</div>
                    <div style="background-color:#f5f7fa; border-left:4px solid #0078d4; padding:12px 16px; margin:12px 0; font-style:italic; line-height:1.6; color:#444;">"${error.text}"</div>
                    <div style="margin:12px 0; line-height:1.5;">
                        <span style="font-weight:600; color:#333; margin-right:4px;">建议:</span> ${error.suggestion}
                    </div>
                    <div style="color:#666; font-size:14px; line-height:1.5; margin:8px 0 15px 0; padding:8px; background-color:#f9f9f9; border-radius:4px;">${error.explanation}</div>
                    <button style="background-color:#0078d4; color:white; border:none; padding:8px 16px; font-size:14px; border-radius:4px; cursor:pointer; display:flex; align-items:center;" onclick="applyCorrection('${encodeURIComponent(error.text)}', '${encodeURIComponent(error.suggestion)}')">
                        <span style="margin-right:6px;"><i class="ms-Icon ms-Icon--CheckMark"></i></span>
                        <span>应用修改</span>
                    </button>
                </div>
            `;
            });

            html += `
                </div>
            </div>
        `;
        }

        // 添加表达建议卡片
        if (analysis.suggestions && analysis.suggestions.length > 0) {
            html += `
            <div style="background-color:#fff; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.1); margin-bottom:20px; overflow:hidden;">
                <div style="background-color:#f0f2f5; padding:12px 16px; font-size:16px; font-weight:600; color:#0078d4; border-bottom:1px solid #e6e9ed;">表达建议</div>
                <div style="padding:16px;">
        `;

            analysis.suggestions.forEach(function (suggestion, index) {
                // 为不同类型的建议设置不同的颜色
                let typeBgColor = "#e6f4ff";
                let typeColor = "#1677ff";

                if (suggestion.type === "conciseness") {
                    typeBgColor = "#e6ffec";
                    typeColor = "#389e0d";
                } else if (suggestion.type === "tone") {
                    typeBgColor = "#f5e8ff";
                    typeColor = "#722ed1";
                }

                html += `
                <div style="position:relative; margin-bottom:${index < analysis.suggestions.length - 1 ? '20px' : '0'}; padding-bottom:${index < analysis.suggestions.length - 1 ? '20px' : '0'}; border-bottom:${index < analysis.suggestions.length - 1 ? '1px solid #eee' : 'none'};">
                    <div style="display:inline-block; font-size:12px; font-weight:600; padding:4px 8px; margin-bottom:8px; border-radius:3px; background-color:${typeBgColor}; color:${typeColor};">表达建议</div>
                    <div style="background-color:#f5f7fa; border-left:4px solid #0078d4; padding:12px 16px; margin:12px 0; font-style:italic; line-height:1.6; color:#444;">"${suggestion.text}"</div>
                    <div style="margin:12px 0; line-height:1.5;">
                        <span style="font-weight:600; color:#333; margin-right:4px;">建议:</span> ${suggestion.suggestion}
                    </div>
                    <div style="color:#666; font-size:14px; line-height:1.5; margin:8px 0 15px 0; padding:8px; background-color:#f9f9f9; border-radius:4px;">${suggestion.explanation}</div>
                    <button style="background-color:#0078d4; color:white; border:none; padding:8px 16px; font-size:14px; border-radius:4px; cursor:pointer; display:flex; align-items:center;" onclick="applyCorrection('${encodeURIComponent(suggestion.text)}', '${encodeURIComponent(suggestion.suggestion)}')">
                        <span style="margin-right:6px;"><i class="ms-Icon ms-Icon--CheckMark"></i></span>
                        <span>应用修改</span>
                    </button>
                </div>
            `;
            });

            html += `
                </div>
            </div>
        `;
        }

        // 如果没有错误和建议
        if ((!analysis.errors || analysis.errors.length === 0) &&
            (!analysis.suggestions || analysis.suggestions.length === 0)) {
            html += `
            <div style="background-color:#fff; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.1); margin-bottom:20px; overflow:hidden;">
                <div style="padding:16px;">
                    <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; padding:30px; text-align:center; color:#107C10;">
                        <i class="ms-Icon ms-Icon--Emoji2" style="font-size:32px; margin-bottom:10px;"></i>
                        <div>未发现问题，您的文本看起来很好！</div>
                    </div>
                </div>
            </div>
        `;
        }

        showResults(html);

        // 添加全局函数以应用修改
        window.applyCorrection = function (originalText, newText) {
            applyTextCorrection(decodeURIComponent(originalText), decodeURIComponent(newText));
        };
    }

    // 应用文本修正
    function applyTextCorrection(originalText, newText) {
        Word.run(function (context) {
            // 获取文档中的所有文本
            var docText = context.document.body;
            var searchResults = docText.search(originalText, { matchCase: true, matchWholeWord: false });
            context.load(searchResults, 'text');

            return context.sync()
                .then(function () {
                    if (searchResults.items.length > 0) {
                        // 替换第一个匹配项
                        searchResults.items[0].insertText(newText, Word.InsertLocation.replace);
                        return context.sync()
                            .then(function () {
                                // 重新分析更新后的文本
                                //checkSelectedText();
                            });
                    } else {
                        showNotification("警告", "无法找到要替换的文本。");
                    }
                });
        }).catch(handleError);
    }

    // 显示结果面板
    function showResults(html) {
        var resultsPanel = document.getElementById("results-panel");
        var resultsContent = document.getElementById("results-content");

        resultsContent.innerHTML = html;
        resultsPanel.style.display = "block";
    }

    // 通用文本处理函数
    //function processText(operation) {
    //    Word.run(function (context) {
    //        var range = context.document.getSelection();
    //        range.load('text');

    //        return context.sync()
    //            .then(function () {
    //                if (range.text.trim() === "") {
    //                    showNotification("提示", "请先选择要处理的文本");
    //                    hideResults();
    //                    return;
    //                }

    //                showResults("<div class='loading-spinner'></div><div class='ms-font-m' style='text-align:center;'>正在处理，请稍候...</div>");

    //                // 确定API端点和操作
    //                let endpoint = "";
    //                let mode = "";

    //                switch (operation) {
    //                    case "improveClarity":
    //                        endpoint = "improve";
    //                        mode = "clarity";
    //                        break;
    //                    case "improveConciseness":
    //                        endpoint = "improve";
    //                        mode = "conciseness";
    //                        break;
    //                    case "improveTone":
    //                        endpoint = "improve";
    //                        mode = "tone";
    //                        break;
    //                    case "checkPlagiarism":
    //                        endpoint = "plagiarism";
    //                        break;
    //                    case "checkReadability":
    //                        endpoint = "readability";
    //                        break;
    //                }

    //                // 调用C#后端API
    //                fetch(`${API_BASE_URL}/TextAnalysis/${endpoint}`, {
    //                    method: "POST",
    //                    headers: {
    //                        "Content-Type": "application/json"
    //                    },
    //                    body: JSON.stringify({
    //                        text: range.text,
    //                        mode: mode
    //                    })
    //                })
    //                    .then(response => {
    //                        if (!response.ok) {
    //                            throw new Error(`API请求失败: ${response.status}`);
    //                        }
    //                        return response.json();
    //                    })
    //                    .then(result => {
    //                        // 处理不同类型的结果
    //                        switch (operation) {
    //                            case "improveClarity":
    //                            case "improveConciseness":
    //                            case "improveTone":
    //                                showImprovedText(result.generatedText, range.text, operation);
    //                                break;
    //                            case "checkPlagiarism":
    //                                showPlagiarismResult(result);
    //                                break;
    //                            case "checkReadability":
    //                                showReadabilityResult(result);
    //                                break;
    //                        }
    //                    })
    //                    .catch(err => {
    //                        console.error(err);
    //                        showResults(`<div class="ms-MessageBar ms-MessageBar--error">
    //                                  <div class="ms-MessageBar-content">
    //                                    <div class="ms-MessageBar-icon">
    //                                      <i class="ms-Icon ms-Icon--ErrorBadge"></i>
    //                                    </div>
    //                                    <div class="ms-MessageBar-text">
    //                                      处理时发生错误: ${err.message}<br>
    //                                      <small>请确保后端服务正在运行并且可以访问。</small>
    //                                    </div>
    //                                  </div>
    //                                </div>`);
    //                    });
    //            })
    //            .catch(handleError);
    //    });
    //}

    // 显示AI改进的文本

    // 通用文本处理函数
    function processText(operation) {
        Word.run(function (context) {
            var range = context.document.getSelection();
            range.load('text');

            return context.sync()
                .then(function () {
                    if (range.text.trim() === "") {
                        showNotification("提示", "请先选择要处理的文本");
                        hideResults();
                        return;
                    }

                    showResults("<div class='loading-spinner'></div><div class='ms-font-m' style='text-align:center;'>正在处理，请稍候...</div>");

                    // 根据操作类型使用OpenAI服务
                    switch (operation) {
                        case "improveClarity":
                        case "improveConciseness":
                        case "improveTone":
                            // 从操作名称获取模式
                            const mode = operation.replace('improve', '').toLowerCase();

                            // 使用OpenAI服务
                            OpenAIService.improveText(range.text, mode)
                                .then(improvedText => {
                                    showImprovedText(improvedText, range.text, operation);
                                })
                                .catch(handleOpenAIError);
                            break;

                        case "checkPlagiarism":
                            OpenAIService.checkPlagiarism(range.text)
                                .then(result => {
                                    showPlagiarismResult(result);
                                })
                                .catch(handleOpenAIError);
                            break;

                        case "checkReadability":
                            OpenAIService.calculateReadability(range.text)
                                .then(result => {
                                    showReadabilityResult(result);
                                })
                                .catch(handleOpenAIError);
                            break;
                    }
                })
                .catch(handleError);
        });
    }

    // 处理OpenAI服务错误
    function handleOpenAIError(err) {
        console.error(err);
        showResults(`<div class="ms-MessageBar ms-MessageBar--error">
              <div class="ms-MessageBar-content">
                <div class="ms-MessageBar-icon">
                  <i class="ms-Icon ms-Icon--ErrorBadge"></i>
                </div>
                <div class="ms-MessageBar-text">
                  处理时发生错误: ${err.message}<br>
                  <small>请检查网络连接和API设置。</small>
                </div>
              </div>
            </div>`);
    }
    function showImprovedText(improvedText, originalText, operation) {
        var operationTitle = "";
        switch (operation) {
            case "improveClarity":
                operationTitle = "更清晰的表达";
                break;
            case "improveConciseness":
                operationTitle = "更简洁的表达";
                break;
            case "improveTone":
                operationTitle = "调整后的语气";
                break;
            default:
                operationTitle = "AI改进建议";
                break;
        }

        var html = `
            <div class='ms-font-m ms-fontWeight-semibold'>${operationTitle}</div>
            <div class='ms-font-s' style='margin-bottom:10px;'>原文本：</div>
            <div class='ms-font-s' style='background-color:#f3f3f3;padding:10px;margin-bottom:15px;'>${originalText}</div>
            <div class='ms-font-s' style='margin-bottom:10px;'>改进建议：</div>
            <div class='ms-font-s' style='background-color:#eaf6ff;padding:10px;margin-bottom:15px;border-left:3px solid #0078d4'>${improvedText}</div>
            <div class='ms-font-s' style='text-align:right;'>
                <button class='ms-Button ms-Button--primary' id='apply-improved'>
                    <span class='ms-Button-label'>应用改进后的文本</span>
                </button>
            </div>
        `;

        showResults(html);

        // 添加应用按钮事件
        document.getElementById('apply-improved').addEventListener('click', function () {
            applyImprovedText(improvedText);
        });
    }

    // 应用改进后的文本
    function applyImprovedText(improvedText) {
        Word.run(function (context) {
            var selection = context.document.getSelection();
            selection.insertText(improvedText, Word.InsertLocation.replace);

            return context.sync()
                .then(function () {
                    showNotification("成功", "已应用改进后的文本");
                    hideResults();
                });
        }).catch(handleError);
    }

    // 显示抄袭检测结果
    function showPlagiarismResult(result) {
        var html = `
            <div class='ms-font-m ms-fontWeight-semibold'>原创度分析</div>
            <div class='ms-font-s' style='margin:15px 0;'>原创度评分:</div>
            <div class='progress-container'>
                <div class='progress-bar' style='width:${result.originalityScore}%;'></div>
            </div>
            <div class='ms-font-xl ms-fontWeight-semibold' style='text-align:center;margin:15px 0;'>${result.originalityScore}%</div>
            <div class='ms-font-s' style='margin-bottom:10px;'>分析报告:</div>
            <div class='ms-font-s' style='background-color:#f9f9f9;padding:10px;'>${result.analysis}</div>
        `;

        if (result.similarSources && result.similarSources.length > 0) {
            html += `<div class='ms-font-s' style='margin:10px 0 5px;'>发现相似来源:</div>`;
            html += `<ul class='ms-font-s'>`;
            result.similarSources.forEach(function (source) {
                html += `<li>${source}</li>`;
            });
            html += `</ul>`;
        }

        showResults(html);
    }

    // 显示可读性评分结果
    function showReadabilityResult(result) {
        var scoreLevel = "";
        if (result.score >= 90) scoreLevel = "非常容易读";
        else if (result.score >= 80) scoreLevel = "容易读";
        else if (result.score >= 70) scoreLevel = "相对容易";
        else if (result.score >= 60) scoreLevel = "中等";
        else if (result.score >= 50) scoreLevel = "相对困难";
        else scoreLevel = "阅读困难";

        var html = `
            <div class='ms-font-m ms-fontWeight-semibold'>可读性分析</div>
            <div class='ms-font-s' style='margin:15px 0;'>可读性评分:</div>
            <div class='progress-container'>
                <div class='progress-bar' style='width:${result.score}%;background-color:${getScoreColor(result.score)}'></div>
            </div>
            <div class='ms-font-xl ms-fontWeight-semibold' style='text-align:center;margin:10px 0;'>${result.score}/100 <span class='ms-font-m'>(${scoreLevel})</span></div>
            
            <div class='ms-Grid' style='margin-top:15px;'>
                <div class='ms-Grid-row'>
                    <div class='ms-Grid-col ms-sm6 ms-md6'>
                        <div class='ms-font-s'>词数: ${result.wordCount}</div>
                        <div class='ms-font-s'>句子数: ${result.sentenceCount}</div>
                    </div>
                    <div class='ms-Grid-col ms-sm6 ms-md6'>
                        <div class='ms-font-s'>平均句长: ${result.averageSentenceLength} 词/句</div>
                        <div class='ms-font-s'>复杂词比例: ${result.complexWordPercentage}%</div>
                    </div>
                </div>
            </div>
            
            <div class='ms-font-s' style='margin:15px 0 5px;'>AI分析:</div>
            <div class='ms-font-s' style='background-color:#f9f9f9;padding:10px;'>${result.aiAnalysis}</div>
        `;

        showResults(html);
    }

    // 获取评分颜色
    function getScoreColor(score) {
        if (score >= 80) return "#107C10"; // 绿色
        if (score >= 60) return "#2D7D9A"; // 蓝色
        if (score >= 40) return "#FF8C00"; // 橙色
        return "#E81123"; // 红色
    }
    // 翻译文本
    function translateText() {
        Word.run(function (context) {
            var range = context.document.getSelection();
            range.load('text');

            return context.sync()
                .then(function () {
                    if (range.text.trim() === "") {
                        showNotification("提示", "请先选择要翻译的文本");
                        hideResults();
                        return;
                    }

                    showResults(`
                <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; padding:20px;">
                    <div style="width:40px; height:40px; border:4px solid #f3f3f3; border-top:4px solid #0078d4; border-radius:50%; animation:spin 1s linear infinite;"></div>
                    <div style="margin-top:16px; text-align:center;">正在翻译，请稍候...</div>
                </div>
                <style>
                    @keyframes spin {
                        0% { transform: rotate(0deg); }
                        100% { transform: rotate(360deg); }
                    }
                </style>
            `);

                    // 使用OpenAI服务进行翻译
                    OpenAIService.translateText(range.text, "zh-CN")
                        .then(translatedText => {
                            showTranslationResult(translatedText, range.text);
                        })
                        .catch(err => {
                            console.error(err);
                            showResults(`
                        <div style="background-color:#fdeceb; border-radius:4px; padding:16px; display:flex; align-items:flex-start;">
                            <div style="font-size:24px; color:#a80000; margin-right:16px;">
                                <i class="ms-Icon ms-Icon--ErrorBadge"></i>
                            </div>
                            <div>
                                <div style="font-weight:600; margin-bottom:4px;">翻译时发生错误: ${err.message}</div>
                                <div style="font-size:14px; opacity:0.9;">请检查网络连接和API设置。</div>
                            </div>
                        </div>
                    `);
                        });
                })
                .catch(handleError);
        });
    }
    // 翻译文本
    //function translateText() {
    //    Word.run(function (context) {
    //        var range = context.document.getSelection();
    //        range.load('text');

    //        return context.sync()
    //            .then(function () {
    //                if (range.text.trim() === "") {
    //                    showNotification("提示", "请先选择要翻译的文本");
    //                    hideResults();
    //                    return;
    //                }

    //                showResults(`
    //                <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; padding:20px;">
    //                    <div style="width:40px; height:40px; border:4px solid #f3f3f3; border-top:4px solid #0078d4; border-radius:50%; animation:spin 1s linear infinite;"></div>
    //                    <div style="margin-top:16px; text-align:center;">正在翻译，请稍候...</div>
    //                </div>
    //                <style>
    //                    @keyframes spin {
    //                        0% { transform: rotate(0deg); }
    //                        100% { transform: rotate(360deg); }
    //                    }
    //                </style>
    //            `);

    //                // 调用翻译API
    //                fetch(`${API_BASE_URL}/TextAnalysis/translate`, {
    //                    method: "POST",
    //                    headers: {
    //                        "Content-Type": "application/json"
    //                    },
    //                    body: JSON.stringify({
    //                        text: range.text,
    //                        targetLanguage: "zh-CN"  // 目标语言：中文
    //                    })
    //                })
    //                    .then(response => {
    //                        if (!response.ok) {
    //                            throw new Error(`API请求失败: ${response.status}`);
    //                        }
    //                        return response.json();
    //                    })
    //                    .then(result => {
    //                        showTranslationResult(result.translatedText, range.text);
    //                    })
    //                    .catch(err => {
    //                        console.error(err);
    //                        showResults(`
    //                        <div style="background-color:#fdeceb; border-radius:4px; padding:16px; display:flex; align-items:flex-start;">
    //                            <div style="font-size:24px; color:#a80000; margin-right:16px;">
    //                                <i class="ms-Icon ms-Icon--ErrorBadge"></i>
    //                            </div>
    //                            <div>
    //                                <div style="font-weight:600; margin-bottom:4px;">翻译时发生错误: ${err.message}</div>
    //                                <div style="font-size:14px; opacity:0.9;">请确保后端服务正在运行并且可以访问。</div>
    //                            </div>
    //                        </div>
    //                    `);
    //                    });
    //            })
    //            .catch(handleError);
    //    });
    //}

    // 显示翻译结果
    function showTranslationResult(translatedText, originalText) {
        var html = `
        <div style="background-color:#fff; border-radius:6px; box-shadow:0 2px 8px rgba(0,0,0,0.1); margin-bottom:20px; overflow:hidden;">
            <div style="background-color:#f0f2f5; padding:12px 16px; font-size:16px; font-weight:600; color:#0078d4; border-bottom:1px solid #e6e9ed;">翻译结果</div>
            <div style="padding:16px;">
                <div style="display:flex; flex-direction:column; gap:16px; margin-bottom:16px;">
                    <div style="flex:1;">
                        <div style="font-weight:600; margin-bottom:8px; color:#333;">原文</div>
                        <div style="padding:12px; background-color:#f5f5f5; border-left:4px solid #666; border-radius:4px; white-space:pre-wrap; line-height:1.5; min-height:120px; max-height:300px; overflow-y:auto;">${originalText}</div>
                    </div>
                    <div style="flex:1;">
                        <div style="font-weight:600; margin-bottom:8px; color:#333;">中文翻译</div>
                        <div style="padding:12px; background-color:#eaf6ff; border-left:4px solid #0078d4; border-radius:4px; white-space:pre-wrap; line-height:1.5; min-height:120px; max-height:300px; overflow-y:auto;">${translatedText}</div>
                    </div>
                </div>
                <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:16px;">
                    <button id="apply-translation" style="display:flex; align-items:center; gap:6px; padding:10px 18px; font-size:14px; border-radius:4px; cursor:pointer; border:none; background-color:#0078d4; color:white;">
                        <span style="display:flex; align-items:center;"><i class="ms-Icon ms-Icon--CheckMark"></i></span>
                        <span>应用翻译结果</span>
                    </button>
                </div>
            </div>
        </div>
    `;

        showResults(html);

        // 添加应用翻译按钮事件
        document.getElementById('apply-translation').addEventListener('click', function () {
            applyTranslationText(translatedText);
        });
    }

    // 应用翻译后的文本
    function applyTranslationText(translatedText) {
        Word.run(function (context) {
            var selection = context.document.getSelection();
            selection.insertText(translatedText, Word.InsertLocation.replace);

            return context.sync()
                .then(function () {
                    showNotification("成功", "已应用翻译后的文本");
                    hideResults();
                });
        }).catch(function (error) {
            handleError(error);
        });
    }
    
})();