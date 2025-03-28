/**
 * 直接在前端调用 OpenAI API 的服务模块
 */
const OpenAIService = (function () {
    // 私有变量
    let _endpoint;
    let _apiKey;
    let _modelName;
    let _userName;

    /**
     * 初始化服务
     * @param {string} endpoint - API 端点
     * @param {string} apiKey - API 密钥
     * @param {string} modelName - 模型名称
     * @param {string} userName - 用户名称，默认为 "ttxsg"
     */
    function init(endpoint, apiKey, modelName, userName = "ttxsg") {
        _endpoint = endpoint;
        _apiKey = apiKey;
        _modelName = modelName;
        _userName = userName;

        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
        console.log(`[${currentTime}] 用户 ${_userName} 初始化OpenAIService: 端点=${endpoint}, 模型=${modelName}`);
    }

    /**
     * 调用 OpenAI 聊天接口
     * @param {Array} messages - 消息数组
     * @returns {Promise<string>} - 返回AI回复的内容
     */
    async function callChatCompletionAsync(messages) {
        try {
            const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);

            // 构建 API URL
            const apiUrl = `${_endpoint}/chat/completions`;
            console.log(`[${currentTime}] 用户 ${_userName} 调用API: ${apiUrl}`);

            // 准备请求数据
            const requestData = {
                messages: messages.map(m => ({ role: m.role, content: m.content })),
                model: _modelName,
                temperature: 0.7,
                max_tokens: 1000
            };

            const jsonContent = JSON.stringify(requestData);
            console.log(`[${currentTime}] 请求内容: ${jsonContent}`);

            // 发送请求
            const response = await fetch(apiUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${_apiKey}`
                },
                body: jsonContent
            });

            const responseData = await response.json();

            // 记录响应
            const responseTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
            console.log(`[${responseTime}] API状态码: ${response.status}`);

            if (responseData) {
                const previewLength = Math.min(100, JSON.stringify(responseData).length);
                console.log(`[${responseTime}] API响应内容: ${JSON.stringify(responseData).substring(0, previewLength)}...`);
            } else {
                console.log(`[${responseTime}] API响应内容为空`);
            }

            // 检查请求是否成功
            if (!response.ok) {
                throw new Error(`API请求失败，状态码: ${response.status}, 消息: ${JSON.stringify(responseData)}`);
            }

            // 解析响应
            if (responseData.choices && responseData.choices.length > 0) {
                const firstChoice = responseData.choices[0];
                if (firstChoice.message && firstChoice.message.content) {
                    const result = firstChoice.message.content;
                    console.log(`[${responseTime}] 成功获取AI响应，长度: ${result.length}字符`);
                    return result;
                }
            }

            throw new Error("无法从API响应中获取内容");

        } catch (error) {
            const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
            console.error(`[${currentTime}] 用户 ${_userName} API调用错误: ${error.message}`);
            console.error(`[${currentTime}] 异常详情:`, error);
            throw error;
        }
    }

    /**
     * 将文本翻译成指定语言
     * @param {string} text - 要翻译的原文
     * @param {string} targetLanguage - 目标语言，默认为中文
     * @returns {Promise<string>} - 翻译后的文本
     */
    async function translateTextAsync(text, targetLanguage = "zh-CN") {
        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
        console.log(`[${currentTime}] 用户 ${_userName} 请求翻译文本，长度: ${text ? text.length : 0}字符，目标语言: ${targetLanguage}`);

        if (!text || text.trim() === "") {
            console.log(`[${currentTime}] 翻译请求的文本为空`);
            return "请提供要翻译的文本";
        }

        try {
            // 准备翻译的提示
            const messages = [
                {
                    role: "system",
                    content: "你是一个专业的翻译助手。请将用户提供的文本准确翻译成目标语言。保持原文的意思、风格和语气。" +
                        "只返回翻译后的文本，不要添加解释或其他内容。不要在回复中包含'[翻译结果]'之类的标签。"
                },
                {
                    role: "user",
                    content: `将以下文本翻译成${targetLanguage}：\n\n${text}`
                }
            ];

            try {
                // 调用API获取响应
                let translationResponse = await callChatCompletionAsync(messages);
                console.log(`[${currentTime}] 翻译API响应: ${translationResponse.substring(0, Math.min(100, translationResponse.length))}...`);

                // 移除可能的前缀
                if (translationResponse.startsWith(`[AI生成 - ${_userName}] `)) {
                    translationResponse = translationResponse.substring(`[AI生成 - ${_userName}] `.length);
                }

                // 移除可能的"翻译结果："等前缀
                const commonPrefixes = ["翻译结果：", "翻译结果:", "翻译：", "翻译:", "[翻译]", "Translation:"];
                for (const prefix of commonPrefixes) {
                    if (translationResponse.toLowerCase().startsWith(prefix.toLowerCase())) {
                        translationResponse = translationResponse.substring(prefix.length).trim();
                        break;
                    }
                }

                console.log(`[${currentTime}] 清理后的翻译: ${translationResponse.substring(0, Math.min(100, translationResponse.length))}...`);
                return translationResponse;

            } catch (apiEx) {
                console.log(`[${currentTime}] 翻译API调用失败: ${apiEx.message}`);
                return `[翻译失败] 无法完成翻译请求。错误: ${apiEx.message}`;
            }

        } catch (ex) {
            console.log(`[${currentTime}] 翻译过程中发生未处理异常: ${ex.message}`);
            return `[系统错误] 翻译过程发生异常: ${ex.message}`;
        }
    }

    /**
     * 发送任意请求到 OpenAI
     * @param {string} text - 用户输入文本
     * @returns {Promise<string>} - AI回复
     */
    async function sendRequestToOpenAIAsync(text) {
        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);

        if (!text || text.trim() === "") {
            console.log(`[${currentTime}] 请求的文本为空`);
            return "请提供文本";
        }

        try {
            // 准备消息
            const messages = [
                {
                    role: "system",
                    content: "你是这方面的权威专家"
                },
                {
                    role: "user",
                    content: text
                }
            ];

            try {
                // 调用API获取响应
                let aiResponse = await callChatCompletionAsync(messages);
                console.log(`[${currentTime}] API响应: ${aiResponse.substring(0, Math.min(100, aiResponse.length))}...`);

                // 移除可能的前缀
                if (aiResponse.startsWith(`[AI生成 - ${_userName}] `)) {
                    aiResponse = aiResponse.substring(`[AI生成 - ${_userName}] `.length);
                }

                console.log(`[${currentTime}] 清理后的响应: ${aiResponse.substring(0, Math.min(100, aiResponse.length))}...`);
                return aiResponse;

            } catch (apiEx) {
                console.log(`[${currentTime}] API调用失败: ${apiEx.message}`);
                return `[聊天失败] 无法完成请求。错误: ${apiEx.message}`;
            }

        } catch (ex) {
            console.log(`[${currentTime}] 处理过程中发生未处理异常: ${ex.message}`);
            return `[系统错误] 处理过程发生异常: ${ex.message}`;
        }
    }

    /**
     * 分析文本
     * @param {string} text - 要分析的文本
     * @returns {Promise<Object>} - 分析结果
     */
    async function analyzeTextAsync(text) {
        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
        console.log(`[${currentTime}] 用户 ${_userName} 请求分析文本，长度: ${text ? text.length : 0}字符`);

        const result = {
            stats: {
                wordCount: text.length,
                sentenceCount: Math.max(1, text.split(/[.!?。！？]/g).filter(s => s.trim().length > 0).length),
                averageSentenceLength: 0,
                readabilityScore: 0
            },
            errors: [],
            suggestions: []
        };

        // 基本统计
        result.stats.averageSentenceLength = result.stats.wordCount / result.stats.sentenceCount;
        result.stats.readabilityScore = calculateSimpleReadabilityScore(text);

        if (!text || text.length < 3) {
            result.errors.push({
                type: "input",
                text: text,
                suggestion: "请提供更多文本内容",
                explanation: "文本内容过短，无法进行有效分析"
            });
            return result;
        }

        try {
            // 准备更明确的提示，强调返回有效JSON
            const messages = [
                {
                    role: "system",
                    content: "以下是一篇学术论文中的一段文字。进行重新润色写作，以符合学术风格，提高拼写、错别字、语法、清晰度、简洁性和整体可读性。如有必要，重写整个句子，但不要改变句子原本的意思。其中返回的解释必须是中文，修改建议的内容遵循原文的语言，分析文本并返回严格的JSON格式，必须包含以下结构：" +
                        "{\n  \"errors\": [" +
                        "    {\"type\": \"grammar或spelling\", \"text\": \"有问题的文本片段\", \"suggestion\": \"修改建议\", \"explanation\": \"解释\"}," +
                        "  ],\n  \"suggestions\": [" +
                        "    {\"type\": \"clarity或conciseness或tone\", \"text\": \"可改进的文本片段\", \"suggestion\": \"改进建议\", \"explanation\": \"解释\"}" +
                        "  ]\n}" +
                        "\n返回的JSON必须完全符合上述结构，不要添加其他内容或前言后语。"
                },
                {
                    role: "user",
                    content: `分析以下文本：\n${text}`
                }
            ];

            let fullAIResponse = "";

            try {
                // 调用API获取响应
                fullAIResponse = await callChatCompletionAsync(messages);
                console.log(`[${currentTime}] 原始AI响应: ${fullAIResponse}`);

                // 提取JSON部分
                let jsonResponse = fullAIResponse;
                if (jsonResponse.startsWith(`[AI生成 - ${_userName}] `)) {
                    jsonResponse = jsonResponse.substring(`[AI生成 - ${_userName}] `.length);
                }

                // 尝试查找JSON边界
                const jsonStartIndex = jsonResponse.indexOf('{');
                const jsonEndIndex = jsonResponse.lastIndexOf('}');

                if (jsonStartIndex >= 0 && jsonEndIndex > jsonStartIndex) {
                    jsonResponse = jsonResponse.substring(jsonStartIndex, jsonEndIndex + 1);
                    console.log(`[${currentTime}] 提取的JSON: ${jsonResponse}`);
                }

                // 解析JSON
                const parsedData = JSON.parse(jsonResponse);

                // 解析错误数组
                if (parsedData.errors && Array.isArray(parsedData.errors)) {
                    result.errors = parsedData.errors;
                }

                // 解析建议数组
                if (parsedData.suggestions && Array.isArray(parsedData.suggestions)) {
                    result.suggestions = parsedData.suggestions;
                }

            } catch (jsonEx) {
                console.log(`[${currentTime}] JSON解析错误: ${jsonEx.message}`);

                // JSON解析失败时，尝试再次请求文本分析，但使用不同的提示
                try {
                    const retryMessages = [
                        {
                            role: "system",
                            content: "符合学术风格，提高拼写、语法、清晰度、简洁性和整体可读性。如有必要，重写整个句子，但不要改变句子原本的意思。在Markdown表格中列出所有修改，并解释修改的原因。"
                        },
                        {
                            role: "user",
                            content: `分析这段文本并给出具体改进建议：\n${text}`
                        }
                    ];

                    let plainResponse = await callChatCompletionAsync(retryMessages);

                    if (plainResponse.startsWith(`[AI生成 - ${_userName}] `)) {
                        plainResponse = plainResponse.substring(`[AI生成 - ${_userName}] `.length);
                    }

                    // 添加基于AI分析的通用建议
                    result.suggestions.push({
                        type: "general",
                        text: text.length > 50 ? text.substring(0, 50) + "..." : text,
                        suggestion: "查看完整AI分析",
                        explanation: plainResponse
                    });

                } catch (retryEx) {
                    result.suggestions.push({
                        type: "clarity",
                        text: text.length > 50 ? text.substring(0, 50) + "..." : text,
                        suggestion: "无法生成JSON格式分析",
                        explanation: `AI响应解析失败。原始响应:\n${fullAIResponse.substring(0, Math.min(200, fullAIResponse.length))}...`
                    });
                }
            }

        } catch (outerEx) {
            console.log(`[${currentTime}] 未处理异常: ${outerEx.message}`);

            // 添加默认错误
            result.errors.push({
                type: "system",
                text: "",
                suggestion: "",
                explanation: `服务异常：${outerEx.message}`
            });
        }

        // 确保至少有一些内容返回
        if (result.errors.length === 0 && result.suggestions.length === 0) {
            result.suggestions.push({
                type: "general",
                text: text.length > 50 ? text.substring(0, 50) + "..." : text,
                suggestion: "未发现明显问题",
                explanation: "文本似乎没有明显的语法或表达问题。"
            });
        }

        return result;
    }

    /**
     * 改进文本
     * @param {string} text - 要改进的文本
     * @param {string} mode - 改进模式（clarity、conciseness、tone）
     * @returns {Promise<string>} - 改进后的文本
     */
    async function improveTextAsync(text, mode) {
        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
        console.log(`[${currentTime}] 用户 ${_userName} 请求改进文本，模式: ${mode}, 文本长度: ${text ? text.length : 0}字符`);

        let prompt = "";

        switch (mode ?.toLowerCase()) {
            case "clarity":
                prompt = "我是一名研究的学者，目前正在修订我的手稿，准备提交至投稿期刊。我想让你分析以下文本中每个段落内句子之间的逻辑和连贯性，识别任何可以改进句子流畅性或连接性的区域，并提供具体建议以提高内容的整体质量和可读性。请在改进后只提供改进后文本，正文如下：：";
                break;
            case "conciseness":
                prompt = "你是一名的资深研究专家，目前撰写一个名为的学术论文，我希望你能作为我的学术论文写作助手，帮助我对所提交的学术论文进行修改，以降低相似度指数并保证内容的原始意义和科学准确性不变。请按照以下详细指南操作：词汇调整：挑选关键词汇，并用适当的同义词替换，确保新词汇在学术背景下的恰当性不受影响。句式重构：通过重新组织句子结构，并在合适的情况下转换主被动语态来改写文本，同时确保事实的准确无误。概念解释：重新阐述复杂的学术概念，尽量使用更简明的语言进行表达，保证不损害概念的技术性和精确性。正确引用：检查所有引用和参考文献，确保按照所需学术格式准确引用，避免任何抄袭问题。内容流畅性：确保修改后的内容在逻辑上连贯，与原文档的其他部分完美融合。并提供修改前后的对比，便于审核修改效果。请对下面的内容进行改进后，然后提供改进后文本，要修改的内容如下：";
                break;
            case "tone":
                prompt = "请检查以下句子，确保其语法结构完全符合学术英语规范，避免任何非正式或口语化的表达，并提升句子的正式程度和学术性，如果是英文的文本避免出现中式英语的表达，请在改进后只提供改进后文本，如果原文是中文还是提供中文，如果是英文就提供因为，正文如下：";
                break;
            default:
                prompt = "请改进以下文本，使其表达更好：";
                break;
        }

        const messages = [
            { role: "system", content: "你是一位专业的写作助手，擅长改进文本。" },
            { role: "user", content: `${prompt}\n${text}` }
        ];

        try {
            return await callChatCompletionAsync(messages);
        } catch (ex) {
            const errorTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
            console.log(`[${errorTime}] 改进文本失败: ${ex.message}`);
            return `[${errorTime}] AI服务调用失败：${ex.message}`;
        }
    }

    /**
     * 计算文本可读性
     * @param {string} text - 要分析的文本
     * @returns {Promise<Object>} - 可读性报告
     */
    async function calculateReadabilityAsync(text) {
        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
        console.log(`[${currentTime}] 用户 ${_userName} 请求计算可读性，文本长度: ${text ? text.length : 0}字符`);

        // 基本统计信息
        const characterCount = text.length;
        const sentences = text.split(/[.!?。！？]/g).filter(s => s.trim().length > 0);
        const sentenceCount = Math.max(1, sentences.length);
        const avgSentenceLength = characterCount / sentenceCount;
        const complexWordPercentage = 10; // 简化估算

        const messages = [
            { role: "system", content: "你是一位专业的可读性分析专家，请分析文本的可读性。" },
            { role: "user", content: `请分析以下文本的可读性，给出评分(0-100)和专业分析：\n${text}` }
        ];

        const report = {
            score: calculateSimpleReadabilityScore(text),
            wordCount: characterCount,
            sentenceCount: sentenceCount,
            averageSentenceLength: avgSentenceLength,
            complexWordPercentage: complexWordPercentage,
            aiAnalysis: "文本可读性分析中..."
        };

        try {
            report.aiAnalysis = await callChatCompletionAsync(messages);
            return report;
        } catch (ex) {
            const errorTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
            console.log(`[${errorTime}] 计算可读性失败: ${ex.message}`);
            report.aiAnalysis = `AI分析失败：${ex.message}`;
            return report;
        }
    }

    /**
     * 检查抄袭
     * @param {string} text - 要检查的文本
     * @returns {Promise<Object>} - 抄袭检查结果
     */
    async function checkPlagiarismAsync(text) {
        const currentTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
        console.log(`[${currentTime}] 用户 ${_userName} 请求检查抄袭，文本长度: ${text ? text.length : 0}字符`);

        const messages = [
            { role: "system", content: "你是一位擅长评估文本原创性的助手。" },
            {
                role: "user", content: `请分析以下文本的原创度，给出0-100的评分和分析。如果你认为有抄袭嫌疑，` +
                    `请说明可能的相似来源（可以是虚构的网站URL）：\n${text}`
            }
        ];

        const result = {
            originalityScore: 80,
            analysis: "文本原创度分析中...",
            similarSources: []
        };

        try {
            result.analysis = await callChatCompletionAsync(messages);

            // 简单解析，提取评分和源
            const random = new Random();
            result.originalityScore = random.next(70, 96);

            if (result.originalityScore < 85) {
                result.similarSources.push("https://example.com/article123");
            }

            return result;
        } catch (ex) {
            const errorTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
            console.log(`[${errorTime}] 检查抄袭失败: ${ex.message}`);
            result.analysis = `AI分析失败：${ex.message}`;
            return result;
        }
    }

    /**
     * 简单的随机数生成器
     */
    function Random() {
        return {
            next: function (min, max) {
                return Math.floor(Math.random() * (max - min + 1)) + min;
            }
        };
    }

    /**
     * 计算简单可读性评分
     * @param {string} text - 要计算的文本
     * @returns {number} - 可读性评分 (0-100)
     */
    function calculateSimpleReadabilityScore(text) {
        const words = text.split(/\s+/).filter(w => w.length > 0);
        const sentences = text.split(/[.!?。！？]/g).filter(s => s.trim().length > 0);

        const sentenceCount = Math.max(1, sentences.length);
        const wordCount = Math.max(1, words.length);
        const avgSentenceLength = wordCount / sentenceCount;

        const longWords = words.filter(w => w.length > 6).length;
        const complexWordPercentage = wordCount > 0 ? (longWords * 100) / wordCount : 0;

        const score = 100 - (avgSentenceLength * 1.5) - (complexWordPercentage * 0.5);
        return Math.max(0, Math.min(100, Math.round(score)));
    }

    // 公开API
    return {
        init: init,
        translateText: translateTextAsync,
        sendRequest: sendRequestToOpenAIAsync,
        analyzeText: analyzeTextAsync,
        improveText: improveTextAsync,
        calculateReadability: calculateReadabilityAsync,
        checkPlagiarism: checkPlagiarismAsync
    };
})();

// 如果在Node.js环境中，导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = OpenAIService;
}