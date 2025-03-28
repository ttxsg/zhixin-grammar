// 安全的API密钥管理
async function initializeSecureOpenAIService() {
    console.log("初始化OpenAI服务...");
    try {
        // 检查OpenAIService是否已加载
        if (typeof OpenAIService === 'undefined') {
            throw new Error("OpenAIService未定义，请确保openai-service.js已正确加载");
        }

        // 获取API密钥和配置
        const config = await getSecureConfig();

        // 使用获取的配置初始化服务
        console.log(`使用模型: ${config.modelName} 初始化服务...`);
        OpenAIService.init(
            config.endpoint,
            config.apiKey,
            config.modelName,
            "ttxsg"
        );

        console.log("OpenAI服务初始化成功");
    } catch (error) {
        console.error("初始化OpenAI服务失败:", error);
        showNotification("错误", "AI服务初始化失败，请检查网络连接后重试");
    }
}

// 获取安全配置 - 首先尝试从本地文件获取，如果失败则从远程API获取
async function getSecureConfig() {
    try {
        // 尝试从本地文件读取配置
        const localConfig = await readLocalConfigFile();

        // 如果本地配置有效，直接使用
        if (localConfig && localConfig.apiKey && localConfig.apiKey.length > 20) {
            console.log("使用本地配置文件中的API密钥");
            return {
                endpoint: localConfig.endpoint || "https://models.inference.ai.azure.com/",
                apiKey: localConfig.apiKey,
                modelName: localConfig.modelName || "GPT-4o"
            };
        }

        // 否则，从远程获取
        console.log("本地配置无效，正在从远程获取API密钥...");
        return await fetchRemoteConfig();
    } catch (error) {
        console.error("获取配置失败:", error);
        throw new Error("无法获取API配置: " + error.message);
    }
}

// 读取本地配置文件
async function readLocalConfigFile() {
    return new Promise((resolve, reject) => {
        try {
            // 在Word加载项中，使用Office.context.document.settings读取设置
            // 而不是直接读取文件系统（这在Web环境中不允许）

            Office.context.document.settings.refreshAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const savedConfig = Office.context.document.settings.get('aiConfigData');
                    if (savedConfig) {
                        console.log("从文档设置中读取到配置");
                        resolve(JSON.parse(savedConfig));
                    } else {
                        console.log("文档中无保存的配置");
                        resolve(null);
                    }
                } else {
                    console.warn("读取设置失败:", asyncResult.error.message);
                    resolve(null);
                }
            });
        } catch (error) {
            console.error("读取本地配置时出错:", error);
            resolve(null); // 出错时返回null，而不是reject
        }
    });
}

// 从远程API获取配置
async function fetchRemoteConfig() {
    console.log("从远程API获取配置...");
    try {
        // 添加一个随机参数防止缓存
        const timestamp = new Date().getTime();
        const response = await fetch(`https://aiwechat-vercel-virid.vercel.app/api/check_api?t=${timestamp}`);

        if (!response.ok) {
            throw new Error(`API请求失败: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();

        if (!data.apiKey) {
            throw new Error("API返回的数据中没有apiKey");
        }

        // 解密API密钥 (如果需要)
        const decryptedKey = decryptApiKey(data.apiKey);

        // 创建并保存配置
        const config = {
            endpoint: data.endpoint || "https://models.inference.ai.azure.com/",
            apiKey: decryptedKey,
            modelName: data.modelName || "GPT-4o",
            timestamp: new Date().toISOString()
        };

        // 保存配置到文档设置中
        saveConfigToDocument(config);

        return config;
    } catch (error) {
        console.error("获取远程配置失败:", error);
        throw error;
    }
}

// 解密API密钥
// 解密API密钥 - 修改为匹配Go端的简单Base64编码
function decryptApiKey(encryptedKey) {
    try {
        // Go端只是简单的Base64编码，所以这里只需要Base64解码
        return atob(encryptedKey);
    } catch (error) {
        console.error("解密API密钥失败:", error);
        return encryptedKey; // 如果解码失败，返回原始密钥
    }
}

// 保存配置到文档设置
function saveConfigToDocument(config) {
    try {
        Office.context.document.settings.set('aiConfigData', JSON.stringify(config));
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("配置已保存到文档设置");
            } else {
                console.error("保存配置失败:", asyncResult.error.message);
            }
        });
    } catch (error) {
        console.error("保存配置时出错:", error);
    }
}