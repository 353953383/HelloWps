//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft:0,
    msoCTPDockPositionRight:2
}

//从配置文件读取服务器配置 - 统一读取index.html配置
function loadServerConfig() {
    return new Promise((resolve, reject) => {
        try {
            const xhr = new XMLHttpRequest();
            xhr.open('GET', '/server-config.txt', true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4) {
                    if (xhr.status === 200) {
                        try {
                            const configText = xhr.responseText;
                            console.log('📄 配置文件内容:', configText);
                            
                            if (configText) {
                                var config = {};
                                
                                // 解析配置文件
                                var lines = configText.split('\n');
                                for (var j = 0; j < lines.length; j++) {
                                    var line = lines[j].trim();
                                    if (line && line.indexOf('=') > 0) {
                                        var parts = line.split('=');
                                        // 处理嵌套对象属性 (如 AI_CONFIG.apiKey)
                                        if (parts[0].includes('.')) {
                                            const keyParts = parts[0].split('.');
                                            let obj = config;
                                            for (let i = 0; i < keyParts.length - 1; i++) {
                                                if (!obj[keyParts[i]]) {
                                                    obj[keyParts[i]] = {};
                                                }
                                                obj = obj[keyParts[i]];
                                            }
                                            // 特殊处理数字和布尔值
                                            const trimmedValue = parts[1].trim();
                                            if (!isNaN(trimmedValue) && trimmedValue !== '') {
                                                obj[keyParts[keyParts.length - 1]] = Number(trimmedValue);
                                            } else if (trimmedValue === 'true' || trimmedValue === 'false') {
                                                obj[keyParts[keyParts.length - 1]] = trimmedValue === 'true';
                                            } else {
                                                obj[keyParts[keyParts.length - 1]] = trimmedValue;
                                            }
                                        } else {
                                            // 特殊处理数字和布尔值
                                            const trimmedValue = parts[1].trim();
                                            if (!isNaN(trimmedValue) && trimmedValue !== '') {
                                                config[parts[0]] = Number(trimmedValue);
                                            } else if (trimmedValue === 'true' || trimmedValue === 'false') {
                                                config[parts[0]] = trimmedValue === 'true';
                                            } else {
                                                config[parts[0]] = trimmedValue;
                                            }
                                        }
                                    }
                                }
                                
                                console.log('⚙️ 解析后的配置对象:', config);
                                
                                // 验证主服务器配置
                                if (config.PRODUCTION) {
                                    resolve(config);
                                } else {
                                    reject(new Error('配置文件中缺少PRODUCTION配置'));
                                }
                            } else {
                                reject(new Error('配置文件为空'));
                            }
                        } catch (parseError) {
                            reject(parseError);
                        }
                    } else {
                        reject(new Error(`无法加载配置文件，HTTP状态码: ${xhr.status}`));
                    }
                }
            };
            xhr.onerror = function() {
                reject(new Error('网络错误，无法加载配置文件'));
            };
            xhr.send();
        } catch (e) {
            // 配置读取失败，记录错误
            if (typeof console !== 'undefined' && console.error) {
                console.error('❌ 严重错误：无法加载外部配置文件，请检查 server-config.txt 是否存在且格式正确:', e);
            }
            reject(new Error('配置加载失败，系统无法启动'));
        }
    });
}

// 统一配置管理
(function() {
    try {
        // 加载配置
        var configResult = null;
        try {
            loadServerConfig().then(function(result) {
                configResult = result;
                if (configResult) {
                    // 将配置暴露到全局作用域
                    window.CONFIG = configResult;
                    window.SERVER_CONFIG = configResult.PRODUCTION;
                    window.PRODUCTION = configResult.PRODUCTION;
                    window.AI_CONFIG_TYPE = configResult.AI_CONFIG_TYPE;
                    window.AI_CONFIG = configResult.AI_CONFIG;
                    window.AI_CONFIG_WLAN = configResult.AI_CONFIG_WLAN;
                    
                    // 根据配置类型设置当前使用的AI配置
                    if (configResult.AI_CONFIG_TYPE === 'AI_CONFIG_WLAN' && configResult.AI_CONFIG_WLAN) {
                        window.CURRENT_AI_CONFIG = configResult.AI_CONFIG_WLAN;
                    } else {
                        window.CURRENT_AI_CONFIG = configResult.AI_CONFIG;
                    }
                    
                    console.log('✅ 配置加载成功:', configResult);
                } else {
                    throw new Error('配置加载返回空结果');
                }
            }).catch(function(error) {
                console.error('配置加载失败:', error);
                applyFallbackConfig();
            });
        } catch (loadError) {
            console.error('配置加载失败:', loadError);
            applyFallbackConfig();
        }
        
    } catch (e) {
        console.error('配置初始化失败:', e);
        applyFallbackConfig();
    }
    
    function applyFallbackConfig() {
        // 添加应急配置以防系统无法启动
        window.PRODUCTION = 'http://192.168.70.26:8080/V6R343/';
        window.SERVER_CONFIG = 'http://192.168.70.26:8080/V6R343/';
        window.AI_CONFIG = {
            apiKey: 'sk-9bacbdffb7dd4b91b240c472d9c5e0c2',
            baseURL: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
            modelName: 'qwen-max',
            maxTokens: 30000,
            temperature: 0.3,
            timeout: 60000,
            description: '阿里云百炼OpenAI格式API（通过OpenAI客户端调用）'
        };
        window.AI_CONFIG_WLAN = {
            apiEndpoint: 'http://192.168.70.26:1234/v1/chat/completions',
            modelName: 'local-model',
            apiKey: 'local-api-key',
            maxTokens: 30000,
            temperature: 0.3,
            timeout: 300000,
            requestFormat: 'compatible',
            description: '局域网部署的AI模型'
        };
        window.CURRENT_AI_CONFIG = window.AI_CONFIG;
        console.warn('⚠️ 配置初始化失败，使用应急配置');
    }
})();

/**
 * 获取主服务器地址
 * @returns {Promise<string>} 主服务器地址
 */
function selectServer() {
    return new Promise((resolve, reject) => {
        // 确保配置已加载
        if (window.PRODUCTION) {
            resolve(window.PRODUCTION);
        } else {
            // 如果配置尚未加载，等待一段时间后重试
            setTimeout(() => {
                if (window.PRODUCTION) {
                    resolve(window.PRODUCTION);
                } else {
                    reject(new Error('无法获取服务器配置，请检查配置文件'));
                }
            }, 1000);
        }
    });
}

function GetUrlPath() {
    let e = document.location.toString()
    return -1!=(e=decodeURI(e)).indexOf("/")&&(e=e.substring(0,e.lastIndexOf("/"))),e
}

/**
 * 通过wps提供的接口执行一段脚本
 * @param {*} param 需要执行的脚本
 */
function shellExecuteByOAAssist(param) {
    if (wps != null) {
        wps.OAAssist.ShellExecute(param)
    }
}