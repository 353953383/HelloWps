/**
 * 服务器配置加载器
 * 用于加载server-config.txt中的配置信息
 */

async function loadServerConfig() {
    try {
        // 读取server-config.txt文件内容
        const response = await fetch('../../server-config.txt');
        if (!response.ok) {
            throw new Error(`配置文件加载失败: ${response.status} ${response.statusText}`);
        }
        
        const configText = await response.text();
        
        // 解析配置文本
        parseConfigText(configText);
        
        console.log('✅ 服务器配置加载成功');
        return window.CONFIG;
    } catch (error) {
        console.error('❌ 服务器配置加载失败:', error);
        throw error;
    }
}

function parseConfigText(configText) {
    // 创建一个临时的script元素来解析配置
    const lines = configText.split('\n');
    const config = {};
    
    lines.forEach(line => {
        line = line.trim();
        if (line && line.indexOf('=') > 0) {
            const parts = line.split('=');
            if (parts.length >= 2) {
                const key = parts[0].trim();
                let value = parts.slice(1).join('=').trim(); // 处理值中包含=的情况
                
                // 尝试解析为JSON或普通字符串
                if (value.startsWith('"') && value.endsWith('"')) {
                    value = value.substring(1, value.length - 1); // 去掉引号
                } else if (value.toLowerCase() === 'true') {
                    value = true;
                } else if (value.toLowerCase() === 'false') {
                    value = false;
                } else if (!isNaN(value) && value.trim() !== '') {
                    value = Number(value);
                } else if (value.includes('{') || value.includes('[')) {
                    try {
                        value = JSON.parse(value);
                    } catch (e) {
                        // 如果不是有效的JSON，保持为字符串
                    }
                }
                
                config[key] = value;
                
                // 特殊处理AI配置
                if (key.startsWith('AI_CONFIG')) {
                    // 处理AI配置对象
                    setAIConfig(key, value);
                }
            }
        }
    });
    
    window.CONFIG = config;
    
    // 从配置文本解析各个AI配置
    // 解析AI_CONFIG
    if (config['AI_CONFIG.apiKey']) {
        if (!window.AI_CONFIG) window.AI_CONFIG = {};
        window.AI_CONFIG.apiKey = config['AI_CONFIG.apiKey'];
    }
    if (config['AI_CONFIG.baseURL']) {
        if (!window.AI_CONFIG) window.AI_CONFIG = {};
        window.AI_CONFIG.baseURL = config['AI_CONFIG.baseURL'];
    }
    if (config['AI_CONFIG.modelName']) {
        if (!window.AI_CONFIG) window.AI_CONFIG = {};
        window.AI_CONFIG.modelName = config['AI_CONFIG.modelName'];
    }
    if (config['AI_CONFIG.maxTokens']) {
        if (!window.AI_CONFIG) window.AI_CONFIG = {};
        window.AI_CONFIG.maxTokens = config['AI_CONFIG.maxTokens'];
    }
    if (config['AI_CONFIG.temperature']) {
        if (!window.AI_CONFIG) window.AI_CONFIG = {};
        window.AI_CONFIG.temperature = config['AI_CONFIG.temperature'];
    }
    if (config['AI_CONFIG.timeout']) {
        if (!window.AI_CONFIG) window.AI_CONFIG = {};
        window.AI_CONFIG.timeout = config['AI_CONFIG.timeout'];
    }
    
    // 解析AI_CONFIG_WLAN
    if (config['AI_CONFIG_WLAN.apiEndpoint']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.apiEndpoint = config['AI_CONFIG_WLAN.apiEndpoint'];
    }
    if (config['AI_CONFIG_WLAN.modelName']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.modelName = config['AI_CONFIG_WLAN.modelName'];
    }
    if (config['AI_CONFIG_WLAN.apiKey']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.apiKey = config['AI_CONFIG_WLAN.apiKey'];
    }
    if (config['AI_CONFIG_WLAN.maxTokens']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.maxTokens = config['AI_CONFIG_WLAN.maxTokens'];
    }
    if (config['AI_CONFIG_WLAN.temperature']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.temperature = config['AI_CONFIG_WLAN.temperature'];
    }
    if (config['AI_CONFIG_WLAN.timeout']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.timeout = config['AI_CONFIG_WLAN.timeout'];
    }
    if (config['AI_CONFIG_WLAN.requestFormat']) {
        if (!window.AI_CONFIG_WLAN) window.AI_CONFIG_WLAN = {};
        window.AI_CONFIG_WLAN.requestFormat = config['AI_CONFIG_WLAN.requestFormat'];
    }
    
    // 根据SERVER_CONFIG和AI_CONFIG_TYPE设置CURRENT_AI_CONFIG
    // 优先使用AI_CONFIG_TYPE变量来确定使用哪个配置
    if (window.AI_CONFIG_TYPE) {
        if (window.AI_CONFIG_TYPE === 'AI_CONFIG_WLAN' && window.AI_CONFIG_WLAN) {
            window.CURRENT_AI_CONFIG = window.AI_CONFIG_WLAN;
            console.log('🎯 当前AI配置设置为局域网配置');
        } else if (window.AI_CONFIG_TYPE === 'AI_CONFIG' && window.AI_CONFIG) {
            window.CURRENT_AI_CONFIG = window.AI_CONFIG;
            console.log('🎯 当前AI配置设置为云端配置');
        } else if (window[window.AI_CONFIG_TYPE]) {
            window.CURRENT_AI_CONFIG = window[window.AI_CONFIG_TYPE];
            console.log('🎯 根据AI_CONFIG_TYPE设置当前AI配置:', window.AI_CONFIG_TYPE);
        }
    } else if (config.AI_CONFIG_TYPE === 'AI_CONFIG_WLAN' && window.AI_CONFIG_WLAN) {
        window.CURRENT_AI_CONFIG = window.AI_CONFIG_WLAN;
        console.log('🎯 当前AI配置设置为局域网配置');
    } else if (config.AI_CONFIG_TYPE === 'AI_CONFIG' && window.AI_CONFIG) {
        window.CURRENT_AI_CONFIG = window.AI_CONFIG;
        console.log('🎯 当前AI配置设置为云端配置');
    } else if (window.SERVER_CONFIG && window[window.SERVER_CONFIG]) {
        window.CURRENT_AI_CONFIG = window[window.SERVER_CONFIG];
        console.log('🎯 根据SERVER_CONFIG设置当前AI配置:', window.SERVER_CONFIG);
    }
    
    // 检测当前配置类型，不强制改变用户选择
    if ((window.wps || window.Application) && window.CURRENT_AI_CONFIG && 
        window.CURRENT_AI_CONFIG.apiKey && window.CURRENT_AI_CONFIG.apiKey.includes('sk-')) {
        console.log('🏠 WPS环境中使用云端AI配置，如需避免CORS问题可考虑使用局域网配置');
    }
    
    console.log('📋 解析后的配置:', window.CONFIG);
    console.log('🎯 当前AI配置:', window.CURRENT_AI_CONFIG);
}

function setAIConfig(key, value) {
    // 处理类似 AI_CONFIG.apiKey 这样的嵌套配置
    const parts = key.split('.');
    if (parts.length === 2) {
        const configName = parts[0]; // 如 AI_CONFIG
        const property = parts[1];   // 如 apiKey
        
        if (!window[configName]) {
            window[configName] = {};
        }
        
        window[configName][property] = value;
    }
}

// 自动加载配置（如果在浏览器环境中）
if (typeof window !== 'undefined') {
    // 延迟加载配置以确保其他脚本已准备就绪
    setTimeout(async () => {
        try {
            await loadServerConfig();
            
            // 触发配置加载完成事件
            const event = new Event('configLoaded');
            window.dispatchEvent(event);
        } catch (error) {
            console.error('自动加载配置失败:', error);
        }
    }, 500);
}