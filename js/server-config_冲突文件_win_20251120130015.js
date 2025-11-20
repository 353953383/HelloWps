// 服务器配置文件 - 公共配置
// 直接设置全局变量，与util.js保持兼容

(function() {
    'use strict';
    
    // 设置全局配置变量
    function setServerConfig() {
        // 默认配置
        const config = {
            PRODUCTION: 'http://192.168.70.26:8080/V6R343/',
            //本地：http://192.168.70.17:8888/V6/
            AI_CONFIG_TYPE:'AI_CONFIG', // 可选: 'AI_CONFIG'(阿里云百炼), 'AI_CONFIG_WLAN'(局域网部署)
            // AI模型配置 - 阿里云百炼 (OpenAI格式配置)
            AI_CONFIG: {
                // 阿里云百炼OpenAI格式配置
                apiKey: 'sk-9bacbdffb7dd4b91b240c472d9c5e0c2', // 实际API密钥，如果使用环境变量请用：process.env.DASHSCOPE_API_KEY
                baseURL: 'https://dashscope.aliyuncs.com/compatible-mode/v1', // 新加坡地域请替换为：https://dashscope-intl.aliyuncs.com/compatible-mode/v1
                modelName: 'qwen-max', // 可选模型：qwen-plus、qwen-max、qwen3-max等
                maxTokens: 30000,
                temperature: 0.3,
                timeout: 60000, // 60秒超时
                description: '阿里云百炼OpenAI格式API（通过OpenAI客户端调用）'
            },
            
            AI_CONFIG_WLAN: {
                // 局域网 - 获取模型
                apiEndpoint: 'http://192.168.70.26:1234/v1/chat/completions', // 修改为实际的聊天完成端点
                modelName: 'local-model', // 局域网模型名称
                apiKey: 'local-api-key', // 局域网API密钥（如果需要）
                maxTokens: 30000,
                temperature: 0.3,
                timeout: 300000, // 300秒超时
                requestFormat: 'compatible', // 指定使用兼容模式
                description: '局域网部署的AI模型'
            }
        };
        
        // 设置到全局作用域
        window.SERVER_CONFIG = 'PRODUCTION';
        window.PRODUCTION = config.PRODUCTION;
        window.AI_CONFIG_TYPE = config.AI_CONFIG_TYPE;
        window.AI_CONFIG = config.AI_CONFIG;
        window.AI_CONFIG_WLAN = config.AI_CONFIG_WLAN;
        
        // 根据配置类型设置当前使用的AI配置
        if (config.AI_CONFIG_TYPE === 'AI_CONFIG_WLAN' && config.AI_CONFIG_WLAN) {
            window.CURRENT_AI_CONFIG = config.AI_CONFIG_WLAN;
        } else {
            window.CURRENT_AI_CONFIG = config.AI_CONFIG;
        }
        
        // 触发配置就绪事件
        if (typeof CustomEvent !== 'undefined') {
            const event = new CustomEvent('serverConfigReady', { detail: config });
            window.dispatchEvent(event);
        }
    }
    
    // 立即设置配置（确保在util.js加载之前）
    setServerConfig();
    
    // 页面加载完成后再设置一次（双重保险）
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', setServerConfig);
    }
    
})();