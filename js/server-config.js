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
            AI_CONFIG_TYPE:'AI_CONFIG',
            // AI模型配置 - 阿里云DashScope
            AI_CONFIG: {
                // 直连DashScope API（兼容模式）
                apiEndpoint: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
                apiKey: 'sk-9bacbdffb7dd4b91b240c472d9c5e0c2', // 实际API密钥
                modelName: 'qwen-plus', // 使用兼容模式支持的模型
                maxTokens: 30000,
                temperature: 0.3,
                timeout: 60000, // 60秒超时
                requestFormat: 'compatible' // 指定使用兼容模式
            },
            AI_CONFIG_WLAN: {
                // 局域网 - 获取模型
                apiEndpoint: 'http://192.168.70.26:1234/v1/models',
                maxTokens: 30000,
                temperature: 0.3,
                timeout: 300000, // 300秒超时
                requestFormat: 'compatible' // 指定使用兼容模式
            }
        };
        
        // 设置到全局作用域
        window.SERVER_CONFIG = 'PRODUCTION';
        window.PRODUCTION = config.PRODUCTION;
        window.AI_CONFIG = config.AI_CONFIG;
        
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