/**
 * 配置验证工具
 * 用于验证AI配置是否正确加载
 */

function validateAIConfig() {
    console.log('🔍 开始验证AI配置...');
    
    // 检查全局配置是否已加载
    console.log('📋 检查全局配置变量:');
    console.log('  - window.CONFIG:', window.CONFIG ? '✅ 已加载' : '❌ 未加载');
    console.log('  - window.AI_CONFIG_TYPE:', window.AI_CONFIG_TYPE || '❌ 未定义');
    console.log('  - window.AI_CONFIG:', window.AI_CONFIG ? '✅ 已加载' : '❌ 未加载');
    console.log('  - window.AI_CONFIG_WLAN:', window.AI_CONFIG_WLAN ? '✅ 已加载' : '❌ 未加载');
    console.log('  - window.CURRENT_AI_CONFIG:', window.CURRENT_AI_CONFIG ? '✅ 已加载' : '❌ 未加载');
    
    if (window.CURRENT_AI_CONFIG) {
        console.log('🎯 当前使用的AI配置:');
        console.log('  - apiKey:', window.CURRENT_AI_CONFIG.apiKey ? 
                   window.CURRENT_AI_CONFIG.apiKey.substring(0, 4) + '...' : '❌ 未设置');
        console.log('  - endpoint:', window.CURRENT_AI_CONFIG.apiEndpoint || 
                   window.CURRENT_AI_CONFIG.baseURL || '❌ 未设置');
        console.log('  - model:', window.CURRENT_AI_CONFIG.modelName || '❌ 未设置');
        console.log('  - type:', window.CURRENT_AI_CONFIG.apiEndpoint ? '局域网配置' : 
                   window.CURRENT_AI_CONFIG.baseURL ? '云端配置' : '❌ 未设置');
    }
    
    // 检查配置类型
    if (window.AI_CONFIG_TYPE === 'AI_CONFIG_WLAN') {
        console.log('✅ 配置类型为 AI_CONFIG_WLAN，应使用局域网配置');
        if (window.CURRENT_AI_CONFIG === window.AI_CONFIG_WLAN) {
            console.log('✅ 正确使用局域网配置');
            console.log('   - API端点:', window.CURRENT_AI_CONFIG.apiEndpoint);
            console.log('   - 模型名称:', window.CURRENT_AI_CONFIG.modelName);
            console.log('   - API密钥:', window.CURRENT_AI_CONFIG.apiKey.substring(0, 4) + '...');
        } else {
            console.log('❌ 配置类型为局域网，但未正确使用局域网配置');
            console.log('   - 预期使用:', 'AI_CONFIG_WLAN');
            console.log('   - 实际使用:', window.CURRENT_AI_CONFIG === window.AI_CONFIG ? 'AI_CONFIG' : '未知');
        }
    } else if (window.AI_CONFIG_TYPE === 'AI_CONFIG') {
        console.log('✅ 配置类型为 AI_CONFIG，应使用云端配置');
        if (window.CURRENT_AI_CONFIG === window.AI_CONFIG) {
            console.log('✅ 正确使用云端配置');
            console.log('   - 基础URL:', window.CURRENT_AI_CONFIG.baseURL);
            console.log('   - 模型名称:', window.CURRENT_AI_CONFIG.modelName);
            console.log('   - API密钥:', window.CURRENT_AI_CONFIG.apiKey.substring(0, 4) + '...');
        } else {
            console.log('❌ 配置类型为云端，但未正确使用云端配置');
            console.log('   - 预期使用:', 'AI_CONFIG');
            console.log('   - 实际使用:', window.CURRENT_AI_CONFIG === window.AI_CONFIG_WLAN ? 'AI_CONFIG_WLAN' : '未知');
        }
    } else {
        console.log('❌ 未知的配置类型:', window.AI_CONFIG_TYPE);
    }
    
    console.log('✅ 配置验证完成');
}

// 检查配置加载状态的函数
function checkConfigStatus() {
    if (window.CONFIG) {
        console.log('✅ 配置已加载');
        validateAIConfig();
    } else {
        console.log('❌ 配置尚未加载，正在尝试加载...');
        // 如果配置未加载，尝试重新加载
        if (typeof loadServerConfig === 'function') {
            loadServerConfig().then(function(config) {
                console.log('✅ 配置重新加载成功:', config);
                validateAIConfig();
            }).catch(function(error) {
                console.error('❌ 配置重新加载失败:', error);
            });
        } else {
            console.log('❌ loadServerConfig函数未定义');
        }
    }
}

// 自动执行验证（如果在浏览器环境中）
if (typeof window !== 'undefined') {
    // 等待配置完全加载后执行验证
    setTimeout(checkConfigStatus, 1000);
    setTimeout(checkConfigStatus, 3000); // 再次检查以确保配置加载完成
}