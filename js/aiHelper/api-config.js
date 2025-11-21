/**
 * API配置管理器
 * 解决API配置不完整和网络请求问题
 */

class APIConfigManager {
    constructor() {
        this.config = {
            // 主要配置 - 使用代理服务器地址（从server-config.js继承）
            endpoint: 'http://127.0.0.1:3889/api/chat/completions', // 正确的代理服务器地址
            apiKey: '',
            modelName: 'qwen-plus',
            
            // 高级配置
            timeout: 60000,
            maxRetries: 3,
            retryDelay: 1000,
            useProxy: true,
            
            // 网络配置
            enableCORS: false,
            requestMode: 'cors',
            credentials: 'include'
        };
        
        this.isConfigured = false;
        this.lastConfigCheck = null;
        
        // 加载保存的配置
        this.loadConfig();
        
        // 验证配置
        this.validateConfig();
    }
    
    /**
     * 加载配置
     */
    loadConfig() {
        try {
            const savedConfig = localStorage.getItem('ai-helper-config');
            if (savedConfig) {
                const parsed = JSON.parse(savedConfig);
                this.config = { ...this.config, ...parsed };
                console.log('✅ API配置已加载');
            } else {
                console.log('ℹ️ 未找到保存的配置，使用默认配置');
            }
        } catch (error) {
            console.warn('⚠️ 加载配置失败:', error);
        }
    }
    
    /**
     * 保存配置
     */
    saveConfig() {
        try {
            // 不保存敏感信息到本地存储
            const safeConfig = {
                ...this.config,
                apiKey: '' // 清空敏感信息
            };
            
            localStorage.setItem('ai-helper-config', JSON.stringify(safeConfig));
            console.log('✅ API配置已保存');
        } catch (error) {
            console.warn('⚠️ 保存配置失败:', error);
        }
    }
    
    /**
     * 验证配置
     */
    validateConfig() {
        const errors = [];
        
        // 检查端点配置
        if (!this.config.endpoint || this.co·nfig.endpoint.trim() === '') {
            errors.push('API端点地址不能为空');
        } else if (!this.isValidURL(this.config.endpoint)) {
            errors.push('API端点地址格式不正确');
        }
        
        // 检查模型名称
        if (!this.config.modelName || this.config.modelName.trim() === '') {
            errors.push('模型名称不能为空');
        }
        
        // 检查超时设置
        if (this.config.timeout < 5000 || this.config.timeout > 300000) {
            errors.push('超时设置应在5秒到5分钟之间');
        }
        
        // 检查重试设置
        if (this.config.maxRetries < 0 || this.config.maxRetries > 10) {
            errors.push('重试次数应在0到10次之间');
        }
        
        this.isConfigured = errors.length === 0;
        this.lastConfigCheck = new Date();
        
        if (errors.length > 0) {
            console.warn('⚠️ API配置验证失败:', errors);
        } else {
            console.log('✅ API配置验证通过');
        }
        
        return {
            isValid: this.isConfigured,
            errors: errors,
            timestamp: this.lastConfigCheck
        };
    }
    
    /**
     * 测试API连接
     */
    async testConnection() {
        try {
            console.log('🔍 开始测试API连接...');
            
            const testUrl = `${this.config.endpoint}/api/test`;
            console.log('📡 测试URL:', testUrl);
            
            const response = await fetch(testUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    test: true,
                    timestamp: new Date().toISOString()
                }),
                timeout: this.config.timeout
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const result = await response.json();
            
            if (result.success) {
                console.log('✅ API连接测试成功');
                return {
                    success: true,
                    message: 'API连接测试成功',
                    response: result
                };
            } else {
                throw new Error(result.error || '未知错误');
            }
            
        } catch (error) {
            console.error('❌ API连接测试失败:', error);
            
            return {
                success: false,
                error: error.message,
                suggestions: this.getErrorSuggestions(error)
            };
        }
    }
    
    /**
     * 获取错误建议
     */
    getErrorSuggestions(error) {
        const suggestions = [];
        
        if (error.message.includes('fetch')) {
            suggestions.push('请检查代理服务器是否已启动');
            suggestions.push('请检查网络连接是否正常');
        }
        
        if (error.message.includes('CORS')) {
            suggestions.push('请使用代理服务器访问API');
            suggestions.push('检查代理服务器是否正确处理CORS');
        }
        
        if (error.message.includes('401') || error.message.includes('403')) {
            suggestions.push('请检查API密钥是否正确');
            suggestions.push('请确认API密钥有调用权限');
        }
        
        if (error.message.includes('timeout')) {
            suggestions.push('请增加超时设置');
            suggestions.push('请检查网络延迟');
        }
        
        return suggestions;
    }
    
    /**
     * 更新配置
     */
    updateConfig(newConfig) {
        this.config = { ...this.config, ...newConfig };
        this.validateConfig();
        this.saveConfig();
        
        return {
            success: true,
            config: this.getSafeConfig(),
            validation: this.validateConfig()
        };
    }
    
    /**
     * 获取安全的配置（不包含敏感信息）
     */
    getSafeConfig() {
        return {
            ...this.config,
            apiKey: this.config.apiKey ? '***' : '',
            hasApiKey: !!this.config.apiKey
        };
    }
    
    /**
     * 智能配置检测
     */
    async detectOptimalConfig() {
        console.log('🔍 智能检测最优配置...');
        
        const tests = [
            {
                name: '代理服务器',
                endpoint: 'http://127.0.0.1:3889/api/chat/completions',
                useProxy: true
            },
            {
                name: '本地代理',
                endpoint: 'http://127.0.0.1:3889/api/chat/completions',
                useProxy: true
            },
            {
                name: '直连模式（不推荐，存在CORS问题）',
                endpoint: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
                useProxy: false
            }
        ];
        
        const results = [];
        
        for (const test of tests) {
            try {
                console.log(`🧪 测试${test.name}...`);
                
                const testUrl = test.useProxy 
                    ? `${test.endpoint}/api/health`
                    : test.endpoint;
                
                const response = await fetch(testUrl, {
                    method: 'GET',
                    timeout: 5000
                });
                
                const success = response.ok;
                
                results.push({
                    ...test,
                    success,
                    responseTime: Date.now(),
                    status: response.status
                });
                
                if (success) {
                    console.log(`✅ ${test.name}可用`);
                } else {
                    console.log(`❌ ${test.name}不可用`);
                }
                
            } catch (error) {
                console.log(`❌ ${test.name}失败:`, error.message);
                
                results.push({
                    ...test,
                    success: false,
                    error: error.message
                });
            }
        }
        
        // 选择最优配置
        const bestConfig = results.find(r => r.success) || tests[0];
        
        this.updateConfig({
            endpoint: bestConfig.endpoint,
            useProxy: bestConfig.useProxy
        });
        
        return {
            tested: results,
            recommended: bestConfig,
            current: this.getSafeConfig()
        };
    }
    
    /**
     * URL格式验证
     */
    isValidURL(string) {
        try {
            new URL(string);
            return true;
        } catch (_) {
            return false;
        }
    }
    
    /**
     * 获取配置状态
     */
    getStatus() {
        return {
            isConfigured: this.isConfigured,
            lastCheck: this.lastConfigCheck,
            config: this.getSafeConfig(),
            features: {
                proxySupported: this.config.useProxy,
                corsEnabled: this.config.enableCORS,
                timeout: this.config.timeout,
                maxRetries: this.config.maxRetries
            }
        };
    }
    
    /**
     * 重置配置
     */
    resetConfig() {
        this.config = {
            endpoint: 'http://127.0.0.1:3889/api/chat/completions',
            apiKey: '',
            modelName: 'qwen-plus',
            timeout: 60000,
            maxRetries: 3,
            retryDelay: 1000,
            useProxy: true,
            enableCORS: false,
            requestMode: 'cors',
            credentials: 'include'
        };
        
        localStorage.removeItem('ai-helper-config');
        this.validateConfig();
        
        console.log('🔄 API配置已重置');
        
        return {
            success: true,
            message: '配置已重置为默认值'
        };
    }
    
    /**
     * 导出配置
     */
    exportConfig() {
        const config = this.getSafeConfig();
        const timestamp = new Date().toISOString();
        
        return {
            version: '1.0',
            timestamp: timestamp,
            config: config,
            instructions: {
                '设置API密钥': '请在代理服务器中设置DASHSCOPE_API_KEY环境变量',
                '启动代理服务器': '运行 server/start-proxy.sh',
                '测试连接': '访问 http://localhost:8080/api/health',
                'Chat接口地址': 'http://localhost:8080/api/chat/completions'
            }
        };
    }
}

// 创建全局实例
window.apiConfigManager = new APIConfigManager();

// 导出配置管理器的状态
console.log('🔧 API配置管理器已初始化');
console.log('📋 配置状态:', window.apiConfigManager.getStatus());