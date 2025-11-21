/**
 * API配置管理器
 * 负责加载、验证和管理AI API配置
 */

var ApiConfigManager = (function() {
    'use strict';
    
    function ApiConfigManager() {
        this.config = null;
        this.isValid = false;
        this.lastCheck = null;
        this.features = {
            supportsStreaming: false,
            supportsFunctions: false,
            supportsImages: false
        };
        
        this.init();
    }
    
    ApiConfigManager.prototype.init = function() {
        try {
            this.loadConfig();
            this.validateConfig();
            this.detectFeatures();
            
            this.lastCheck = new Date();
        } catch (error) {
            console.error('❌ API配置初始化失败:', error.message);
        }
    };
    
    /**
     * 加载配置
     */
    ApiConfigManager.prototype.loadConfig = function() {
        try {
            // 首先尝试从全局配置加载，优先使用 CURRENT_AI_CONFIG（如果已设置）
            if (window.CURRENT_AI_CONFIG) {
                this.config = window.CURRENT_AI_CONFIG;
                return;
            }
            
            if (window.AI_CONFIG) {
                this.config = window.AI_CONFIG;
                return;
            }
            
            // 尝试从localStorage加载
            var savedConfig = this.loadSavedConfig();
            if (savedConfig) {
                this.config = savedConfig;
                return;
            }
            
            // 使用默认配置
            this.config = {
                apiKey: '',
                baseURL: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
                modelName: 'qwen-plus',
                maxTokens: 2000,
                temperature: 0.7,
                timeout: 30000
            };
            
        } catch (error) {
            console.error('❌ 配置加载失败:', error.message);
            throw error;
        }
    };
    
    /**
     * 从本地存储加载配置
     */
    ApiConfigManager.prototype.loadSavedConfig = function() {
        try {
            var saved = localStorage.getItem('aiHelper_apiConfig');
            if (saved) {
                var config = JSON.parse(saved);
                return config;
            }
        } catch (error) {
            // 忽略解析错误
        }
        return null;
    };
    
    /**
     * 验证配置
     */
    ApiConfigManager.prototype.validateConfig = function() {
        if (!this.config) {
            this.isValid = false;
            throw new Error('配置未加载');
        }
        
        if (!this.config.apiKey) {
            this.isValid = false;
            throw new Error('缺少API密钥');
        }
        
        // 检查是否是局域网配置（使用apiEndpoint）或标准配置（使用baseURL）
        if (!this.config.baseURL && !this.config.apiEndpoint) {
            this.isValid = false;
            throw new Error('缺少基础URL或API端点');
        }
        
        // 验证URL格式（如果提供了baseURL）
        if (this.config.baseURL) {
            try {
                new URL(this.config.baseURL);
            } catch (error) {
                this.isValid = false;
                throw new Error('基础URL格式无效');
            }
        }
        
        // 验证API端点格式（如果提供了apiEndpoint）
        if (this.config.apiEndpoint) {
            try {
                new URL(this.config.apiEndpoint);
            } catch (error) {
                this.isValid = false;
                throw new Error('API端点格式无效');
            }
        }
        
        this.isValid = true;
    };
    
    /**
     * 检测API功能支持
     */
    ApiConfigManager.prototype.detectFeatures = function() {
        if (!this.isValid) return;
        
        // 根据模型和端点推断功能支持
        var baseURL = this.config.baseURL ? this.config.baseURL.toLowerCase() : '';
        var apiEndpoint = this.config.apiEndpoint ? this.config.apiEndpoint.toLowerCase() : '';
        var modelName = this.config.modelName ? this.config.modelName.toLowerCase() : '';
        
        // 重置功能支持
        this.features = {
            supportsStreaming: false,
            supportsFunctions: false,
            supportsImages: false
        };
        
        // DashScope支持
        if (baseURL.includes('dashscope')) {
            this.features.supportsStreaming = true;
            this.features.supportsFunctions = true;
        }
        
        // OpenAI支持
        if (baseURL.includes('openai') || baseURL.includes('api.com')) {
            this.features.supportsStreaming = true;
            this.features.supportsFunctions = true;
            this.features.supportsImages = true;
        }
        
        // 局域网模型默认功能支持（保守估计）
        if (apiEndpoint && !baseURL) {
            this.features.supportsStreaming = false; // 局域网模型通常不支持流式传输
            this.features.supportsFunctions = true;  // 假设支持函数调用
            this.features.supportsImages = false;    // 局域网模型通常不支持图像处理
        }
    };
    
    /**
     * 更新配置
     */
    ApiConfigManager.prototype.updateConfig = function(newConfig) {
        try {
            // 合并配置
            this.config = Object.assign({}, this.config, newConfig);
            
            // 验证新配置
            this.validateConfig();
            
            // 重新检测功能
            this.detectFeatures();
            
            // 保存配置
            this.saveConfig();
            
            this.lastCheck = new Date();
            
        } catch (error) {
            console.error('❌ 配置更新失败:', error.message);
            throw error;
        }
    };
    
    /**
     * 保存配置到本地存储
     */
    ApiConfigManager.prototype.saveConfig = function() {
        try {
            localStorage.setItem('aiHelper_apiConfig', JSON.stringify(this.config));
        } catch (error) {
            // 忽略存储错误
        }
    };
    
    /**
     * 获取配置状态
     */
    ApiConfigManager.prototype.getConfigStatus = function() {
        return {
            isConfigured: this.isValid,
            lastCheck: this.lastCheck,
            config: this.config,
            features: this.features
        };
    };
    
    /**
     * 测试API连接
     */
    ApiConfigManager.prototype.testConnection = function() {
        var self = this;
        return new Promise(function(resolve, reject) {
            if (!self.isValid) {
                reject(new Error('配置无效'));
                return;
            }
            
            // 创建一个简单的测试请求
            var testMessages = [
                {
                    role: 'user',
                    content: '你好'
                }
            ];
            
            var requestBody = {
                model: self.config.modelName,
                messages: testMessages,
                max_tokens: 10
            };
            
            // 确定API端点
            var url;
            if (self.config.apiEndpoint) {
                // 局域网配置使用apiEndpoint
                url = self.config.apiEndpoint;
            } else if (self.config.baseURL) {
                // 标准配置使用baseURL + 路径
                url = self.config.baseURL + (self.config.baseURL.endsWith('/') ? 'chat/completions' : '/chat/completions');
            } else {
                reject(new Error('缺少API端点配置'));
                return;
            }
            
            fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': 'Bearer ' + self.config.apiKey,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestBody),
                timeout: self.config.timeout || 10000
            })
            .then(function(response) {
                if (response.ok) {
                    resolve({
                        success: true,
                        status: response.status
                    });
                } else {
                    reject(new Error('连接测试失败: ' + response.status));
                }
            })
            .catch(function(error) {
                reject(new Error('网络错误: ' + error.message));
            });
        });
    };
    
    return ApiConfigManager;
})();

// 初始化配置管理器
window.apiConfigManager = new ApiConfigManager();

// 导出配置管理器的状态
if (window.apiConfigManager.isValid) {
    console.log('✅ API配置初始化成功');
} else {
    console.log('❌ API配置初始化失败');
}