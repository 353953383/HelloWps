/**
 * 标准AI API接口 - 严格遵循AIapi.txt规范
 * 实现OpenAI兼容格式的API调用
 */

var AIapiStandard = (function() {
    'use strict';
    
    function AIapiStandard(config) {
        // 验证必要配置
        if (!config || (!config.apiKey && !config.baseURL && !config.apiEndpoint)) {
            throw new Error('缺少必要的API配置参数');
        }
        
        this.apiKey = config.apiKey;
        this.baseURL = config.baseURL;
        this.apiEndpoint = config.apiEndpoint;
        this.model = config.modelName || 'qwen-plus';
        
        // 确定API端点
        if (this.apiEndpoint) {
            // 局域网配置直接使用apiEndpoint
            this.endpoint = this.apiEndpoint;
        } else if (this.baseURL) {
            // 标准配置使用baseURL + 路径
            this.endpoint = this.baseURL + (this.baseURL.endsWith('/') ? '' : '/') + "chat/completions";
        } else {
            throw new Error('缺少API端点配置');
        }
        
        this.defaultHeaders = {
            'Authorization': 'Bearer ' + this.apiKey,
            'Content-Type': 'application/json'
        };
        
        // 验证API配置
        this.validateConfig();
    }
    
    /**
     * 验证API配置
     */
    AIapiStandard.prototype.validateConfig = function() {
        if (!this.apiKey) {
            throw new Error('API密钥未配置');
        }
        
        if (!this.endpoint) {
            throw new Error('API端点未配置');
        }
    };
    
    /**
     * 发起聊天请求 - 标准OpenAI格式
     */
    AIapiStandard.prototype.chat = function(messagesOrPrompt, options) {
        var self = this;
        
        return new Promise(function(resolve, reject) {
            try {
                // 格式化消息
                var messages = self.formatMessages(messagesOrPrompt);
                
                // 构建请求体
                var requestBody = {
                    model: self.model,
                    messages: messages,
                    max_tokens: 30000,
                    temperature: 0.3
                };
                
                // 合并选项
                if (options) {
                    for (var key in options) {
                        if (options.hasOwnProperty(key)) {
                            requestBody[key] = options[key];
                        }
                    }
                }
                
                // 创建请求选项
                var requestOptions = {
                    method: 'POST',
                    headers: self.defaultHeaders,
                    body: JSON.stringify(requestBody)
                };
                
                // 发起请求
                fetch(self.endpoint, requestOptions)
                    .then(function(response) {
                        if (!response.ok) {
                            throw new Error('HTTP错误: ' + response.status + ' ' + response.statusText);
                        }
                        return response.json();
                    })
                    .then(function(data) {
                        resolve(data);
                    })
                    .catch(function(error) {
                        reject(error);
                    });
                    
            } catch (error) {
                reject(error);
            }
        });
    };
    
    /**
     * 格式化消息
     */
    AIapiStandard.prototype.formatMessages = function(messagesOrPrompt) {
        if (typeof messagesOrPrompt === 'string') {
            return [
                {
                    role: 'user',
                    content: messagesOrPrompt
                }
            ];
        }
        
        if (Array.isArray(messagesOrPrompt)) {
            return messagesOrPrompt;
        }
        
        throw new Error('消息格式无效');
    };
    
    return AIapiStandard;
})();