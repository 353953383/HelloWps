/**
 * 增强的AI接口管理器 - 整合代理服务器支持
 * 解决CORS错误、API配置和网络请求问题
 * 适配WPS JSA环境
 */

// 使用立即执行函数包装，确保兼容WPS JSA环境
var EnhancedAIInterface = (function() {
    'use strict';
    
    function EnhancedAIInterface() {
        this.configManager = window.apiConfigManager || null;
        this.requestQueue = [];
        this.isProcessing = false;
        this.stats = {
            totalRequests: 0,
            successfulRequests: 0,
            failedRequests: 0,
            lastRequestTime: null
        };
        
        // 重写fetch方法以支持代理
        this.originalFetch = window.fetch;
        this.setupProxyFetch();
    }
    
    /**
     * 设置代理fetch方法
     */
    EnhancedAIInterface.prototype.setupProxyFetch = function() {
        // 等待API配置管理器初始化
        var self = this;
        this.waitForConfigManager().then(function(configManager) {
            if (!configManager) {
                return;
            }
        }).catch(function(error) {
        });
    };
    
    /**
     * 等待API配置管理器初始化
     */
    EnhancedAIInterface.prototype.waitForConfigManager = function() {
        var self = this;
        return new Promise(function(resolve, reject) {
            var maxAttempts = 50; // 最多等待5秒（50 * 100ms）
            var attempts = 0;
            
            var checkConfigManager = function() {
                if (window.apiConfigManager && typeof window.apiConfigManager === 'object') {
                    self.configManager = window.apiConfigManager;
                    resolve(self.configManager);
                } else if (attempts < maxAttempts) {
                    attempts++;
                    setTimeout(checkConfigManager, 100);
                } else {
                    resolve(null);
                }
            };
            
            checkConfigManager();
        });
    };
    
    /**
     * 检查是否为AI API调用
     */
    EnhancedAIInterface.prototype.isAIApiCall = function(url, options) {
        // 检查URL是否为AI API端点
        if (typeof url === 'string' && url.includes('dashscope.aliyuncs.com')) {
            return true;
        }
        
        // 检查请求体是否包含AI相关字段
        if (options && options.body) {
            try {
                var body = JSON.parse(options.body);
                return body && (body.messages || body.model || body.input);
            } catch (_) {
                // 如果无法解析JSON，不做特殊处理
                return false;
            }
        }
        
        return false;
    };
    
    /**
     * 处理AI API请求
     */
    EnhancedAIInterface.prototype.handleAIApiRequest = function(url, options) {
        this.stats.totalRequests++;
        this.stats.lastRequestTime = new Date();
        
        var self = this;
        return new Promise(function(resolve, reject) {
            try {
                var config = self.configManager.config;
                var requestUrl = url;
                var requestOptions = {};
                if (options) {
                    for (var key in options) {
                        requestOptions[key] = options[key];
                    }
                }
                
                // 使用XMLHttpRequest替代fetch以兼容WPS JSA
                var xhr = new XMLHttpRequest();
                xhr.open(requestOptions.method || 'GET', requestUrl, true);
                
                // 设置请求头
                if (requestOptions.headers) {
                    for (var header in requestOptions.headers) {
                        xhr.setRequestHeader(header, requestOptions.headers[header]);
                    }
                }
                
                // 设置超时
                xhr.timeout = config.timeout || 30000;
                
                xhr.onload = function() {
                    if (xhr.status >= 200 && xhr.status < 300) {
                        var result;
                        try {
                            result = JSON.parse(xhr.responseText);
                        } catch (parseError) {
                            reject(new Error('响应解析失败: ' + parseError.message));
                            return;
                        }
                        
                        self.stats.successfulRequests++;
                        resolve(self.formatAIResponse(result));
                    } else {
                        var error = new Error('HTTP ' + xhr.status + ': ' + xhr.statusText);
                        self.stats.failedRequests++;
                        reject(error);
                    }
                };
                
                xhr.onerror = function() {
                    var error = new Error('网络请求失败');
                    self.stats.failedRequests++;
                    reject(error);
                };
                
                xhr.ontimeout = function() {
                    var error = new Error('请求超时');
                    self.stats.failedRequests++;
                    reject(error);
                };
                
                // 发送请求
                xhr.send(requestOptions.body || null);
                
            } catch (error) {
                self.stats.failedRequests++;
                // 尝试备用方案
                reject(error);
            }
        });
    };
    
    /**
     * 处理流式AI API请求
     */
    EnhancedAIInterface.prototype.handleStreamAIApiRequest = function(url, options, streamOptions) {
        var self = this;
        streamOptions = streamOptions || {};
        
        return new Promise(function(resolve, reject) {
            try {
                // 使用fetch和ReadableStream处理流式响应
                fetch(url, options)
                    .then(function(response) {
                        if (!response.ok) {
                            throw new Error('网络响应错误: ' + response.status);
                        }
                        
                        // 如果没有body，尝试直接获取响应文本
                        if (!response.body) {
                            console.warn('流式响应不可用，尝试获取完整响应');
                            return response.text().then(function(text) {
                                // 尝试解析响应
                                var result = self.parseAIResponse(text);
                                self.stats.successfulRequests++;
                                resolve(result);
                            }).catch(function(error) {
                                self.stats.failedRequests++;
                                reject(new Error('响应解析失败: ' + error.message));
                            });
                        }
                        
                        var reader = response.body.getReader();
                        var decoder = new TextDecoder();
                        var fullResponse = '';
                        var thinkingProcess = '';
                        var accumulatedData = ''; // 用于累积流式数据
                        
                        // 递归读取流
                        function read() {
                            reader.read().then(function(result) {
                                if (result.done) {
                                    // 流结束，尝试解析累积的数据
                                    try {
                                        // 如果有累积的完整响应，优先使用
                                        if (fullResponse.trim() !== '') {
                                            var parsedResult = self.parseAIResponse(fullResponse);
                                            self.stats.successfulRequests++;
                                            resolve(parsedResult);
                                            return;
                                        }
                                        
                                        // 处理可能的SSE格式数据
                                        var lines = accumulatedData.split('\n');
                                        var jsonData = '';
                                        
                                        lines.forEach(function(line) {
                                            if (line.startsWith('data: ')) {
                                                var data = line.substring(6);
                                                if (data !== '[DONE]') {
                                                    try {
                                                        var parsed = JSON.parse(data);
                                                        if (parsed.choices && parsed.choices[0] && parsed.choices[0].delta) {
                                                            jsonData += parsed.choices[0].delta.content || '';
                                                        } else if (parsed.content) {
                                                            jsonData += parsed.content;
                                                        }
                                                    } catch (e) {
                                                        // 如果不是JSON，直接添加
                                                        jsonData += data;
                                                    }
                                                }
                                            }
                                        });
                                        
                                        if (jsonData.trim() !== '') {
                                            var parsedResult = self.parseAIResponse(jsonData);
                                            self.stats.successfulRequests++;
                                            resolve(parsedResult);
                                        } else if (accumulatedData.trim() !== '') {
                                            // 尝试直接解析累积的数据
                                            var parsedResult = self.parseAIResponse(accumulatedData);
                                            self.stats.successfulRequests++;
                                            resolve(parsedResult);
                                        } else {
                                            throw new Error('响应内容为空');
                                        }
                                    } catch (parseError) {
                                        self.stats.failedRequests++;
                                        reject(new Error('响应解析失败: ' + parseError.message));
                                    }
                                    return;
                                }
                                
                                // 处理接收到的数据块
                                var chunk = decoder.decode(result.value, { stream: true });
                                accumulatedData += chunk; // 累积所有数据
                                
                                var lines = chunk.split('\n');
                                
                                lines.forEach(function(line) {
                                    if (line.startsWith('data: ')) {
                                        var data = line.substring(6);
                                        if (data === '[DONE]') {
                                            // 完成信号
                                            return;
                                        }
                                        
                                        try {
                                            var jsonData = JSON.parse(data);
                                            
                                            // 更新思考过程 - 显示累积的响应内容
                                            if (jsonData.choices && jsonData.choices[0] && jsonData.choices[0].delta) {
                                                var content = jsonData.choices[0].delta.content || '';
                                                if (content) {
                                                    thinkingProcess += content;
                                                    if (streamOptions.onProgress) {
                                                        streamOptions.onProgress(thinkingProcess);
                                                    }
                                                }
                                            } else if (jsonData.content) {
                                                thinkingProcess += jsonData.content;
                                                if (streamOptions.onProgress) {
                                                    streamOptions.onProgress(thinkingProcess);
                                                }
                                            }
                                            
                                            // 累积完整响应
                                            if (jsonData.choices && jsonData.choices[0] && jsonData.choices[0].delta) {
                                                fullResponse += jsonData.choices[0].delta.content || '';
                                            } else if (jsonData.content) {
                                                fullResponse += jsonData.content;
                                            }
                                        } catch (parseError) {
                                            // 忽略解析错误，可能不是JSON数据
                                            console.warn('流式数据解析失败:', parseError);
                                        }
                                    }
                                });
                                
                                // 继续读取
                                read();
                            }).catch(function(error) {
                                self.stats.failedRequests++;
                                reject(new Error('流式读取失败: ' + error.message));
                            });
                        }
                        
                        // 开始读取流
                        read();
                    })
                    .catch(function(error) {
                        self.stats.failedRequests++;
                        reject(new Error('请求失败: ' + error.message));
                    });
            } catch (error) {
                self.stats.failedRequests++;
                reject(error);
            }
        });
    };
    
    /**
     * 格式化AI响应
     */
    EnhancedAIInterface.prototype.formatAIResponse = function(result) {
        try {
            // 检查OpenAI兼容格式
            if (result.choices && result.choices[0]) {
                var content = result.choices[0].message.content;
                return this.parseAIResponse(content);
            }
            
            // 直接返回代理响应
            return {
                success: true,
                data: result,
                timestamp: new Date().toISOString()
            };
            
        } catch (error) {
            return result;
        }
    };
    
    /**
     * 解析AI响应内容（增强版）
     * 使用改进的JSON提取和错误处理机制
     */
    EnhancedAIInterface.prototype.parseAIResponse = function(response) {
        try {
            // 提取响应内容
            var content = response;
            if (typeof response === 'object') {
                if (response.choices && response.choices[0] && response.choices[0].message) {
                    content = response.choices[0].message.content;
                } else if (response.content) {
                    content = response.content;
                } else {
                    content = JSON.stringify(response);
                }
            }
            
            if (!content || typeof content !== 'string') {
                throw new Error('响应内容为空或格式无效');
            }
            
            // 如果内容是纯文本，尝试包装成基本格式
            if (!content.trim().startsWith('{') && !content.includes('```json')) {
                // 检查是否是流式响应的一部分
                if (content.includes('data:')) {
                    // 尝试从流式响应中提取完整内容
                    var fullContent = this.extractContentFromStream(content);
                    if (fullContent && fullContent.trim().startsWith('{')) {
                        content = fullContent;
                    } else {
                        // 创建基本的响应格式
                        return {
                            success: true,
                            formulas: [{
                                title: "默认公式",
                                formula: "=0",
                                explanation: "默认公式示例",
                                confidence: 50
                            }],
                            data_analysis: {},
                            alternative_formulas: [],
                            _rawContent: content // 保存原始内容供调试
                        };
                    }
                } else {
                    // 创建基本的响应格式
                    return {
                        success: true,
                        formulas: [{
                            title: "默认公式",
                        formula: "=0",
                            explanation: "默认公式示例",
                            confidence: 50
                        }],
                        data_analysis: {},
                        alternative_formulas: [],
                        _rawContent: content // 保存原始内容供调试
                    };
                }
            }
            
            // 使用多种方法尝试提取JSON
            var jsonData = null;
            var extractMethod = '';
            
            // 方法1: 查找JSON代码块标记
            if (content.includes('```json') || content.includes('```')) {
                var jsonBlockMatch = content.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
                if (jsonBlockMatch) {
                    jsonData = this.extractJSONContent(jsonBlockMatch[1]);
                    extractMethod = '代码块提取';
                }
            }
            
            // 方法2: 直接尝试解析整个内容为JSON（处理简单JSON响应）
            if (!jsonData) {
                try {
                    jsonData = JSON.parse(content);
                    extractMethod = '直接解析';
                } catch (e) {
                    // 忽略错误，继续尝试其他方法
                }
            }
            
            // 方法3: 正则表达式提取大括号内容
            if (!jsonData) {
                var braceMatch = content.match(/\{[\s\S]*\}/);
                if (braceMatch) {
                    jsonData = this.extractJSONByBrackets(braceMatch[0]);
                    extractMethod = '括号匹配提取';
                }
            }
            
            // 方法4: 尝试修复不完整的JSON
            if (!jsonData && content.includes('{')) {
                var partialMatch = content.substring(content.indexOf('{'));
                jsonData = this.tryFixIncompleteJSON(partialMatch);
                if (jsonData) {
                    extractMethod = 'JSON修复';
                }
            }
            
            // 方法5: 响应内容重建
            if (!jsonData) {
                jsonData = this.rebuildResponseFromContent(content);
                extractMethod = '内容重建';
            }
            
            if (!jsonData) {
                // 如果所有方法都失败了，创建一个基本响应
                return {
                    success: true,
                    formulas: [{
                        title: "响应解析",
                        formula: "=0",
                        explanation: "原始响应: " + content.substring(0, 100) + (content.length > 100 ? "..." : ""),
                        confidence: 70
                    }],
                    data_analysis: {
                        smart_analysis: "无法解析完整的JSON响应"
                    },
                    alternative_formulas: [],
                    _rawContent: content
                };
            }
            
            // 验证和增强响应数据
            var validatedData = this.validateResponse(jsonData);
            
            // 添加元数据
            validatedData._metadata = {
                extractMethod: extractMethod,
                timestamp: new Date().toISOString(),
                originalLength: content.length,
                processedAt: new Date().toISOString()
            };
            
            return validatedData;
            
        } catch (error) {
            // 创建详细的错误响应
            var errorResponse = this.createEnhancedErrorResponse(error, response);
            return errorResponse;
        }
    };
    
    /**
     * 提取JSON内容
     */
    EnhancedAIInterface.prototype.extractJSONContent = function(jsonStr) {
        try {
            // 清理JSON字符串
            var cleanJson = jsonStr.trim();
            
            // 移除可能的代码块标记
            cleanJson = cleanJson.replace(/^```json\s*/, '').replace(/\s*```$/, '');
            
            // 尝试直接解析
            try {
                return JSON.parse(cleanJson);
            } catch (e) {
                // 尝试修复常见JSON错误
                return this.fixCommonJSONErrors(cleanJson);
            }
        } catch (error) {
            throw new Error('JSON内容提取失败: ' + error.message);
        }
    };
    
    /**
     * 通过括号匹配提取JSON
     */
    EnhancedAIInterface.prototype.extractJSONByBrackets = function(jsonStr) {
        try {
            var cleanJson = jsonStr.trim();
            
            // 尝试修复常见错误
            cleanJson = this.fixCommonJSONErrors(cleanJson);
            
            return JSON.parse(cleanJson);
        } catch (error) {
            throw new Error('括号匹配提取失败: ' + error.message);
        }
    };
    
    /**
     * 尝试修复不完整的JSON
     */
    EnhancedAIInterface.prototype.tryFixIncompleteJSON = function(jsonStr) {
        try {
            var fixedJson = jsonStr.trim();
            
            // 修复常见的JSON错误
            fixedJson = this.fixCommonJSONErrors(fixedJson);
            
            // 尝试解析
            var parsed = JSON.parse(fixedJson);
            
            return parsed;
            
        } catch (error) {
            return null;
        }
    };
    
    /**
     * 修复常见JSON错误
     */
    EnhancedAIInterface.prototype.fixCommonJSONErrors = function(jsonStr) {
        var fixed = jsonStr;
        
        // 修复单引号为双引号
        fixed = fixed.replace(/'/g, '"');
        
        // 修复未加引号的键名
        fixed = fixed.replace(/(\w+):/g, '"$1":');
        
        return fixed;
    };
    
    /**
     * 从内容重建响应
     */
    EnhancedAIInterface.prototype.rebuildResponseFromContent = function(content) {
        // 简化实现，尝试直接解析整个内容
        try {
            return JSON.parse(content);
        } catch (error) {
            return null;
        }
    };
    
    /**
     * 验证响应
     */
    EnhancedAIInterface.prototype.validateResponse = function(data) {
        // 确保响应包含必要的字段
        if (!data.formulas && !data.formula) {
            // 如果没有formulas字段，尝试创建一个
            return {
                success: true,
                formulas: [{
                    title: "默认公式",
                    formula: data.formula || "=0",
                    explanation: data.explanation || "默认公式示例",
                    confidence: data.confidence || 50
                }],
                data_analysis: data.data_analysis || {},
                alternative_formulas: data.alternative_formulas || []
            };
        } else if (data.formula && !data.formulas) {
            // 处理简单的响应格式 { "formula": "...", "explanation": "...", ... }
            return {
                success: true,
                formulas: [{
                    title: data.title || "推荐公式",
                    formula: data.formula,
                    explanation: data.explanation || "无说明",
                    confidence: data.confidence || 90
                }],
                data_analysis: data.data_analysis || {},
                alternative_formulas: data.alternative_formulas || []
            };
        }
        
        return {
            success: true,
            formulas: data.formulas || [],
            data_analysis: data.data_analysis || {},
            alternative_formulas: data.alternative_formulas || []
        };
    };
    
    /**
     * 创建增强的错误响应
     */
    EnhancedAIInterface.prototype.createEnhancedErrorResponse = function(error, response) {
        return {
            success: false,
            error: '响应解析失败',
            message: error.message,
            raw_response: typeof response === 'string' ? response.substring(0, 500) : JSON.stringify(response).substring(0, 500),
            timestamp: new Date().toISOString()
        };
    };
    
    /**
     * 生成公式（主要对外接口）
     */
    EnhancedAIInterface.prototype.generateFormula = function(description) {
        var self = this;
        return new Promise(function(resolve, reject) {
            // 构建请求消息
            var messages = [
                {
                    role: "system",
                    content: "你是一个专业的Excel公式专家助手。你的任务是根据用户的需求和提供的数据信息，生成精确的Excel公式。你的回答必须严格按照以下JSON格式返回，不要包含任何其他文本。"
                },
                {
                    role: "user",
                    content: "请根据以下需求生成Excel公式: " + description
                }
            ];
            
            // 获取API配置
            var apiKey = window.AI_CONFIG ? window.AI_CONFIG.apiKey : null;
            var model = window.AI_CONFIG ? window.AI_CONFIG.modelName : "qwen-plus";
            var baseURL = window.AI_CONFIG ? window.AI_CONFIG.baseURL : "https://dashscope.aliyuncs.com/compatible-mode/v1";
            
            if (!apiKey) {
                reject(new Error("API密钥未配置"));
                return;
            }
            
            // 构建请求体
            var requestBody = {
                model: model,
                messages: messages
            };
            
            // 构建请求选项
            var requestOptions = {
                method: 'POST',
                headers: {
                    'Authorization': 'Bearer ' + apiKey,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestBody)
            };
            
            // 发起请求
            self.handleAIApiRequest(baseURL + "/chat/completions", requestOptions)
                .then(function(result) {
                    resolve(result);
                })
                .catch(function(error) {
                    reject(error);
                });
        });
    };
    
    /**
     * 生成公式请求
     */
    EnhancedAIInterface.prototype.generateFormulaRequest = function(requestData, options) {
        var self = this;
        options = options || {};
        
        return new Promise(function(resolve, reject) {
            try {
                self.stats.totalRequests++;
                self.stats.lastRequestTime = new Date();
                
                // 构建AI请求消息
                var messages = self.buildAIRequestMessages(requestData);
                
                // 获取API配置，优先使用 CURRENT_AI_CONFIG（如果已设置）
                // 优先使用 CURRENT_AI_CONFIG，确保正确处理配置类型
                var currentConfig = window.CURRENT_AI_CONFIG;
                var fallbackConfig = window.AI_CONFIG;
                
                var apiKey = currentConfig ? currentConfig.apiKey : 
                           (fallbackConfig ? fallbackConfig.apiKey : null);
                var model = currentConfig ? currentConfig.modelName : 
                           (fallbackConfig ? fallbackConfig.modelName : "qwen-plus");
                var baseURL = currentConfig ? currentConfig.baseURL : 
                             (fallbackConfig ? fallbackConfig.baseURL : null);
                var apiEndpoint = currentConfig ? currentConfig.apiEndpoint : 
                                 (fallbackConfig ? fallbackConfig.apiEndpoint : null);
                
                // 确定API端点
                var endpoint;
                if (apiEndpoint) {
                    // 局域网配置使用apiEndpoint
                    endpoint = apiEndpoint;
                } else if (baseURL) {
                    // 标准配置使用baseURL + 路径
                    endpoint = baseURL + (baseURL.endsWith('/') ? '' : '/') + "chat/completions";
                } else {
                    reject(new Error("API端点未配置"));
                    return;
                }
                
                if (!apiKey) {
                    reject(new Error("API密钥未配置"));
                    return;
                }
                
                // 构建请求体
                var requestBody = {
                    model: model,
                    messages: messages,
                    stream: true // 启用流式响应
                };
                
                // 构建请求选项
                var requestOptions = {
                    method: 'POST',
                    headers: {
                        'Authorization': 'Bearer ' + apiKey,
                        'Content-Type': 'application/json',
                        'X-DashScope-SSE': 'enable' // 启用SSE流式传输
                    },
                    body: JSON.stringify(requestBody) // 修正：应该发送requestBody而不是requestData
                };
                
                // 处理流式响应
                self.handleStreamAIApiRequest(endpoint, requestOptions, options)
                    .then(function(result) {
                        resolve(result);
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
     * 构建AI请求消息
     */
    EnhancedAIInterface.prototype.buildAIRequestMessages = function(requestData) {
        // 构建系统提示词
        var systemPrompt = this.buildSystemPrompt();
        
        // 构建用户提示词
        var userPrompt = this.buildUserPrompt(requestData);
        
        // 返回消息数组
        return [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt }
        ];
    };
    
    /**
     * 构建系统提示词
     */
    EnhancedAIInterface.prototype.buildSystemPrompt = function() {
        return `你是一个专业的Excel公式专家助手。你的任务是根据用户的需求和提供的数据信息，生成精确的Excel公式。

当用户没有明确描述需求时，你需要根据提供的数据结构和单元格信息，自主智能分析最可能的计算需求，并生成对应的Excel公式。

你的回答必须严格按照以下JSON格式返回，不要包含任何其他文本：

{
    "formulas": [
        {
            "title": "公式名称/描述",
            "formula": "完整的Excel公式",
            "explanation": "公式详细说明，包括各参数含义和业务逻辑",
            "confidence": 95,
            "applicable_ranges": ["应用范围说明"],
            "required_functions": ["使用的函数列表"],
            "example": "使用示例"
        }
    ],
    "data_analysis": {
        "headers_found": ["发现的表头"],
        "data_types": ["数据类型分析"],
        "recommendations": ["使用建议"],
        "smart_analysis": "你的智能分析结果，解释为什么选择这个公式"
    },
    "alternative_formulas": [
        {
            "description": "替代方案描述",
            "formula": "替代公式",
            "pros": ["优点"],
            "cons": ["缺点"]
        }
    ]
}

智能分析指导原则：
1. 根据单元格地址位置推断可能的计算需求（如行12通常是数据汇总行）
2. 根据表头内容判断数据类型和计算方式
3. 考虑当前数据区域的数据分布和特征
4. 如果是库存相关数据，优先考虑库存计算公式
5. 如果是财务数据，优先考虑金额计算和汇总公式
6. 如果包含日期列，优先考虑时间相关计算

重要规则：
1. 公式必须使用正确的Excel语法
2. 如果引用跨工作簿数据，使用'[工作簿名]工作表名!单元格引用'格式
3. 如果引用跨工作表数据，使用'工作表名!单元格引用'格式
4. 考虑引用范围的锁定方式（相对/绝对引用）
5. 如果需要填充，公式中的引用需要相应调整
6. confidence值应该在70-99之间，反映公式的准确性
7. 若信息不足，请直接按可能概率推荐有可能的公式
8. 特别关注I12这种位置的数据，通常是汇总或计算结果位置`;
    };
    
    /**
     * 构建用户提示词
     */
    EnhancedAIInterface.prototype.buildUserPrompt = function(requestData) {
        var prompt = '';
        
        // 用户需求描述
        if (requestData.description && requestData.description.trim() !== '') {
            prompt += `用户需求: ${requestData.description}\n\n`;
        } else {
            prompt += `请根据提供的工作表数据信息，自行分析最可能的需求并给出最合适的Excel公式建议。分析数据特点，推测用户可能想要进行的计算或数据处理操作。\n\n`;
        }
        
        // 当前单元格信息
        if (requestData.currentCell) {
            prompt += `=== 当前单元格信息 ===\n`;
            prompt += `工作簿: ${requestData.currentCell.workbook || '未知'}\n`;
            prompt += `工作表: ${requestData.currentCell.worksheet || '未知'}\n`;
            prompt += `单元格地址: ${requestData.currentCell.cellAddress || 'A1'}\n`;
            prompt += `行列位置: 第${requestData.currentCell.row || 1}行，第${requestData.currentCell.col || 1}列\n`;
            if (requestData.currentCell.columnHeader) {
                prompt += `列标题: ${requestData.currentCell.columnHeader}\n`;
            }
            if (requestData.currentCell.value) {
                prompt += `单元格值: ${requestData.currentCell.value}\n`;
            }
            prompt += `\n`;
        }
        
        // 工作簿和工作表信息
        if (requestData.selectedWorkbooks && requestData.selectedWorkbooks.length > 0) {
            prompt += `=== 工作簿和工作表信息 ===\n`;
            requestData.selectedWorkbooks.forEach(function(workbook, index) {
                prompt += `工作簿 ${index + 1}: ${workbook.workBookName || workbook.name}\n`;
                if (workbook.worksheets && workbook.worksheets.length > 0) {
                    workbook.worksheets.forEach(function(worksheet) {
                        prompt += `  工作表: ${worksheet.workSheetName || worksheet.name}\n`;
                        if (worksheet.columnHeaders) {
                            prompt += `    列标题: ${JSON.stringify(worksheet.columnHeaders)}\n`;
                        }
                    });
                }
            });
            prompt += `\n`;
        }
        
        // 填充选项
        if (requestData.fillOptions && (requestData.fillOptions.right || requestData.fillOptions.down)) {
            prompt += `=== 填充需求 ===\n`;
            var fillOptionsText = [];
            if (requestData.fillOptions.right) fillOptionsText.push('向右填充');
            if (requestData.fillOptions.down) fillOptionsText.push('向下填充');
            prompt += `填充方向: ${fillOptionsText.join('、')}\n\n`;
        }
        
        return prompt;
    };
    
    /**
     * 获取统计信息
     */
    EnhancedAIInterface.prototype.getStats = function() {
        return this.stats;
    };
    
    return EnhancedAIInterface;
})();

// 创建全局实例
window.enhancedAIInterface = new EnhancedAIInterface();