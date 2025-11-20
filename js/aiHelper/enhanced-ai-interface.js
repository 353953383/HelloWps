/**
 * 增强的AI接口管理器 - 整合代理服务器支持
 * 解决CORS错误、API配置和网络请求问题
 */

class EnhancedAIInterface {
    constructor() {
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
    setupProxyFetch() {
        if (!this.configManager) {
            console.warn('⚠️ API配置管理器未初始化，使用原始fetch');
            return;
        }
        
        window.fetch = async (url, options = {}) => {
            // 特殊处理AI API调用
            if (this.isAIApiCall(url, options)) {
                return this.handleAIApiRequest(url, options);
            }
            
            // 其他请求使用原始fetch
            return this.originalFetch(url, options);
        };
    }
    
    /**
     * 检查是否为AI API调用
     */
    isAIApiCall(url, options) {
        if (!options || !options.body) return false;
        
        try {
            const body = JSON.parse(options.body);
            return body && body.system && body.user;
        } catch (_) {
            return false;
        }
    }
    
    /**
     * 处理AI API请求
     */
    async handleAIApiRequest(url, options) {
        this.stats.totalRequests++;
        this.stats.lastRequestTime = new Date();
        
        try {
            const config = this.configManager.config;
            let requestUrl = url;
            let requestOptions = { ...options };
            
            // 如果使用代理，重定向到代理服务器
            if (config.useProxy && config.endpoint) {
                if (url.includes('dashscope.aliyuncs.com')) {
                    requestUrl = `${config.endpoint}/api/chat/completions`;
                    
                    // 移除不必要的headers（避免代理冲突）
                    delete requestOptions.headers['Authorization'];
                    delete requestOptions.headers['Content-Type'];
                    
                    // 添加代理所需的headers
                    requestOptions.headers = {
                        ...requestOptions.headers,
                        'Content-Type': 'application/json',
                        'X-Client-Version': 'ai-helper-v2.0'
                    };
                }
            }
            
            console.log('🌐 使用代理模式:', requestUrl);
            
            // 设置超时控制器
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), config.timeout);
            
            requestOptions.signal = controller.signal;
            
            // 执行请求
            const response = await this.originalFetch(requestUrl, requestOptions);
            clearTimeout(timeoutId);
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const result = await response.json();
            
            // 检查代理响应格式
            if (result.error) {
                throw new Error(result.error);
            }
            
            this.stats.successfulRequests++;
            console.log('✅ AI API请求成功');
            
            return this.formatAIResponse(result);
            
        } catch (error) {
            this.stats.failedRequests++;
            console.error('❌ AI API请求失败:', error);
            
            // 尝试备用方案
            return this.handleFallbackRequest(error);
        }
    }
    
    /**
     * 处理备用请求
     */
    async handleFallbackRequest(error) {
        console.log('🔄 尝试备用方案...');
        
        // 检查是否为CORS错误
        if (error.message.includes('CORS') || error.message.includes('blocked')) {
            return this.createCORSErrorResponse(error);
        }
        
        // 检查是否为网络错误
        if (error.message.includes('fetch') || error.message.includes('NetworkError')) {
            return this.createNetworkErrorResponse(error);
        }
        
        // 其他错误
        return this.createGenericErrorResponse(error);
    }
    
    /**
     * 创建CORS错误响应
     */
    createCORSErrorResponse(error) {
        console.log('🚫 CORS错误处理');
        
        return {
            success: false,
            error: 'CORS跨域错误',
            message: '无法直接访问API，请使用代理服务器',
            suggestions: [
                '启动代理服务器: server/start-proxy.sh',
                '或使用浏览器插件绕过CORS限制',
                '确保API端点配置正确'
            ],
            timestamp: new Date().toISOString()
        };
    }
    
    /**
     * 创建网络错误响应
     */
    createNetworkErrorResponse(error) {
        console.log('🌐 网络错误处理');
        
        return {
            success: false,
            error: '网络连接错误',
            message: '无法连接到API服务器',
            suggestions: [
                '检查网络连接是否正常',
                '确认代理服务器是否已启动',
                '检查API端点地址是否正确'
            ],
            timestamp: new Date().toISOString()
        };
    }
    
    /**
     * 创建通用错误响应
     */
    createGenericErrorResponse(error) {
        return {
            success: false,
            error: '请求失败',
            message: error.message,
            timestamp: new Date().toISOString()
        };
    }
    
    /**
     * 格式化AI响应
     */
    formatAIResponse(result) {
        try {
            // 检查OpenAI兼容格式
            if (result.choices && result.choices[0]) {
                const content = result.choices[0].message.content;
                return this.parseAIResponse(content);
            }
            
            // 直接返回代理响应
            return {
                success: true,
                data: result,
                timestamp: new Date().toISOString()
            };
            
        } catch (error) {
            console.warn('⚠️ 响应格式化失败:', error);
            return result;
        }
    }
    
    /**
     * 解析AI响应内容
     */
    parseAIResponse(content) {
        try {
            // 尝试提取JSON内容
            const jsonMatch = content.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                const jsonStr = jsonMatch[0];
                return JSON.parse(jsonStr);
            }
            
            // 如果没有JSON，尝试解析为文本
            return {
                success: true,
                formulas: [
                    {
                        description: content,
                        formula: '=0',
                        confidence: 0.5,
                        explanation: 'AI响应解析失败，使用默认公式'
                    }
                ],
                rawContent: content
            };
            
        } catch (error) {
            console.warn('⚠️ AI响应解析失败:', error);
            return this.createFallbackResponse('解析错误');
        }
    }
    
    /**
     * 创建备用响应
     */
    createFallbackResponse(reason) {
        return {
            success: true,
            formulas: [
                {
                    description: '基于数据计算的通用公式',
                    formula: '=SUM(数值1,数值2)',
                    confidence: 0.6,
                    explanation: `由于"${reason}"，提供通用公式模板`
                }
            ],
            timestamp: new Date().toISOString()
        };
    }
    
    /**
     * 生成公式（增强版）
     */
    async generateFormula(description, options = {}) {
        console.log('🤖 开始增强版公式生成...');
        console.log('📝 描述:', description);
        
        try {
            // 验证输入
            if (!description || description.trim() === '') {
                throw new Error('描述不能为空');
            }
            
            // 检查配置状态
            const status = this.configManager.getStatus();
            if (!status.isConfigured) {
                throw new Error('API配置未完成，请先进行配置');
            }
            
            // 构建请求数据
            const requestData = this.buildRequestData(description, options);
            
            // 添加到队列
            return await this.addToQueue(async () => {
                return await this.sendAIRequest(requestData);
            });
            
        } catch (error) {
            console.error('❌ 公式生成失败:', error);
            return this.createErrorFormula(error.message);
        }
    }
    
    /**
     * 构建请求数据
     */
    buildRequestData(description, options) {
        const currentCell = this.getCurrentCellInfo();
        const headers = this.getColumnHeaders();
        
        const systemPrompt = this.createSystemPrompt(headers);
        const userPrompt = this.createUserPrompt(description, currentCell, options);
        
        return {
            system: systemPrompt,
            user: userPrompt,
            description: description,
            currentCell: currentCell,
            columnHeaders: headers,
            options: options
        };
    }
    
    /**
     * 创建系统提示词
     */
    createSystemPrompt(headers) {
        const headerList = headers.join(', ');
        
        return `你是一个Excel公式专家。根据用户需求和数据结构，生成准确的Excel公式。

数据列: ${headerList}

要求:
1. 生成精确的Excel公式
2. 公式要符合实际业务逻辑
3. 给出清晰的解释
4. 评估公式的可信度(0-1)

响应格式:
{
  "formulas": [
    {
      "description": "描述公式用途",
      "formula": "具体公式",
      "confidence": 0.95,
      "explanation": "公式解释"
    }
  ]
}

确保返回的是有效的JSON格式。`;
    }
    
    /**
     * 创建用户提示词
     */
    createUserPrompt(description, currentCell, options) {
        let prompt = `当前单元格: ${currentCell.address} (列名: ${currentCell.columnName})\n`;
        prompt += `需要: ${description}\n`;
        
        if (options.context) {
            prompt += `附加信息: ${options.context}\n`;
        }
        
        prompt += `\n请生成相应的Excel公式。`;
        
        return prompt;
    }
    
    /**
     * 发送AI请求
     */
    async sendAIRequest(requestData) {
        console.log('📡 发送AI请求...');
        
        // 构建API请求
        const apiRequest = {
            model: this.configManager.config.modelName,
            messages: [
                {
                    role: 'system',
                    content: requestData.system
                },
                {
                    role: 'user',
                    content: requestData.user
                }
            ],
            temperature: 0.7,
            max_tokens: 2000
        };
        
        // 发送请求
        const response = await this.callAIApi(apiRequest);
        
        // 解析响应
        return this.parseAIResponse(response);
    }
    
    /**
     * 调用AI API
     */
    async callAIApi(request) {
        // 使用增强版fetch方法
        return await fetch(this.configManager.config.endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(request)
        });
    }
    
    /**
     * 队列管理
     */
    async addToQueue(task) {
        return new Promise((resolve, reject) => {
            this.requestQueue.push({
                task,
                resolve,
                reject,
                timestamp: Date.now()
            });
            
            this.processQueue();
        });
    }
    
    /**
     * 处理队列
     */
    async processQueue() {
        if (this.isProcessing || this.requestQueue.length === 0) {
            return;
        }
        
        this.isProcessing = true;
        
        try {
            const item = this.requestQueue.shift();
            
            // 检查超时（30秒）
            if (Date.now() - item.timestamp > 30000) {
                item.reject(new Error('请求超时'));
                return;
            }
            
            const result = await item.task();
            item.resolve(result);
            
        } catch (error) {
            console.error('❌ 队列处理错误:', error);
            
            if (this.requestQueue.length > 0) {
                const item = this.requestQueue.shift();
                item.reject(error);
            }
        } finally {
            this.isProcessing = false;
            
            // 继续处理队列
            if (this.requestQueue.length > 0) {
                setTimeout(() => this.processQueue(), 100);
            }
        }
    }
    
    /**
     * 获取当前单元格信息
     */
    getCurrentCellInfo() {
        try {
            if (typeof Excel !== 'undefined' && Excel.context) {
                const cell = Excel.context.workbook.worksheets.getActiveWorksheet().getRange(Excel.context.workbook.worksheets.getActiveWorksheet().rangeAddress);
                return {
                    address: cell.address,
                    columnName: this.getColumnName(cell.columnIndex - 1),
                    row: cell.row,
                    column: cell.columnIndex
                };
            }
        } catch (error) {
            console.warn('⚠️ 无法获取单元格信息:', error);
        }
        
        return {
            address: '未知单元格',
            columnName: '未知列',
            row: 0,
            column: 0
        };
    }
    
    /**
     * 获取列标题
     */
    getColumnHeaders() {
        try {
            if (typeof Excel !== 'undefined' && Excel.context) {
                const range = Excel.context.workbook.worksheets.getActiveWorksheet().getRange('1:1');
                const values = range.values[0];
                return values.filter(val => val && val.trim() !== '');
            }
        } catch (error) {
            console.warn('⚠️ 无法获取列标题:', error);
        }
        
        return ['数据列'];
    }
    
    /**
     * 获取列名
     */
    getColumnName(index) {
        const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        if (index < 26) {
            return alphabet[index];
        } else {
            return alphabet[Math.floor(index / 26) - 1] + alphabet[index % 26];
        }
    }
    
    /**
     * 创建错误公式响应
     */
    createErrorFormula(errorMessage) {
        return {
            success: false,
            error: errorMessage,
            formulas: [
                {
                    description: '错误处理公式',
                    formula: '=IFERROR(0,"发生错误")',
                    confidence: 0.1,
                    explanation: `由于错误"${errorMessage}"，提供错误处理公式`
                }
            ],
            timestamp: new Date().toISOString()
        };
    }
    
    /**
     * 获取统计信息
     */
    getStats() {
        return {
            ...this.stats,
            queueLength: this.requestQueue.length,
            isProcessing: this.isProcessing,
            successRate: this.stats.totalRequests > 0 
                ? (this.stats.successfulRequests / this.stats.totalRequests * 100).toFixed(2) + '%'
                : '0%'
        };
    }
    
    /**
     * 重置统计
     */
    resetStats() {
        this.stats = {
            totalRequests: 0,
            successfulRequests: 0,
            failedRequests: 0,
            lastRequestTime: null
        };
        console.log('📊 统计信息已重置');
    }
    
    /**
     * 测试连接
     */
    async testConnection() {
        console.log('🔍 开始连接测试...');
        
        if (!this.configManager) {
            return {
                success: false,
                error: '配置管理器未初始化'
            };
        }
        
        return await this.configManager.testConnection();
    }
}

// 创建全局实例
window.enhancedAIInterface = new EnhancedAIInterface();

// 替换原始接口（向后兼容）
window.aiInterface = window.enhancedAIInterface;

// 导出增强接口
console.log('🚀 增强AI接口已初始化');
console.log('📊 统计:', window.enhancedAIInterface.getStats());