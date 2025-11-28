/**
 * AI接口模块
 * 负责与AI大模型通信，配置固定的JSON格式规范
 */

class AIInterface {
    constructor() {
        this.apiEndpoint = '';
        this.apiKey = '';
        this.modelName = '';
        this.timeout = 30000; // 30秒超时
        
        this.init();
    }
    
    init() {
        this.loadConfig();
        this.setupEventListeners();
    }
    
    /**
     * 加载配置
     */
    loadConfig() {
        try {
            // 读取全局配置并设置到实例变量
            this.loadAIConfig();
            
    
            
        } catch (error) {
            console.error('❌ AI配置加载失败:', error);
            throw new Error(`AI配置加载失败: ${error.message}`);
        }
    }
    
    /**
     * 从全局配置文件加载AI设置（OpenAI格式）
     */
    loadAIConfig() {
        // 从全局变量读取AI配置，优先使用 CURRENT_AI_CONFIG（如果已设置）
        const globalAIConfig = window.CURRENT_AI_CONFIG || window.AI_CONFIG;
        
        if (!globalAIConfig) {
            throw new Error('❌ 全局AI配置不存在，请确保util.js已正确加载AI_CONFIG配置');
        }
        
        // 验证必要配置项（OpenAI格式）
        if (!globalAIConfig.baseURL && !globalAIConfig.apiEndpoint) {
            throw new Error('❌ AI配置错误：缺少baseURL或apiEndpoint参数');
        }
        
        if (!globalAIConfig.apiKey) {
            throw new Error('❌ AI配置错误：缺少apiKey参数');
        }
        
        if (!globalAIConfig.modelName) {
            throw new Error('❌ AI配置错误：缺少modelName参数');
        }
        
        // 读取配置值（OpenAI格式）
        // 支持两种配置格式：标准OpenAI格式和局域网格式
        if (globalAIConfig.baseURL) {
            // 标准OpenAI格式
            this.baseURL = globalAIConfig.baseURL;
            this.apiEndpoint = `${this.baseURL}/chat/completions`;
        } else if (globalAIConfig.apiEndpoint) {
            // 局域网格式
            this.apiEndpoint = globalAIConfig.apiEndpoint;
        }
        
        this.apiKey = globalAIConfig.apiKey;
        this.modelName = globalAIConfig.modelName;
        this.maxTokens = globalAIConfig.maxTokens || 30000;
        this.temperature = globalAIConfig.temperature || 0.3;
        this.timeout = globalAIConfig.timeout || 60000;
        this.requestFormat = globalAIConfig.requestFormat || 'openai'; // 添加请求格式支持
        
        // 如果有备用配置，设置为备用配置
        if (window.AI_CONFIG_LOCAL) {
            this.localConfig = window.AI_CONFIG_LOCAL;
        } else if (window.AI_CONFIG) {
            // 如果当前使用的是局域网配置，则将云端配置作为备用
            if (window.CURRENT_AI_CONFIG === window.AI_CONFIG_WLAN && window.AI_CONFIG) {
                this.localConfig = window.AI_CONFIG;
            }
            // 如果当前使用的是云端配置，则将局域网配置作为备用
            else if (window.CURRENT_AI_CONFIG !== window.AI_CONFIG_WLAN && window.AI_CONFIG_WLAN) {
                this.localConfig = window.AI_CONFIG_WLAN;
            }
        }
        
        // 显示当前配置信息
        console.log('🤖 AI配置已加载:');
        console.log(`   API端点: ${this.apiEndpoint}`);
        console.log(`   模型: ${this.modelName}`);
        console.log(`   最大Token: ${this.maxTokens}`);
        console.log(`   温度: ${this.temperature}`);
        if (globalAIConfig.description) {
            console.log(`   说明: ${globalAIConfig.description}`);
        }
    }
    
    /**
     * 设置事件监听器
     */
    setupEventListeners() {
        // 可以在这里添加配置更新的监听器
        document.addEventListener('aiConfigUpdated', (e) => {
            this.updateConfig(e.detail);
        });
    }
    
    /**
     * 更新配置
     */
    updateConfig(newConfig) {
        this.apiEndpoint = newConfig.apiEndpoint || this.apiEndpoint;
        this.apiKey = newConfig.apiKey || this.apiKey;
        this.modelName = newConfig.modelName || this.modelName;
    }
    
    /**
     * 生成公式建议
     */
    async generateFormula(requestData) {
        if (!this.apiEndpoint || !this.modelName) {
            throw new Error('❌ 主API端点或模型未配置');
        }
        
        try {
            return await this.generateFormulaWithEndpoint(requestData, this.apiEndpoint, this.modelName);
        } catch (error) {
            console.warn('⚠️ 主API调用失败:', error.message);
            
            // 如果有局域网配置，尝试切换到局域网配置
            if (this.localConfig) {
                console.log('🔄 尝试切换到局域网AI服务...');
                try {
                    return await this.generateFormulaWithEndpoint(requestData, this.localConfig.apiEndpoint, this.localConfig.modelName, this.localConfig);
                } catch (localError) {
                    console.error('❌ 局域网API调用也失败:', localError.message);
                    throw new Error(`AI服务连接失败:\n主服务: ${error.message}\n局域网服务: ${localError.message}`);
                }
            } else {
                console.error('❌ 没有可用的备用AI服务配置');
                throw new Error(`AI服务错误: ${error.message}`);
            }
        }
    }

    /**
     * 使用指定端点生成公式
     */
    async generateFormulaWithEndpoint(requestData, endpoint, modelName = null, configOverrides = null) {
        try {
            console.log('🔄 开始AI公式生成流程...');
            
            // 🔧 关键修复：创建请求数据的完全独立副本
            // 使用深拷贝确保即使原始对象被修改也不会影响AI调用
            const requestDataCopy = JSON.parse(JSON.stringify(requestData));
            console.log('🔍 [调试] 创建请求数据副本，原对象ID:', requestData.__id || '无');
            console.log('🔍 [调试] 副本对象ID:', requestDataCopy.__id || '新副本');
            console.log('🔍 [调试] 副本description长度:', requestDataCopy.description ? requestDataCopy.description.length : 0);
            
            // 🔧 新增：保护性措施 - 在全局变量中保存完整的description
            if (requestDataCopy.description && typeof requestDataCopy.description === 'string' && requestDataCopy.description.length > 0) {
                window.__protectedDescription = {
                    content: requestDataCopy.description,
                    length: requestDataCopy.description.length,
                    timestamp: Date.now(),
                    requestId: Date.now() + '_' + Math.random().toString(36).substr(2, 9)
                };
                console.log('🔒 [保护] 已将完整description保存到全局保护变量');
                console.log('🔒 [保护] 保护描述长度:', window.__protectedDescription.length);
                console.log('🔒 [保护] 请求ID:', window.__protectedDescription.requestId);
            }
            
            // 使用副本数据进行AI请求构建
            const aiRequest = this.buildAIRequest(requestDataCopy);
            console.log('📊 数据完整性检查完成，使用副本数据');
            
            // 🔧 新增：检查是否需要使用智能分析器
            // 当description为空或无效时，使用智能分析器分析需求
            const needsIntelligentAnalysis = !requestDataCopy.description || 
                                          requestDataCopy.description === "" ||
                                          requestDataCopy.description === "undefined" ||
                                          requestDataCopy.description === "null" ||
                                          (typeof requestDataCopy.description === "string" && requestDataCopy.description.trim() === "") ||
                                          requestDataCopy.description.includes('自行分析最可能的需求');
            
            if (needsIntelligentAnalysis) {
                console.log('🔍 检测到空描述，启用智能分析器');
                const analyzer = new IntelligentAnalyzer();
                const analysisResult = analyzer.analyze(requestDataCopy);
                console.log('🧠 智能分析结果:', analysisResult);
                
                // 更新AI请求中的描述
                if (analysisResult.suggestedDescription) {
                    aiRequest.description = analysisResult.suggestedDescription;
                    if (aiRequest.messages && aiRequest.messages[1]) {
                        aiRequest.messages[1].content = analysisResult.suggestedDescription;
                    }
                }
            }
            
            // 显示请求构建完成状态
            console.log('✅ AI请求构建完成');
            console.log('📝 请求描述预览:', aiRequest.description ? 
                         (aiRequest.description.length > 100 ? 
                          aiRequest.description.substring(0, 100) + '...' : 
                          aiRequest.description) : 
                         '无描述');
            
            // 获取当前配置（考虑覆盖配置）
            const effectiveConfig = configOverrides || this;
            const apiKey = effectiveConfig.apiKey;
            const effectiveModelName = modelName || effectiveConfig.modelName;
            const requestFormat = effectiveConfig.requestFormat || 'openai';
            
            // 🔧 关键修复：确保AI请求中的模型名称正确
            aiRequest.model = effectiveModelName;
            
            // 显示发送给AI的请求信息
            console.log('📤 发送给AI的请求信息:');
            console.log(`   模型: ${effectiveModelName}`);
            console.log(`   端点: ${endpoint}`);
            console.log(`   请求格式: ${requestFormat}`);
            
            // 显示消息内容详情
            if (aiRequest.messages) {
                console.log(`   消息数量: ${aiRequest.messages.length}`);
                aiRequest.messages.forEach((msg, index) => {
                    const role = msg.role === 'system' ? '系统' : '用户';
                    // 系统消息显示1200字符，用户消息显示200字符
                    const maxLength = msg.role === 'system' ? 1200 : 200;
                    const content = msg.content.length > maxLength ? msg.content.substring(0, maxLength) + '...' : msg.content;
                    console.log(`   消息${index + 1} (${role}): ${content}`);
                });
            } else if (aiRequest.input && aiRequest.input.messages) {
                console.log(`   消息数量: ${aiRequest.input.messages.length}`);
                aiRequest.input.messages.forEach((msg, index) => {
                    const role = msg.role === 'system' ? '系统' : '用户';
                    const maxLength = msg.role === 'system' ? 1200 : 200;
                    const content = msg.content.length > maxLength ? msg.content.substring(0, maxLength) + '...' : msg.content;
                    console.log(`   消息${index + 1} (${role}): ${content}`);
                });
            } else {
                console.log('   消息内容: 无法解析');
            }
            
            // 显示请求参数
            if (aiRequest.parameters) {
                console.log(`   温度: ${aiRequest.parameters.temperature}, 最大Token: ${aiRequest.parameters.max_tokens}`);
            } else if (aiRequest.max_tokens) {
                console.log(`   最大Token: ${aiRequest.max_tokens}`);
            }
            
            // 根据请求格式设置不同的请求头和请求体
            const headers = {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            };
            
            // 构建适用于不同格式的请求体
            let finalRequestBody;
            if (requestFormat === 'compatible' || endpoint.includes('/v1/chat/completions')) {
                // 局域网兼容格式或标准OpenAI格式
                finalRequestBody = {
                    model: effectiveModelName,
                    messages: aiRequest.messages || aiRequest.input?.messages,
                    max_tokens: aiRequest.max_tokens || aiRequest.parameters?.max_tokens || 30000,
                    temperature: aiRequest.temperature || aiRequest.parameters?.temperature || 0.3
                };
            } else {
                // 其他格式保持原有结构
                finalRequestBody = aiRequest;
            }
            
            // 发送请求（针对CORS进行优化）
            const fetchOptions = {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(finalRequestBody),
                mode: 'cors',
                credentials: 'omit',
                // 添加cache控制避免缓存问题
                cache: 'no-cache'
            };
            
            let response;
            try {
                console.log('🌐 发送请求到:', endpoint);
                response = await fetch(endpoint, fetchOptions);
            } catch (corsError) {
                // 如果是局域网配置且CORS错误，尝试使用代理
                if (window.CURRENT_AI_CONFIG === window.AI_CONFIG_WLAN && window.enhancedAIInterface) {
                    console.log('🔄 局域网配置遇到CORS问题，尝试使用代理...');
                    response = await window.enhancedAIInterface.handleAIApiRequest(endpoint, fetchOptions);
                } else {
                    throw corsError;
                }
            }
            
            // 检查响应状态
            if (!response.ok) {
                const errorText = await response.text();
                console.error('❌ AI响应错误:', errorText);
                throw new Error(`AI响应错误: ${errorText}`);
            }
            
            // 解析响应
            const responseData = await response.json();
            console.log('✅ AI响应解析完成');
            console.log('📝 响应内容预览:', responseData.choices ? 
                         (responseData.choices[0].message.content.length > 100 ? 
                          responseData.choices[0].message.content.substring(0, 100) + '...' : 
                          responseData.choices[0].message.content) : 
                         '无内容');
            
            // 🔧 清理：成功后清理全局保护变量
            delete window.__protectedDescription;
            
            return responseData.choices[0].message.content;
        } catch (error) {
            console.error('❌ AI公式生成失败:', error.message);
            // 🔧 清理：错误时也清理全局保护变量
            delete window.__protectedDescription;
            throw new Error(`AI公式生成失败: ${error.message}`);
        }
    }
    
    /**
     * 构建AI请求（支持兼容模式和标准格式）
     */
    buildAIRequest(requestData) {
        // 🔧 关键修复：使用独立的深拷贝，避免原始对象被修改
        const originalData = JSON.parse(JSON.stringify(requestData));
        
        console.log('🔍 [调试] 原始requestData详情:');
        console.log('  - 对象类型:', typeof requestData);
        console.log('  - 对象字段数:', requestData ? Object.keys(requestData).length : 0);
        console.log('  - 对象字段列表:', requestData ? Object.keys(requestData) : []);
        console.log('  - description长度:', requestData?.description ? requestData.description.length : 0);
        console.log('  - description前100字符:', requestData?.description ? requestData.description.substring(0, 100) : 'null/undefined');
        
        // 🔧 关键修复：在数据完整性检查之前先记录原始description
        console.log('🔍 [调试] buildAIRequest接收到原始description:', typeof requestData?.description, requestData?.description ? requestData.description.length : 0);
        console.log('🔍 [调试] 原始description前100字符:', requestData?.description ? requestData.description.substring(0, 100) : 'null/undefined');
        
        // 数据完整性检查和修复 - 使用独立的拷贝
        const safeRequestData = this.ensureDataIntegrity(originalData);
        
        // 🔧 关键修复：在构建AI请求前再次确认description完整性
        console.log('🔍 [调试] 数据完整性检查后description:', typeof safeRequestData.description, safeRequestData.description ? safeRequestData.description.length : 0);
        console.log('🔍 [调试] 检查后description前100字符:', safeRequestData.description ? safeRequestData.description.substring(0, 100) : 'null/undefined');
        
        let requestBody;
        
        // 🔧 关键修复：确保description在传递过程中不被修改
        const userPrompt = this.buildUserPrompt(safeRequestData);
        const systemPrompt = this.buildSystemPrompt(safeRequestData);
        
        // 使用OpenAI格式构建请求体
        requestBody = {
            model: this.modelName,
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: userPrompt }
            ],
            max_tokens: this.maxTokens || 30000,
            temperature: this.temperature || 0.3,
            top_p: 0.8
        };
        console.log('🎯 [调试] 使用OpenAI格式API，maxTokens设置为:', this.maxTokens);
        
        // 🔧 关键修复：确认构建的请求体中的description内容
        console.log('🔍 [调试] 构建的AI请求体中用户提示词长度:', userPrompt.length);
        console.log('🔍 [调试] 用户提示词前200字符:', userPrompt.substring(0, 200));
        
        return requestBody;
    }
    
    /**
     * 确保数据完整性 - 检查和修复可能的undefined数据
     */
    ensureDataIntegrity(requestData) {
        // 数据完整性检查开始
        console.log('🔍 开始数据完整性检查...');
        
        // 如果requestData为null或undefined，创建默认对象
        if (!requestData) {
            console.log('⚠️ 使用默认数据配置');
            return {
                description: "智能分析当前工作表数据，提供Excel公式建议。",
                referenceType: "current",
                currentCell: null,
                selectedWorkbooks: [],
                selectedWorksheets: [],
                fillOptions: { right: false, down: false },
                headers: []
            };
        }
        
        // 🔧 关键修复：使用深拷贝创建完全独立的安全副本
        // 确保嵌套对象也不会与原始对象共享引用
        console.log('🔍 [调试] 原始requestData对象ID:', requestData.__id || '无ID');
        const safeData = JSON.parse(JSON.stringify(requestData));
        safeData.__id = 'safe_' + Date.now(); // 添加唯一标识
        
        let fixes = 0;
        
        // 🔧 关键修复：增强description检查逻辑，确保不会丢失
        console.log('🔍 [调试] 数据完整性检查 - 检查description字段...');
        console.log('🔍 [调试] 原始description:', typeof safeData.description, safeData.description);
        console.log('🔍 [调试] description是否存在:', !!safeData.description);
        console.log('🔍 [调试] description长度:', safeData.description ? safeData.description.length : 0);
        console.log('🔍 [调试] description前100字符:', safeData.description ? safeData.description.substring(0, 100) : 'null/undefined');
        
        // 🔧 增强的description处理逻辑
        if (safeData.description === undefined || safeData.description === null) {
            // 🔧 关键修复：首先尝试从全局保护变量恢复完整的description
            if (window.__protectedDescription && window.__protectedDescription.content) {
                console.log('🔒 [恢复] 从全局保护变量恢复完整description');
                console.log('🔒 [恢复] 保护描述长度:', window.__protectedDescription.length);
                safeData.description = window.__protectedDescription.content;
                console.log('🔒 [恢复] 已恢复description，长度:', safeData.description.length);
            } else {
                console.log('⚠️ [调试] description为undefined/null且无保护描述，使用默认值');
                safeData.description = "根据所提供信息，以当前单元格信息（当前工作簿、工作表名和表头信息）为主，其他信息为辅智能分析需求。";
            }
            fixes++;
        } else if (typeof safeData.description === 'string') {
            const trimmedDesc = safeData.description.trim();
            if (trimmedDesc === '') {
                // 🔧 关键修复：空字符串时也尝试从全局保护变量恢复
                if (window.__protectedDescription && window.__protectedDescription.content) {
                    console.log('🔒 [恢复] 从全局保护变量恢复空description');
                    safeData.description = window.__protectedDescription.content;
                    console.log('🔒 [恢复] 已恢复空description，长度:', safeData.description.length);
                } else {
                    console.log('⚠️ [调试] description为空字符串且无保护描述，使用默认智能分析描述');
                    safeData.description = "根据当前工作表的数据结构、表头信息和单元格位置，智能分析计算需求并提供合适的Excel公式建议。";
                }
                fixes++;
            } else {
                safeData.description = trimmedDesc; // 确保没有前后空格
                console.log('✅ [调试] description字段有效且已清理，保留原始值');
            }
        } else {
            console.log('⚠️ [调试] description类型异常，转换为字符串');
            safeData.description = String(safeData.description);
            fixes++;
        }
        
        // 🔧 最终确认description的完整性
        console.log('🔍 [调试] 最终确认description:', typeof safeData.description, safeData.description ? safeData.description.length : 0);
        if (!safeData.description || safeData.description === 'undefined' || safeData.description === 'null') {
            console.error('❌ [调试] description仍然无效，强制设置默认值');
            safeData.description = "根据当前工作表数据智能分析并提供Excel公式建议。";
            fixes++;
        }
        
        // 检查并修复referenceType
        if (!safeData.referenceType || safeData.referenceType === 'undefined') {
            safeData.referenceType = "current";
            fixes++;
        }
        
        // 检查并修复currentCell
        if (!safeData.currentCell || safeData.currentCell === 'undefined' || 
            (typeof safeData.currentCell === 'object' && (
                !safeData.currentCell.cellAddress || 
                !safeData.currentCell.worksheet || 
                safeData.currentCell.cellAddress === 'undefined' ||
                safeData.currentCell.worksheet === 'undefined'
            ))
        ) {
            // 如果currentCell存在但缺少必要字段，尝试从Excel获取
            try {
                if (window.Application && window.Application.ActiveSheet) {
                    const activeSheet = window.Application.ActiveSheet;
                    const selection = window.Application.Selection;
                    if (selection) {
                        safeData.currentCell = {
                            workbook: window.Application.ActiveWorkbook ? window.Application.ActiveWorkbook.Name : '',
                            worksheet: activeSheet ? activeSheet.Name : '',
                            row: selection.Row || 1,
                            col: selection.Column || 1,
                            cellAddress: getCellAddressSafe(selection.Row || 1, selection.Column || 1)
                        };
                    } else {
                        safeData.currentCell = null;
                    }
                } else {
                    safeData.currentCell = null;
                }
            } catch (error) {
                console.warn('获取当前单元格信息失败:', error);
                safeData.currentCell = null;
            }
            fixes++;
        }
        
        // 辅助函数：安全获取单元格地址
        function getCellAddressSafe(row, col) {
            try {
                const columnLetters = getColumnLettersSafe(col);
                return `${columnLetters}${row}`;
            } catch (error) {
                return `A${row}`; // 默认返回A列
            }
        }
        
        // 辅助函数：安全获取列字母
        function getColumnLettersSafe(col) {
            try {
                let temp = '';
                let columnNumber = col;
                
                while (columnNumber > 0) {
                    let remainder = (columnNumber - 1) % 26;
                    temp = String.fromCharCode(65 + remainder) + temp;
                    columnNumber = Math.floor((columnNumber - 1) / 26);
                }
                
                return temp;
            } catch (error) {
                return 'A'; // 默认返回A列
            }
        }
        
        // 检查并修复数组字段
        if (!Array.isArray(safeData.selectedWorkbooks)) {
            safeData.selectedWorkbooks = [];
            fixes++;
        }
        
        if (!Array.isArray(safeData.selectedWorksheets)) {
            safeData.selectedWorksheets = [];
            fixes++;
        }
        
        if (!Array.isArray(safeData.headers)) {
            safeData.headers = [];
            fixes++;
        }
        
        // 检查并修复fillOptions
        if (!safeData.fillOptions || safeData.fillOptions === 'undefined') {
            safeData.fillOptions = { right: false, down: false };
            fixes++;
        }
        
        // 数据完整性检查完成 - 添加description最终值调试日志
        console.log(`✅ 数据完整性检查完成，修复了${fixes}处异常`);
        console.log('🔍 [调试] 修复后的description字段信息:');
        console.log('🔍 [调试] 修复后description长度:', safeData.description ? safeData.description.length : 0);
        console.log('🔍 [调试] 修复后description类型:', typeof safeData.description);
        console.log('🔍 [调试] 修复后description前150字符:', safeData.description ? safeData.description.substring(0, 150) : 'null/undefined');
        return safeData;
    }
    
    /**
     * 构建系统提示词
     */
    buildSystemPrompt(requestData) {
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
    }
    
    /**
     * 构建用户提示词
     */
    buildUserPrompt(requestData) {
        // 🔧 安全处理：先深拷贝requestData，确保不修改原始对象
        console.log('🔍 [调试] buildUserPrompt接收到requestData:', typeof requestData, requestData ? Object.keys(requestData) : 'null');
        const safeRequestData = requestData ? JSON.parse(JSON.stringify(requestData)) : null;
        
        // 安全地处理可能为undefined的属性
        const {
            description = "",
            referenceType = "current",
            currentCell,
            selectedWorkbooks,
            selectedWorksheets,
            fillOptions,
            headers,
            currentWorkbook,
            currentWorksheet,
            allWorksheets,
            columnHeaders
        } = safeRequestData || {};
        
        let prompt = '';
        
        // 检查是否用户没有输入具体需求信息 - 更精确的检查，添加调试日志
        console.log('🔍 [调试] 构建用户提示词 - 检查description字段...');
        console.log('🔍 [调试] buildUserPrompt中description类型:', typeof description);
        console.log('🔍 [调试] buildUserPrompt中description存在:', !!description);
        console.log('🔍 [调试] buildUserPrompt中description长度:', description ? description.length : 0);
        console.log('🔍 [调试] buildUserPrompt中description前200字符:', description ? description.substring(0, 200) : 'null/undefined');
        
        // 🔧 简化空值检测逻辑 - 只检测真正的空值和无效值
        const isEmptyDescription = !description || 
                                  description === "undefined" || 
                                  description === "null" ||
                                  (typeof description === "string" && description.trim() === "");
        
        console.log('🔍 [调试] isEmptyDescription检查结果:', isEmptyDescription);
        console.log('🔍 [调试] 检查条件详情:', {
            '!description': !description,
            'description类型': typeof description,
            'description === "undefined"': description === "undefined",
            'description === "null"': description === "null",
            'description.trim() === ""': description ? description.trim() === "" : 'N/A',
            'description有效内容': description ? description.substring(0, 50) + '...' : 'null/undefined'
        });
        
        if (isEmptyDescription) {
            console.log('⚠️ [调试] description被判定为空，使用默认提示词');
            prompt = `请根据提供的工作表数据信息，自行分析最可能的需求并给出最合适的Excel公式建议。分析数据特点，推测用户可能想要进行的计算或数据处理操作。`;
        } else {
            console.log('✅ [调试] description有效，maxTokens已提升至30000，无需长度限制');
            console.log('🔍 [调试] 当前description长度:', description.length);
            
            // 🎯 优化：既然maxTokens已提升至30000，直接使用完整description
            // 移除不必要的压缩逻辑，让AI获得完整的用户需求信息
            const originalDescriptionLength = description.length;
            const descriptionForPrompt = description;
            
            // 🔧 保留：将完整description存储到全局，供clearAll()恢复使用
            if (typeof window !== 'undefined') {
                window.__originalDescription = description;
                console.log('💾 [调试] 完整description已保存到window.__originalDescription，长度:', description.length);
            }
            
            prompt = `用户需求: ${descriptionForPrompt}`;
            console.log('🔍 [调试] 处理后的description长度:', descriptionForPrompt.length);
            console.log('🎯 [调试] Token预算充足，无需截断处理');
        }
        
        console.log('🔍 [调试] 构建的用户提示词长度:', prompt.length);
        console.log('🔍 [调试] 用户提示词前300字符:', prompt.substring(0, 300));
        
        // =================== 当前工作表详细信息 ===================
        if (referenceType === 'current' && currentWorkbook && currentWorksheet) {
            prompt += `\n\n=== 当前工作表详细信息 ===`;
            
            // 当前工作簿信息
            prompt += `\n📁 当前工作簿: ${currentWorkbook.name || '未知'}`;
            
            // 所有工作表列表
            if (allWorksheets && allWorksheets.length > 0) {
                prompt += `\n📋 工作簿中的所有工作表 (${allWorksheets.length}个):`;
                allWorksheets.forEach((sheet, index) => {
                    prompt += `\n  ${index + 1}. ${sheet.name} (${sheet.usedRange?.rows || '?'}行 x ${sheet.usedRange?.columns || '?'}列)`;
                });
            }
            
            // 当前工作表信息
            prompt += `\n📄 当前工作表: ${currentWorksheet.name || '未知'}`;
            if (currentWorksheet.usedRange) {
                prompt += `\n   数据范围: ${currentWorksheet.usedRange.rows}行 x ${currentWorksheet.usedRange.columns}列`;
            }
        }
        
        // =================== 当前单元格详细信息 ===================
        if (currentCell && (currentCell.cellAddress || currentCell.worksheet)) {
            prompt += `\n\n=== 当前单元格信息 ===`;
            const cellInfo = [];
            if (currentCell.worksheet) cellInfo.push(`工作表: ${currentCell.worksheet}`);
            if (currentCell.cellAddress) cellInfo.push(`单元格地址: ${currentCell.cellAddress}`);
            if (currentCell.row !== undefined && currentCell.column !== undefined) {
                cellInfo.push(`位置: 第${currentCell.row + 1}行，第${currentCell.column + 1}列`);
            }
            if (currentCell.columnHeader) {
                cellInfo.push(`列标题: ${currentCell.columnHeader}`);
            }
            prompt += `\n${cellInfo.join(', ')}`;
        }
        
        // =================== 表头详细信息 ===================
        let allHeaders = [];
        
        // 当前工作表的所有列标题
        if (referenceType === 'current' && columnHeaders && columnHeaders.length > 0) {
            prompt += `\n\n=== 当前工作表所有列标题 ===`;
            prompt += `\n📊 总计${columnHeaders.length}列:`;
            columnHeaders.forEach((header, index) => {
                const colLetter = this.getColumnLetter(index + 1);
                prompt += `\n  ${colLetter}列: ${header || '(空)'}`;
            });
            allHeaders = allHeaders.concat(columnHeaders);
        }
        
        // 其他工作表的表头信息
        if (headers && headers.length > 0) {
            prompt += `\n\n=== 相关工作表表头信息 ===`;
            headers.forEach(headerGroup => {
                if (headerGroup.headers && headerGroup.headers.length > 0) {
                    prompt += `\n📋 ${headerGroup.worksheet}: ${headerGroup.headers.join(', ')}`;
                    allHeaders = allHeaders.concat(headerGroup.headers);
                }
            });
        }
        
        // =================== 引用数据源详细信息 ===================
        if (referenceType !== 'current') {
            if (referenceType === 'worksheet' && selectedWorksheets && selectedWorksheets.length > 0) {
                prompt += `\n\n=== 跨工作表数据源 ===`;
                prompt += `\n📂 引用工作表列表 (${selectedWorksheets.length}个):`;
                selectedWorksheets.forEach((worksheet, index) => {
                    prompt += `\n  ${index + 1}. ${worksheet.name}`;
                    if (worksheet.headers && worksheet.headers.length > 0) {
                        prompt += `\n     表头 (${worksheet.headers.length}列): ${worksheet.headers.join(', ')}`;
                    }
                    if (worksheet.usedRange) {
                        prompt += `\n     数据范围: ${worksheet.usedRange.rows}行 x ${worksheet.usedRange.columns}列`;
                    }
                });
                
            } else if (referenceType === 'workbook' && selectedWorkbooks && selectedWorkbooks.length > 0) {
                prompt += `\n\n=== 跨工作簿数据源 ===`;
                selectedWorkbooks.forEach((workbook, wbIndex) => {
                    prompt += `\n📁 工作簿 ${wbIndex + 1}: ${workbook.name}`;
                    if (workbook.worksheets && workbook.worksheets.length > 0) {
                        workbook.worksheets.forEach((worksheet, wsIndex) => {
                            prompt += `\n  📄 工作表 ${wsIndex + 1}: ${worksheet.name}`;
                            if (worksheet.headers && worksheet.headers.length > 0) {
                                prompt += `\n     表头 (${worksheet.headers.length}列): ${worksheet.headers.join(', ')}`;
                            }
                            if (worksheet.usedRange) {
                                prompt += `\n     数据范围: ${worksheet.usedRange.rows}行 x ${worksheet.usedRange.columns}列`;
                            }
                        });
                    }
                });
            }
        }
        
        // =================== 填充需求 ===================
        if (fillOptions && (fillOptions.right || fillOptions.down)) {
            prompt += `\n\n=== 填充需求 ===`;
            const fillOptionsText = [];
            if (fillOptions.right) fillOptionsText.push('向右填充');
            if (fillOptions.down) fillOptionsText.push('向下填充');
            prompt += `\n📋 填充方向: ${fillOptionsText.join('、')}`;
        }
        
        // =================== 智能分析指导 ===================
        if (isEmptyDescription && allHeaders.length > 0) {
            prompt += `\n\n=== 智能分析指导 ===`;
            prompt += `\n🔍 数据特点分析提示:`;
            prompt += `\n- 检测到的列标题: ${allHeaders.slice(0, 10).join(', ')}${allHeaders.length > 10 ? '...' : ''}`;
            prompt += `\n- 建议重点关注的可能操作:`;
            
            // 基于表头内容智能推测可能的计算需求
            const headerStr = allHeaders.join(' ').toLowerCase();
            const suggestions = [];
            
            if (headerStr.includes('金额') || headerStr.includes('价格') || headerStr.includes('成本')) {
                suggestions.push('数值计算(求和、平均值、汇总)');
            }
            if (headerStr.includes('日期') || headerStr.includes('时间')) {
                suggestions.push('日期计算(时间差、工作日计算)');
            }
            if (headerStr.includes('数量') || headerStr.includes('件数')) {
                suggestions.push('统计计算(计数、最大最小值)');
            }
            if (headerStr.includes('比例') || headerStr.includes('率')) {
                suggestions.push('比例计算(百分比、占比)');
            }
            if (headerStr.includes('排名') || headerStr.includes('排序')) {
                suggestions.push('排名计算(RANK、排序公式)');
            }
            
            if (suggestions.length === 0) {
                suggestions.push('数据查询(VLOOKUP、INDEX+MATCH)');
                suggestions.push('条件判断(IF、AND、OR)');
                suggestions.push('文本处理(CONCATENATE、LEFT、RIGHT)');
            }
            
            suggestions.forEach(suggestion => {
                prompt += `\n  • ${suggestion}`;
            });
            
            prompt += `\n\n请根据以上信息自主判断最合适的Excel公式，并提供详细的解释。`;
        }
        
        return prompt.trim();
    }
    
    /**
     * 获取列号对应的字母表示 (1 -> A, 2 -> B, ..., 26 -> Z, 27 -> AA)
     */
    getColumnLetter(columnNumber) {
        let result = '';
        while (columnNumber > 0) {
            columnNumber--;
            result = String.fromCharCode(65 + (columnNumber % 26)) + result;
            columnNumber = Math.floor(columnNumber / 26);
        }
        return result;
    }
    
    /**
     * 获取引用类型的中文描述
     */
    getReferenceTypeText(type) {
        const typeMap = {
            'current': '当前工作表',
            'worksheet': '跨工作表',
            'workbook': '跨工作簿'
        };
        return typeMap[type] || type;
    }
    
    /**
     * 发送AI请求
     */
    async sendAIRequest(requestData, configOverrides = null) {
        try {
            // 使用配置覆盖（如果有的话）
            const effectiveConfig = configOverrides || this;
            const apiEndpoint = configOverrides ? configOverrides.apiEndpoint : this.apiEndpoint;
            const apiKey = configOverrides ? configOverrides.apiKey : this.apiKey;
            const modelName = configOverrides ? configOverrides.modelName : this.modelName;
            const requestFormat = configOverrides ? configOverrides.requestFormat : this.requestFormat;
            
            // 检查网络连接状态
            if (!navigator.onLine) {
                throw new Error('网络连接已断开，请检查网络设置');
            }
            
            // 验证API端点格式
            try {
                new URL(apiEndpoint);
            } catch (urlError) {
                throw new Error(`API端点格式无效: ${apiEndpoint}`);
            }
            
            const aiRequest = this.buildAIRequest(requestData);
            aiRequest.model = modelName; // 确保使用正确的模型名
            
            // 显示发送给AI的请求信息
            console.log('📤 发送给AI的请求信息:');
            console.log(`   模型: ${modelName}`);
            console.log(`   端点: ${apiEndpoint}`);
            
            // 显示消息内容详情
            if (aiRequest.messages) {
                console.log(`   消息数量: ${aiRequest.messages.length}`);
                aiRequest.messages.forEach((msg, index) => {
                    const role = msg.role === 'system' ? '系统' : '用户';
                    // 系统消息显示1200字符，用户消息显示200字符
                    const maxLength = msg.role === 'system' ? 1200 : 200;
                    const content = msg.content.length > maxLength ? msg.content.substring(0, maxLength) + '...' : msg.content;
                    console.log(`   消息${index + 1} (${role}): ${content}`);
                });
            } else if (aiRequest.input && aiRequest.input.messages) {
                console.log(`   消息数量: ${aiRequest.input.messages.length}`);
                aiRequest.input.messages.forEach((msg, index) => {
                    const role = msg.role === 'system' ? '系统' : '用户';
                    const maxLength = msg.role === 'system' ? 1200 : 200;
                    const content = msg.content.length > maxLength ? msg.content.substring(0, maxLength) + '...' : msg.content;
                    console.log(`   消息${index + 1} (${role}): ${content}`);
                });
            }
            
            // 显示请求参数
            if (aiRequest.parameters) {
                console.log(`   温度: ${aiRequest.parameters.temperature}, 最大Token: ${aiRequest.parameters.max_tokens}`);
            }
            
            // 根据API端点类型设置不同的请求头
            const headers = {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            };
            
            // 发送请求（针对CORS进行优化）
            const fetchOptions = {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(aiRequest),
                mode: 'cors',
                credentials: 'omit',
                // 添加cache控制避免缓存问题
                cache: 'no-cache'
            };
            
            let response;
            try {
                console.log('🌐 发送CORS请求到:', apiEndpoint);
                response = await fetch(apiEndpoint, fetchOptions);
            } catch (corsError) {
                console.warn('⚠️ CORS请求失败:', corsError.message);
                
                // 检查是否是网络问题
                if (!navigator.onLine) {
                    throw new Error('网络连接已断开，请检查网络设置');
                }
                
                throw new Error(`CORS策略阻止访问API: ${apiEndpoint}。请确保API服务器支持CORS。`);
            }
            
        if (!response.ok) {
            // 提供更详细的HTTP错误诊断
            let errorMessage = `HTTP错误: ${response.status} ${response.statusText}`;
            
            // 🔧 关键修复：增强HTTP 400错误处理，显示更多调试信息
            if (response.status === 400) {
                errorMessage = `请求参数错误 (HTTP 400): 请检查API请求格式和参数设置`;
                
                // 🔧 添加详细的请求体信息用于调试
                try {
                    const requestBody = JSON.stringify(aiRequest);
                    const requestSize = requestBody.length;
                    const descLength = aiRequest.input?.messages?.[1]?.content?.length || 
                                     aiRequest.messages?.[1]?.content?.length || 0;
                    
                    console.error('🔍 HTTP 400错误详细分析:');
                    console.error(`   📏 请求体大小: ${requestSize} 字符`);
                    console.error(`   📝 描述字段长度: ${descLength} 字符`);
                    console.error(`   🔑 API密钥长度: ${this.apiKey ? this.apiKey.length : 0}`);
                    console.error(`   🎯 模型名称: ${this.modelName}`);
                    console.error(`   🔧 使用模式: ${this.apiEndpoint.includes('/compatible-mode/') ? '兼容模式' : '标准模式'}`);
                    
                    // 检查请求体格式
                    if (this.apiEndpoint.includes('/compatible-mode/')) {
                        if (!aiRequest.model) console.error('❌ 缺少model字段');
                        if (!aiRequest.messages || !Array.isArray(aiRequest.messages)) console.error('❌ 缺少messages数组');
                        if (!aiRequest.max_tokens) console.error('❌ 缺少max_tokens字段');
                    } else {
                        if (!aiRequest.model) console.error('❌ 缺少model字段');
                        if (!aiRequest.input?.messages || !Array.isArray(aiRequest.input.messages)) console.error('❌ 缺少input.messages数组');
                        if (!aiRequest.parameters?.max_tokens) console.error('❌ 缺少parameters.max_tokens字段');
                    }
                    
                } catch (debugError) {
                    console.error('❌ 获取调试信息失败:', debugError.message);
                }
                
            } else if (response.status === 401) {
                errorMessage = `认证失败 (HTTP 401): 请检查API密钥是否有效`;
            } else if (response.status === 403) {
                errorMessage = `权限拒绝 (HTTP 403): 请检查API访问权限`;
            } else if (response.status === 429) {
                errorMessage = `请求频率过高 (HTTP 429): 请稍后重试`;
            } else if (response.status >= 500) {
                errorMessage = `服务器内部错误 (HTTP ${response.status}): 请稍后重试`;
            }
            
            console.error(`❌ API请求失败: ${errorMessage}`);
            console.error('📊 错误详情:', {
                status: response.status,
                statusText: response.statusText,
                url: this.apiEndpoint,
                method: 'POST',
                headers: fetchOptions.headers,
                timestamp: new Date().toISOString()
            });
            
            // 🔧 关键修复：添加错误堆栈跟踪
            const errorStack = new Error().stack;
            console.error('📍 错误堆栈跟踪:', errorStack);
            
            throw new Error(errorMessage);
        }
            
            const result = await response.json();
            
            // 显示AI响应信息
            console.log('📥 AI响应信息:');
            console.log(`   ✅ AI响应成功 - 模型: ${this.modelName}`);
            console.log(`   响应时间: ${new Date().toLocaleString()}`);
            
            // 显示响应内容详情
            if (result.choices && result.choices.length > 0) {
                const content = result.choices[0].message.content;
                console.log(`   响应内容 (${content.length}字符):`);
                // 显示前500字符作为详细预览
                if (content.length > 500) {
                    console.log(`   ${content.substring(0, 500)}...`);
                } else {
                    console.log(`   ${content}`);
                }
            } else if (result.data && result.data.length > 0) {
                const content = result.data[0].content;
                console.log(`   响应内容 (${content.length}字符):`);
                if (content.length > 500) {
                    console.log(`   ${content.substring(0, 500)}...`);
                } else {
                    console.log(`   ${content}`);
                }
            }
            
            // 显示使用统计
            if (result.usage) {
                console.log(`   Token统计:`);
                console.log(`     总计: ${result.usage.total_tokens || 0}`);
                console.log(`     输入: ${result.usage.prompt_tokens || 0}`);
                console.log(`     输出: ${result.usage.completion_tokens || 0}`);
            }
            
            // 显示请求ID和完成状态
            if (result.id) {
                console.log(`   请求ID: ${result.id}`);
            }
            
            // 显示响应格式信息
            console.log(`   响应状态: 正常`);
            
            if (result.error) {
                throw new Error(`API错误: ${result.error.message}`);
            }
            
            return result;
            
        } catch (error) {
            // 只保留核心错误信息
            console.error('❌ AI请求失败:', error.message);
            
            if (error.name === 'TimeoutError') {
                throw new Error('AI服务请求超时');
            }
            if (error.name === 'TypeError' && error.message.includes('fetch')) {
                throw new Error('网络连接失败，请检查网络设置');
            }
            throw error;
        }
    }
    
    /**
     * 解析AI响应 - 增强版，支持错误恢复
     */
    parseAIResponse(response) {
        console.log('🔄 [parseAIResponse] 开始解析AI响应...');
        
        try {
            // 获取AI返回的内容
            let content;
            if (response.choices && response.choices.length > 0) {
                content = response.choices[0].message.content;
            } else if (response.data && response.data.length > 0) {
                content = response.data[0].content;
            } else {
                console.error('❌ [parseAIResponse] AI响应格式异常');
                return this.createEnhancedErrorResponse('AI响应格式异常', response);
            }
            
            console.log('📄 [parseAIResponse] 原始AI响应内容长度:', content.length);
            console.log('📄 [parseAIResponse] 原始内容预览:', content.substring(0, 300));
            
            // 提取JSON内容 - 增强版解析逻辑
            const extractedJSON = this.extractJSONContent(content);
            
            if (!extractedJSON.success) {
                console.warn('⚠️ [parseAIResponse] JSON提取失败，尝试内容重建');
                return this.rebuildResponseFromContent(content, response);
            }
            
            console.log('🔍 [parseAIResponse] 完整提取的JSON内容长度:', extractedJSON.jsonContent.length);
            console.log('🔍 [parseAIResponse] 完整JSON内容:', extractedJSON.jsonContent);
            
            let parsed;
            try {
                parsed = JSON.parse(extractedJSON.jsonContent);
            } catch (parseError) {
                console.error('❌ [parseAIResponse] JSON解析失败:', parseError);
                console.log('📄 [parseAIResponse] 尝试修复不完整的JSON...');
                
                // 尝试修复不完整的JSON
                const fixedJSON = this.tryFixIncompleteJSON(extractedJSON.jsonContent);
                if (fixedJSON) {
                    try {
                        parsed = JSON.parse(fixedJSON);
                        console.log('✅ [parseAIResponse] JSON修复成功');
                    } catch (fixError) {
                        console.error('❌ [parseAIResponse] JSON修复失败:', fixError);
                        return this.rebuildResponseFromContent(content, response);
                    }
                } else {
                    return this.rebuildResponseFromContent(content, response);
                }
            }
            
            // 添加响应元数据
            parsed.metadata = {
                model: response.model || this.modelName,
                tokens_used: response.usage?.total_tokens || 0,
                timestamp: new Date().toISOString(),
                request_id: response.id || Date.now().toString(),
                extraction_method: extractedJSON.method
            };
            
            console.log('✅ [parseAIResponse] 响应解析成功');
            return parsed;
            
        } catch (error) {
            console.error('❌ [parseAIResponse] 解析过程发生错误:', error);
            return this.createEnhancedErrorResponse(error.message, response);
        }
    }
    
    /**
     * 提取JSON内容的增强逻辑
     */
    extractJSONContent(content) {
        const methods = [
            '```json标记',
            '括号匹配',
            '正则表达式提取'
        ];
        
        // 方法1: 查找```json标记
        const jsonCodeBlockMatch = content.match(/```json\s*(\{[\s\S]*?\})\s*```/i);
        if (jsonCodeBlockMatch) {
            console.log('✅ [extractJSONContent] 使用```json标记提取JSON');
            return {
                success: true,
                jsonContent: jsonCodeBlockMatch[1],
                method: methods[0]
            };
        }
        
        // 方法2: 括号匹配
        const bracketResult = this.extractJSONByBrackets(content);
        if (bracketResult.success) {
            console.log('✅ [extractJSONContent] 使用括号匹配提取JSON');
            return {
                success: true,
                jsonContent: bracketResult.jsonContent,
                method: methods[1]
            };
        }
        
        // 方法3: 正则表达式提取
        const regexResult = this.extractJSONByRegex(content);
        if (regexResult.success) {
            console.log('✅ [extractJSONContent] 使用正则表达式提取JSON');
            return {
                success: true,
                jsonContent: regexResult.jsonContent,
                method: methods[2]
            };
        }
        
        console.log('❌ [extractJSONContent] 所有JSON提取方法都失败');
        return {
            success: false,
            jsonContent: '',
            method: '无'
        };
    }
    
    /**
     * 通过括号匹配提取JSON
     */
    extractJSONByBrackets(content) {
        let braceCount = 0;
        let firstBrace = -1;
        let lastBrace = -1;
        
        for (let i = 0; i < content.length; i++) {
            if (content[i] === '{') {
                if (firstBrace === -1) {
                    firstBrace = i;
                }
                braceCount++;
            } else if (content[i] === '}') {
                braceCount--;
                if (braceCount === 0) {
                    lastBrace = i;
                    break;
                }
            }
        }
        
        if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
            return {
                success: true,
                jsonContent: content.substring(firstBrace, lastBrace + 1)
            };
        }
        
        return { success: false, jsonContent: '' };
    }
    
    /**
     * 通过正则表达式提取JSON
     */
    extractJSONByRegex(content) {
        const jsonMatch = content.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            return {
                success: true,
                jsonContent: jsonMatch[0]
            };
        }
        return { success: false, jsonContent: '' };
    }
    
    /**
     * 尝试修复不完整的JSON
     */
    tryFixIncompleteJSON(jsonContent) {
        try {
            // 检查是否缺少闭合括号
            const openBraces = (jsonContent.match(/\{/g) || []).length;
            const closeBraces = (jsonContent.match(/\}/g) || []).length;
            
            if (openBraces > closeBraces) {
                const missingBraces = openBraces - closeBraces;
                const fixedJSON = jsonContent + '}'.repeat(missingBraces);
                console.log('🔧 [tryFixIncompleteJSON] 补全了', missingBraces, '个闭合括号');
                return fixedJSON;
            }
            
            // 尝试修复常见的JSON格式问题
            let fixed = jsonContent
                // 移除尾随逗号
                .replace(/,(\s*[}\]])/g, '$1')
                // 修复单引号为双引号（在value中）
                .replace(/:\s*'([^']*)'/g, ': "$1"')
                // 修复未加引号的键
                .replace(/([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:/g, '$1"$2":');
            
            return fixed;
            
        } catch (error) {
            console.warn('⚠️ [tryFixIncompleteJSON] JSON修复失败:', error);
            return null;
        }
    }
    
    /**
     * 从内容重建响应
     */
    rebuildResponseFromContent(content, originalResponse) {
        console.log('🔄 [rebuildResponseFromContent] 开始重建响应...');
        
        // 尝试从内容中提取公式信息
        const formulas = this.extractFormulasFromRawContent(content);
        const dataAnalysis = this.extractDataAnalysisFromRawContent(content);
        
        const rebuiltResponse = {
            formulas: formulas.length > 0 ? formulas : this.createFallbackFormulas(),
            data_analysis: dataAnalysis,
            alternative_formulas: [],
            reconstructed: true,
            reconstruction_note: '根据AI响应内容自动重建',
            metadata: {
                model: originalResponse.model || this.modelName,
                tokens_used: originalResponse.usage?.total_tokens || 0,
                timestamp: new Date().toISOString(),
                request_id: originalResponse.id || Date.now().toString(),
                reconstruction_method: 'content_analysis'
            }
        };
        
        console.log('✅ [rebuildResponseFromContent] 响应重建成功');
        return rebuiltResponse;
    }
    
    /**
     * 从原始内容提取公式
     */
    extractFormulasFromRawContent(content) {
        const formulas = [];
        const formulaRegex = /[=][A-Z]+\([^\)]*\)/g;
        const formulaMatches = content.match(formulaRegex);
        
        if (formulaMatches) {
            formulaMatches.forEach((formula, index) => {
                formulas.push({
                    title: `公式${index + 1}`,
                    formula: formula,
                    explanation: `从AI响应中提取的公式: ${formula}`,
                    confidence: 60,
                    applicable_ranges: ['需手动确定'],
                    required_functions: [formula.match(/[A-Z]+/)?.[0] || '未知'],
                    example: formula
                });
            });
        }
        
        return formulas;
    }
    
    /**
     * 从原始内容提取数据分析信息
     */
    extractDataAnalysisFromRawContent(content) {
        return {
            headers_found: [],
            data_types: ['待分析'],
            recommendations: ['建议重新提交请求以获得更精确的分析']
        };
    }
    
    /**
     * 创建备用公式
     */
    createFallbackFormulas() {
        return [
            {
                title: '基础求和公式',
                formula: '=SUM(选中范围)',
                explanation: '对指定范围内的数值进行求和计算',
                confidence: 50,
                applicable_ranges: ['数值列'],
                required_functions: ['SUM'],
                example: '=SUM(A1:A10)'
            }
        ];
    }
    
    /**
     * 创建增强错误响应
     */
    createEnhancedErrorResponse(errorMessage, originalResponse) {
        console.error('❌ [createEnhancedErrorResponse] 创建错误响应:', errorMessage);
        
        return {
            formulas: [],
            error: true,
            error_message: errorMessage,
            error_type: 'parse_error',
            fallback_response: this.createFallbackResponse(),
            metadata: {
                model: originalResponse.model || this.modelName,
                timestamp: new Date().toISOString(),
                error_timestamp: new Date().toISOString()
            }
        };
    }
    
    /**
     * 创建备用响应（当AI解析失败时）
     */
    createFallbackResponse() {
        return {
            formulas: [
                {
                    title: '基础求和公式',
                    formula: '=SUM(选中范围)',
                    explanation: '对指定范围内的数值进行求和计算',
                    confidence: 50,
                    applicable_ranges: ['数值列'],
                    required_functions: ['SUM'],
                    example: '=SUM(A1:A10)'
                }
            ],
            data_analysis: {
                headers_found: [],
                data_types: ['未分析'],
                recommendations: ['请检查数据源和需求描述']
            },
            alternative_formulas: []
        };
    }
    
    /**
     * 验证响应格式 - 增强版，支持错误恢复
     */
    validateResponse(response) {
        console.log('🔍 [validateResponse] 开始验证响应格式...');
        
        if (response.error) {
            console.log('⚠️ [validateResponse] 检测到错误响应，跳过格式验证');
            return; // 错误响应不需要验证格式
        }
        
        console.log('✅ [validateResponse] 开始结构验证');
        
        // 基本结构验证 - 增强版
        if (!response.formulas) {
            console.warn('⚠️ [validateResponse] 缺少formulas字段，尝试修复');
            response.formulas = this.createFallbackFormulas();
        }
        
        if (!Array.isArray(response.formulas)) {
            console.warn('⚠️ [validateResponse] formulas不是数组，尝试转换为数组');
            if (typeof response.formulas === 'object' && response.formulas !== null) {
                // 如果是对象，尝试提取为单个公式
                response.formulas = [response.formulas];
            } else {
                response.formulas = this.createFallbackFormulas();
            }
        }
        
        if (response.formulas.length === 0) {
            console.warn('⚠️ [validateResponse] formulas数组为空，使用备用公式');
            response.formulas = this.createFallbackFormulas();
        }
        
        console.log('📊 [validateResponse] 验证公式数量:', response.formulas.length);
        
        // 验证每个公式的结构 - 宽容模式
        const validatedFormulas = [];
        
        response.formulas.forEach((formula, index) => {
            try {
                const validatedFormula = this.validateAndFixFormula(formula, index);
                validatedFormulas.push(validatedFormula);
            } catch (validationError) {
                console.warn(`⚠️ [validateResponse] 第${index + 1}个公式验证失败，添加默认公式:`, validationError.message);
                validatedFormulas.push(this.createDefaultFormula(index));
            }
        });
        
        // 替换为验证后的公式
        response.formulas = validatedFormulas;
        
        // 验证其他可选字段 - 自动补全
        if (!response.data_analysis) {
            console.log('🔧 [validateResponse] 补全data_analysis字段');
            response.data_analysis = {
                headers_found: [],
                data_types: ['待分析'],
                recommendations: ['数据需要进一步分析']
            };
        } else {
            // 确保data_analysis有必要的字段
            if (!Array.isArray(response.data_analysis.headers_found)) {
                response.data_analysis.headers_found = [];
            }
            if (!Array.isArray(response.data_analysis.data_types)) {
                response.data_analysis.data_types = ['待分析'];
            }
            if (!Array.isArray(response.data_analysis.recommendations)) {
                response.data_analysis.recommendations = ['数据需要进一步分析'];
            }
        }
        
        if (!response.alternative_formulas) {
            console.log('🔧 [validateResponse] 补全alternative_formulas字段');
            response.alternative_formulas = [];
        }
        
        // 添加验证元数据
        if (!response.metadata) {
            response.metadata = {};
        }
        response.metadata.validation_status = 'passed';
        response.metadata.validation_timestamp = new Date().toISOString();
        response.metadata.formulas_count = response.formulas.length;
        
        console.log('✅ [validateResponse] 响应格式验证完成');
    }
    
    /**
     * 验证并修复单个公式
     */
    validateAndFixFormula(formula, index) {
        if (!formula || typeof formula !== 'object') {
            throw new Error('公式不是有效对象');
        }
        
        const validated = { ...formula };
        
        // 验证和修复必要字段
        if (!validated.title) {
            validated.title = `公式${index + 1}`;
            console.log(`🔧 [validateAndFixFormula] 为第${index + 1}个公式补全title字段`);
        }
        
        if (!validated.formula) {
            if (validated.content) {
                // 从content字段尝试提取公式
                const formulaMatch = validated.content.match(/[=][A-Z]+\([^\)]*\)/);
                if (formulaMatch) {
                    validated.formula = formulaMatch[0];
                } else {
                    validated.formula = '=0'; // 默认公式
                }
            } else {
                validated.formula = '=0'; // 默认公式
            }
            console.log(`🔧 [validateAndFixFormula] 为第${index + 1}个公式补全formula字段`);
        }
        
        if (!validated.explanation) {
            validated.explanation = `公式说明: ${validated.formula}`;
            console.log(`🔧 [validateAndFixFormula] 为第${index + 1}个公式补全explanation字段`);
        }
        
        // 验证和修复置信度
        if (typeof validated.confidence !== 'number' || isNaN(validated.confidence)) {
            validated.confidence = 70; // 默认置信度
            console.log(`🔧 [validateAndFixFormula] 为第${index + 1}个公式设置默认置信度`);
        } else if (validated.confidence < 0) {
            validated.confidence = 0;
        } else if (validated.confidence > 100) {
            validated.confidence = 100;
        }
        
        // 验证和修复可选字段
        if (!Array.isArray(validated.applicable_ranges)) {
            validated.applicable_ranges = ['需手动确定'];
        }
        
        if (!Array.isArray(validated.required_functions)) {
            // 尝试从公式中提取函数名
            const functionMatch = validated.formula.match(/[A-Z]+/);
            validated.required_functions = functionMatch ? [functionMatch[0]] : ['未知'];
        }
        
        if (!validated.example) {
            validated.example = validated.formula;
        }
        
        console.log(`✅ [validateAndFixFormula] 第${index + 1}个公式验证成功`);
        return validated;
    }
    
    /**
     * 创建默认公式
     */
    createDefaultFormula(index) {
        return {
            title: `默认公式${index + 1}`,
            formula: '=SUM(选中范围)',
            explanation: '基础求和公式，请根据实际需求调整',
            confidence: 50,
            applicable_ranges: ['数值列'],
            required_functions: ['SUM'],
            example: '=SUM(A1:A10)',
            is_default: true
        };
    }
    
    /**
     * 测试AI连接
     */
    async testConnection() {
        try {
            // AI连接诊断测试日志已简化
            // console.group('🧪 AI连接诊断测试');
            // console.log('🔍 诊断时间:', new Date().toLocaleString());
            // console.log('📡 目标API端点:', this.apiEndpoint);
            // console.log('🔑 API密钥长度:', this.apiKey ? this.apiKey.length : 0);
            // console.log('📝 使用模型:', this.modelName);
            
            const testRequest = {
                model: this.modelName,
                messages: [
                    {
                        role: 'user',
                        content: '请回复"连接测试成功"'
                    }
                ],
                max_tokens: 10
            };
            
            // console.log('📤 发送测试请求:', JSON.stringify(testRequest, null, 2));
            
            const response = await this.sendAIRequest(testRequest);
            
            // console.log('📥 收到响应:', JSON.stringify(response, null, 2));
            // console.groupEnd();
            
            return {
                success: true,
                message: 'AI连接测试成功',
                response: response
            };
            
        } catch (error) {
            // console.error('❌ AI连接测试失败:', error);
            // console.groupEnd();
            
            return {
                success: false,
                message: `AI连接测试失败: ${error.message}`,
                error: error
            };
        }
    }
    
    /**
     * 诊断API问题
     */
    async diagnoseAPI() {
        // API诊断报告日志已简化
        // console.group('🔧 API诊断报告');
        // console.log('📊 诊断时间:', new Date().toLocaleString());
        
        // 检查配置
        const configCheck = {
            apiEndpoint: {
                value: this.apiEndpoint,
                valid: this.apiEndpoint && this.apiEndpoint.startsWith('http'),
                message: this.apiEndpoint ? 'API端点已配置' : 'API端点未配置'
            },
            apiKey: {
                value: this.apiKey ? `${this.apiKey.substring(0, 10)}...` : null,
                valid: this.apiKey && this.apiKey.length > 10,
                message: this.apiKey ? 'API密钥已配置' : 'API密钥未配置'
            },
            modelName: {
                value: this.modelName,
                valid: this.modelName && this.modelName.length > 0,
                message: this.modelName ? '模型名称已配置' : '模型名称未配置'
            }
        };
        
        // console.log('⚙️ 配置检查:', configCheck);
        
        // 检查网络状态
        const networkCheck = {
            online: navigator.onLine,
            userAgent: navigator.userAgent.substring(0, 50),
            hasFetch: typeof fetch !== 'undefined',
            hasURL: typeof URL !== 'undefined'
        };
        
        // console.log('🌐 网络检查:', networkCheck);
        
        // CORS预检请求测试
        if (this.apiEndpoint) {
            try {
                // console.log('🔍 正在进行CORS预检请求...');
                const preflightResponse = await fetch(this.apiEndpoint, {
                    method: 'OPTIONS',
                    mode: 'cors',
                    headers: {
                        'Origin': window.location.origin,
                        'Access-Control-Request-Method': 'POST',
                        'Access-Control-Request-Headers': 'content-type,authorization'
                    }
                });
                // console.log('✅ CORS预检成功:', preflightResponse.status, preflightResponse.statusText);
                
                const corsHeaders = {
                    'access-control-allow-origin': preflightResponse.headers.get('access-control-allow-origin'),
                    'access-control-allow-methods': preflightResponse.headers.get('access-control-allow-methods'),
                    'access-control-allow-headers': preflightResponse.headers.get('access-control-allow-headers')
                };
                // console.log('🌍 CORS头信息:', corsHeaders);
                
            } catch (corsError) {
                // console.error('❌ CORS预检失败:', corsError.message);
            }
        }
        
        // console.groupEnd();
        
        // 返回诊断结果
        return {
            timestamp: new Date().toISOString(),
            config: configCheck,
            network: networkCheck,
            recommendations: this.generateRecommendations(configCheck, networkCheck)
        };
    }
    
    /**
     * 生成修复建议
     */
    generateRecommendations(configCheck, networkCheck) {
        const recommendations = [];
        
        if (!configCheck.apiEndpoint.valid) {
            recommendations.push('❌ 请检查API端点配置，确保以http://或https://开头');
        }
        
        if (!configCheck.apiKey.valid) {
            recommendations.push('❌ 请检查API密钥配置，确保密钥有效且未过期');
        }
        
        if (!configCheck.modelName.valid) {
            recommendations.push('❌ 请检查模型名称配置');
        }
        
        if (!networkCheck.online) {
            recommendations.push('❌ 请检查网络连接状态');
        }
        
        if (!networkCheck.hasFetch) {
            recommendations.push('❌ 当前浏览器不支持Fetch API，请升级浏览器');
        }
        
        if (recommendations.length === 0) {
            recommendations.push('✅ 基本配置检查通过，如果仍有问题，请检查API服务商状态');
        }
        
        return recommendations;
    }
     
    /**
     * 获取配置信息（用于UI显示）
     */
    getConfig() {
        return {
            apiEndpoint: this.apiEndpoint,
            modelName: this.modelName,
            timeout: this.timeout,
            isConfigured: !!(this.apiKey && this.apiEndpoint)
        };
    }
    
    /**
     * 更新API密钥
     */
    updateApiKey(newApiKey) {
        this.apiKey = newApiKey;
        this.saveConfig();
    }
    
    /**
     * 保存配置到本地存储
     */
    saveConfig() {
        try {
            const config = {
                apiEndpoint: this.apiEndpoint,
                apiKey: this.apiKey,
                modelName: this.modelName
            };
            
            localStorage.setItem('aiHelper_config', JSON.stringify(config));
        } catch (error) {
            console.warn('保存AI配置失败:', error);
        }
    }
    
    /**
     * 从本地存储加载配置
     */
    loadSavedConfig() {
        try {
            const saved = localStorage.getItem('aiHelper_config');
            if (saved) {
                const config = JSON.parse(saved);
                this.updateConfig(config);
                return true;
            }
        } catch (error) {
            console.warn('加载保存的AI配置失败:', error);
        }
        return false;
    }
    
    /**
     * 获取使用统计
     */
    getUsageStats() {
        const stats = {
            total_requests: parseInt(localStorage.getItem('aiHelper_requests') || '0'),
            successful_requests: parseInt(localStorage.getItem('aiHelper_success') || '0'),
            failed_requests: parseInt(localStorage.getItem('aiHelper_failed') || '0'),
            total_tokens: parseInt(localStorage.getItem('aiHelper_tokens') || '0')
        };
        
        stats.success_rate = stats.total_requests > 0 ? 
            Math.round((stats.successful_requests / stats.total_requests) * 100) : 0;
        
        return stats;
    }
    
    /**
     * 记录使用统计
     */
    recordUsage(isSuccess, tokensUsed = 0) {
        try {
            const totalRequests = parseInt(localStorage.getItem('aiHelper_requests') || '0') + 1;
            const successRequests = parseInt(localStorage.getItem('aiHelper_success') || '0') + (isSuccess ? 1 : 0);
            const failedRequests = parseInt(localStorage.getItem('aiHelper_failed') || '0') + (isSuccess ? 0 : 1);
            const totalTokens = parseInt(localStorage.getItem('aiHelper_tokens') || '0') + tokensUsed;
            
            localStorage.setItem('aiHelper_requests', totalRequests.toString());
            localStorage.setItem('aiHelper_success', successRequests.toString());
            localStorage.setItem('aiHelper_failed', failedRequests.toString());
            localStorage.setItem('aiHelper_tokens', totalTokens.toString());
            
        } catch (error) {
            // 使用统计记录失败日志已简化
            // console.warn('记录使用统计失败:', error);
        }
    }

    /**
     * 从响应内容重建JSON结构
     */
    rebuildJSONFromResponse(content) {
        try {
            console.log('🔄 [rebuildJSONFromResponse] 开始重建JSON结构...');
            
            // 尝试提取关键信息并重建最小JSON结构
            const formulas = this.extractFormulasFromContent(content);
            const dataAnalysis = this.extractDataAnalysisFromContent(content);
            const alternatives = this.extractAlternativesFromContent(content);
            
            if (formulas.length === 0) {
                console.log('⚠️ [rebuildJSONFromResponse] 未找到任何公式信息');
                return null;
            }
            
            const rebuilt = {
                formulas: formulas,
                data_analysis: dataAnalysis,
                alternative_formulas: alternatives,
                reconstructed: true,
                reconstruction_note: 'JSON结构根据AI响应内容自动重建'
            };
            
            console.log('✅ [rebuildJSONFromResponse] JSON重建成功');
            return JSON.stringify(rebuilt, null, 2);
            
        } catch (error) {
            console.error('❌ [rebuildJSONFromResponse] JSON重建失败:', error.message);
            return null;
        }
    }

    /**
     * 从内容中提取公式信息
     */
    extractFormulasFromContent(content) {
        const formulas = [];
        
        // 尝试匹配不同的公式格式
        const patterns = [
            // 匹配title和formula的结构
            /"title"\s*:\s*"([^"]+)"[^{}]*"formula"\s*:\s*"([^"]+)"/gi,
            // 匹配公式名称和公式的结构
            /(?:公式名称|标题|描述)\s*[:：]\s*([^\n]+)[^:]*?(?:公式|表达式)\s*[:：]\s*([^\n]+)/gi,
            // 匹配等号开头的公式
            /=\s*[^=\n]+/g
        ];
        
        for (const pattern of patterns) {
            let match;
            while ((match = pattern.exec(content)) !== null && formulas.length < 3) {
                if (match.length >= 3) {
                    // 处理JSON格式匹配
                    const title = match[1].trim();
                    const formula = match[2].trim();
                    if (title && formula) {
                        formulas.push({
                            title: title,
                            formula: formula.startsWith('=') ? formula : '=' + formula,
                            explanation: '根据AI响应内容重建的公式',
                            confidence: 75,
                            applicable_ranges: ['自动检测'],
                            required_functions: [this.extractFunctionName(formula)],
                            example: formula
                        });
                    }
                } else if (match.length === 1 && match[0].startsWith('=')) {
                    // 处理简单公式匹配
                    const formula = match[0].trim();
                    formulas.push({
                        title: '重建公式',
                        formula: formula,
                        explanation: '从AI响应中提取的公式',
                        confidence: 70,
                        applicable_ranges: ['自动检测'],
                        required_functions: [this.extractFunctionName(formula)],
                        example: formula
                    });
                }
            }
        }
        
        // 如果没有找到任何公式，创建一个默认公式
        if (formulas.length === 0) {
            formulas.push({
                title: '默认公式',
                formula: '=SUM(A1:A10)',
                explanation: '基于AI响应创建的默认求和公式',
                confidence: 50,
                applicable_ranges: ['数值列'],
                required_functions: ['SUM'],
                example: '=SUM(A1:A10)'
            });
        }
        
        return formulas;
    }

    /**
     * 从内容中提取数据分析信息
     */
    extractDataAnalysisFromContent(content) {
        const analysis = {
            headers_found: [],
            data_types: [],
            recommendations: ['基于AI响应内容自动分析']
        };
        
        // 尝试提取表头信息
        const headerMatches = content.match(/(?:表头|列名|字段)[：:]\s*([^\n,，]+(?:\s*[，,]\s*[^\n,，]+)*)/gi);
        if (headerMatches) {
            headerMatches.forEach(match => {
                const headers = match.replace(/(?:表头|列名|字段)[：:]\s*/gi, '').split(/[，,]/);
                headers.forEach(header => {
                    const trimmed = header.trim();
                    if (trimmed) {
                        analysis.headers_found.push(trimmed);
                    }
                });
            });
        }
        
        // 尝试提取数据类型
        const typeMatches = content.match(/(?:数据类型|类型)[：:]\s*([^\n,，]+)/gi);
        if (typeMatches) {
            typeMatches.forEach(match => {
                const type = match.replace(/(?:数据类型|类型)[：:]\s*/gi, '').trim();
                if (type) {
                    analysis.data_types.push(type);
                }
            });
        }
        
        return analysis;
    }

    /**
     * 从内容中提取替代公式
     */
    extractAlternativesFromContent(content) {
        const alternatives = [];
        
        // 尝试匹配替代方案
        const altPattern = /(?:替代方案|备选|其他方案)[：:]([^\n]+)(?:公式|表达式)[：:]([^\n]+)/gi;
        let match;
        
        while ((match = altPattern.exec(content)) !== null && alternatives.length < 2) {
            const desc = match[1].trim();
            const formula = match[2].trim();
            if (desc && formula) {
                alternatives.push({
                    description: desc,
                    formula: formula.startsWith('=') ? formula : '=' + formula,
                    pros: ['替代方案'],
                    cons: ['可能不是最佳选择']
                });
            }
        }
        
        return alternatives;
    }

    /**
     * 从公式中提取函数名
     */
    extractFunctionName(formula) {
        const functionPattern = /^=\s*([A-Z]+)\s*\(/i;
        const match = formula.match(functionPattern);
        return match ? match[1].toUpperCase() : 'UNKNOWN';
    }
}

// 创建全局实例
window.aiInterface = new AIInterface();

// 导出类供其他模块使用
window.AIInterface = AIInterface;