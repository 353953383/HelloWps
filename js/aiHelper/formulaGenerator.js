/**
 * 智能公式生成器 - 主逻辑
 * 负责处理用户交互、公式生成和数据处理
 */

class FormulaGenerator {
    constructor() {
        this.currentCell = null;
        this.selectedWorkbooks = [];
        this.selectedWorksheets = [];
        this.referenceType = 'current';
        this.formulaDescription = '';
        this.fillRight = false;
        this.fillDown = false;
        this.selectedFormula = null;
        
        // 防重复调用状态
        this.isGenerating = false;
        
        // 初始化AI接口引用
        this.aiInterface = window.aiInterface || null;
        if (!this.aiInterface) {
            console.warn('⚠️ [FormulaGenerator] AI接口未初始化，尝试延迟初始化...');
            // 延迟初始化，确保DOM加载完成
            setTimeout(() => {
                this.aiInterface = window.aiInterface || null;
                if (!this.aiInterface) {
                    console.error('❌ [FormulaGenerator] AI接口初始化失败');
                } else {
                    console.log('✅ [FormulaGenerator] AI接口初始化成功');
                }
            }, 100);
        } else {
            console.log('✅ [FormulaGenerator] AI接口初始化成功');
        }
        
        this.init();
    }
    
    init() {
        this.bindEvents();
        this.updateCurrentCell();
        this.loadWorkbookData();
    }
    
    bindEvents() {
        // 引用类型切换
        document.querySelectorAll('input[name="referenceType"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.referenceType = e.target.value;
                this.toggleReferenceSelection();
            });
        });
        
        // 填充方向设置
        document.getElementById('fillRight').addEventListener('change', (e) => {
            this.fillRight = e.target.checked;
        });
        
        document.getElementById('fillDown').addEventListener('change', (e) => {
            this.fillDown = e.target.checked;
        });
        
        // 公式描述输入
        document.getElementById('formulaDescription').addEventListener('input', (e) => {
            this.formulaDescription = e.target.value;
        });
        
        // 生成按钮
        document.getElementById('generateFormula').addEventListener('click', () => {
            this.generateFormula();
        });
        
        // 刷新工作簿按钮
        document.getElementById('refreshWorkbooks').addEventListener('click', () => {
            this.loadWorkbookData();
        });
        
        // 应用公式按钮
        document.getElementById('applyFormula').addEventListener('click', () => {
            this.applySelectedFormula();
        });
        
        // 清空按钮
        document.getElementById('clearAll').addEventListener('click', () => {
            this.clearAll();
        });
        
        // 工作簿搜索
        document.getElementById('workbookSearch').addEventListener('input', (e) => {
            this.filterWorkbooks(e.target.value);
        });
    }
    
    updateCurrentCell() {
        try {
            // 获取当前选中的单元格信息
            if (window.Application && window.Application.ActiveSheet) {
                const activeSheet = window.Application.ActiveSheet;
                const selection = window.Application.Selection;
                
                if (selection) {
                    this.currentCell = {
                        workbook: window.Application.ActiveWorkbook ? window.Application.ActiveWorkbook.Name : '',
                        worksheet: activeSheet ? activeSheet.Name : '',
                        row: selection.Row || 1,
                        col: selection.Column || 1,
                        cellAddress: this.getCellAddress(selection.Row || 1, selection.Column || 1)
                    };
                    
                    document.getElementById('currentCell').textContent = 
                        `${this.currentCell.cellAddress} (${this.currentCell.worksheet})`;
                }
            }
        } catch (error) {
            console.error('获取当前单元格信息失败:', error);
            document.getElementById('currentCell').textContent = '无法获取';
        }
    }
    
    getCellAddress(row, col) {
        const columnLetters = this.getColumnLetter(col);
        return `${columnLetters}${row}`;
    }
    
    getColumnLetter(col) {
        let temp = '';
        let columnNumber = col;
        
        while (columnNumber > 0) {
            let remainder = (columnNumber - 1) % 26;
            temp = String.fromCharCode(65 + remainder) + temp;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        
        return temp;
    }
    
    loadWorkbookData() {
        try {
            // 更新状态指示器
            this.updateStatus('正在加载工作簿...');
            
            // 获取所有工作簿
            if (window.Application && window.Application.Workbooks) {
                const workbooks = window.Application.Workbooks;
                const workbookData = [];
                
                for (let i = 1; i <= workbooks.Count; i++) {
                    const wb = workbooks.Item(i);
                    const worksheets = [];
                    
                    if (wb.Worksheets) {
                        for (let j = 1; j <= wb.Worksheets.Count; j++) {
                            const ws = wb.Worksheets.Item(j);
                            worksheets.push({
                                name: ws.Name,
                                usedRange: this.getUsedRangeInfo(ws)
                            });
                        }
                    }
                    
                    workbookData.push({
                        name: wb.Name,
                        path: wb.Path || '',
                        worksheets: worksheets
                    });
                }
                
                this.updateWorkbookList(workbookData);
                this.updateStatus('工作簿加载完成');
            }
        } catch (error) {
            console.error('加载工作簿数据失败:', error);
            this.updateStatus('加载工作簿失败');
            this.showNotification('加载工作簿数据失败: ' + error.message, 'error');
        }
    }
    
    getUsedRangeInfo(worksheet) {
        try {
            const usedRange = worksheet.UsedRange;
            if (usedRange) {
                return {
                    rows: usedRange.Rows.Count,
                    columns: usedRange.Columns.Count,
                    startRow: usedRange.Row,
                    startCol: usedRange.Column,
                    endRow: usedRange.Row + usedRange.Rows.Count - 1,
                    endCol: usedRange.Column + usedRange.Columns.Count - 1
                };
            }
        } catch (error) {
            console.warn('获取工作表使用范围失败:', error);
        }
        
        return { rows: 0, columns: 0, startRow: 1, startCol: 1, endRow: 1, endCol: 1 };
    }
    
    updateWorkbookList(workbooks) {
        const workbookList = document.getElementById('workbookList');
        const worksheetList = document.getElementById('worksheetList');
        
        // 清空现有内容
        workbookList.innerHTML = '';
        worksheetList.innerHTML = '';
        
        // 生成工作簿列表
        workbooks.forEach(workbook => {
            const item = document.createElement('div');
            item.className = 'workbook-item';
            item.innerHTML = `
                <span>${workbook.name}</span>
                <span class="count">(${workbook.worksheets.length}个工作表)</span>
            `;
            item.addEventListener('click', () => {
                this.selectWorkbook(workbook);
                this.updateWorksheetList(workbook.worksheets);
            });
            
            workbookList.appendChild(item);
        });
        
        // 默认选择第一个工作簿
        if (workbooks.length > 0) {
            this.selectWorkbook(workbooks[0]);
            this.updateWorksheetList(workbooks[0].worksheets);
        }
    }
    
    selectWorkbook(workbook) {
        // 移除其他选中状态
        document.querySelectorAll('.workbook-item').forEach(item => {
            item.classList.remove('selected');
        });
        
        // 添加选中状态到当前项
        const items = document.querySelectorAll('.workbook-item');
        items.forEach((item, index) => {
            if (item.textContent.includes(workbook.name)) {
                item.classList.add('selected');
            }
        });
        
        this.selectedWorkbooks = [workbook];
        this.updateSelectedSources();
    }
    
    updateWorksheetList(worksheets) {
        const worksheetList = document.getElementById('worksheetList');
        worksheetList.innerHTML = '';
        
        worksheets.forEach(worksheet => {
            const item = document.createElement('div');
            item.className = 'worksheet-item';
            item.innerHTML = `
                <span>${worksheet.name}</span>
                <span class="range">(${worksheet.usedRange.rows}x${worksheet.usedRange.columns})</span>
            `;
            item.addEventListener('click', () => {
                this.toggleWorksheetSelection(worksheet);
            });
            
            worksheetList.appendChild(item);
        });
    }
    
    toggleWorksheetSelection(worksheet) {
        const items = document.querySelectorAll('.worksheet-item');
        items.forEach(item => {
            if (item.textContent.includes(worksheet.name)) {
                item.classList.toggle('selected');
            }
        });
        
        // 更新选中列表
        this.selectedWorksheets = this.selectedWorksheets.filter(ws => ws.name !== worksheet.name);
        if (!this.selectedWorksheets.find(ws => ws.name === worksheet.name)) {
            this.selectedWorksheets.push(worksheet);
        }
        
        this.updateSelectedSources();
    }
    
    updateSelectedSources() {
        const totalSelected = this.selectedWorkbooks.length + this.selectedWorksheets.length;
        document.getElementById('selectedSources').textContent = totalSelected;
    }
    
    toggleReferenceSelection() {
        const selectionArea = document.getElementById('worksheetSelection');
        
        if (this.referenceType === 'current') {
            selectionArea.style.display = 'none';
        } else {
            selectionArea.style.display = 'block';
        }
    }
    
    async generateFormula() {
        // 防止重复调用的保护机制
        if (this.isGenerating) {
            console.log('⚠️ [调试] 公式生成已在进行中，跳过重复调用');
            return;
        }
        this.isGenerating = true;
        console.log('🔄 [调试] 开始公式生成流程');
        
        // 增强错误处理：捕获所有可能的异常
        try {
            console.log('🔍 [调试] 开始generateFormula方法');
            console.log('📍 [调试] 调用堆栈信息:');
            console.trace();
            
            // 收集请求数据的增强错误处理
            let requestData;
            try {
                console.log('🔍 [调试] 开始收集请求数据');
                requestData = this.collectRequestData();
                console.log('✅ [调试] 请求数据收集成功');
            } catch (dataCollectionError) {
                console.error('❌ [调试] 请求数据收集失败:', dataCollectionError);
                console.error('📊 [调试] 数据收集错误堆栈:', dataCollectionError.stack);
                console.error('📊 [调试] 当前Excel对象状态检查:', {
                    'window.Application存在': !!window.Application,
                    'ActiveWorkbook存在': !!window.Application?.ActiveWorkbook,
                    'ActiveSheet存在': !!window.Application?.ActiveSheet,
                    'Selection存在': !!window.Application?.Selection
                });
                throw new Error(`数据收集失败: ${dataCollectionError.message}`);
            }
            
            // 添加详细的数据验证日志
            console.log('🔍 [调试] 数据收集完成，开始验证');
            console.log('📊 [调试] requestData详情:');
            console.log('  - description:', requestData.description || '未设置');
            console.log('  - currentCell:', requestData.currentCell ? JSON.stringify(requestData.currentCell) : 'null/undefined');
            console.log('  - columnHeaders:', requestData.columnHeaders ? `${requestData.columnHeaders.length}项` : 'null/undefined');
            console.log('  - allWorksheets:', requestData.allWorksheets ? `${requestData.allWorksheets.length}项` : 'null/undefined');
            
            // 特别详细地显示description内容（因为这是最关键的）
            if (requestData.description) {
                console.log('📝 [调试] description完整内容:');
                console.log('========== DESCRIPTION START ==========');
                console.log(requestData.description);
                console.log('========== DESCRIPTION END ==========');
                console.log(`📏 [调试] description长度: ${requestData.description.length} 字符`);
                
                // 检查是否包含关键信息
                const keyTerms = ['组件在制', '库存', 'Excel公式', '单元格'];
                keyTerms.forEach(term => {
                    const contains = requestData.description.includes(term);
                    console.log(`🔍 [调试] 关键词"${term}": ${contains ? '✅ 包含' : '❌ 缺失'}`);
                });
            } else {
                console.warn('⚠️ [调试] 警告: description为空');
            }
            
            // 如果关键数据缺失，显示具体问题
            if (!requestData.description) {
                console.warn('⚠️ [调试] 警告: description为空');
            }
            if (!requestData.currentCell) {
                console.warn('⚠️ [调试] 警告: currentCell为null/undefined');
            }
            if (!requestData.columnHeaders || requestData.columnHeaders.length === 0) {
                console.warn('⚠️ [调试] 警告: columnHeaders为空或未设置');
            }
            if (!requestData.allWorksheets || requestData.allWorksheets.length === 0) {
                console.warn('⚠️ [调试] 警告: allWorksheets为空或未设置');
            }
            
            // 显示传递给AI的完整数据
            console.log('📤 [调试] 传递给AIInterface的完整requestData:');
            console.log(JSON.stringify(requestData, null, 2));
            
            // 检查是否有AI接口实例
            if (!this.aiInterface) {
                console.error('❌ [调试] AI接口未初始化');
                console.error('📊 [调试] this对象检查:', {
                    'this存在': !!this,
                    'this.aiInterface存在': !!this?.aiInterface,
                    'AI接口类型': typeof this?.aiInterface
                });
                throw new Error('AI接口未初始化');
            }
            
            // AI接口调用的增强错误处理和详细调试日志
            let result;
            try {
                console.log('🔄 [调试] 开始调用AIInterface.generateFormulaWithEndpoint');
                console.log('🔍 [调试] AI接口调用堆栈:');
                console.trace();
                
                // 详细的API调用前检查和日志记录
                console.log('📡 [调试] API调用前的完整检查:');
                console.log('  📋 [调试] 请求数据大小:', JSON.stringify(requestData).length, '字符');
                console.log('  📋 [调试] 请求数据字段:', Object.keys(requestData));
                console.log('  📋 [调试] description字段长度:', requestData.description?.length || 0);
                console.log('  📋 [调试] 当前单元格地址:', requestData.currentCell?.cellAddress || '无');
                console.log('  📋 [调试] 选中工作簿数量:', requestData.selectedWorkbooks?.length || 0);
                console.log('  📋 [调试] 列标题数量:', requestData.columnHeaders?.length || 0);
                
                // API配置检查
                console.log('⚙️ [调试] AI接口配置检查:');
                console.log('  - API端点存在:', !!this.aiInterface?.config?.apiEndpoint);
                console.log('  - API端点地址:', this.aiInterface?.config?.apiEndpoint || '未设置');
                console.log('  - API密钥存在:', !!this.aiInterface?.config?.apiKey);
                console.log('  - API密钥长度:', this.aiInterface?.config?.apiKey?.length || 0);
                console.log('  - 模型名称:', this.aiInterface?.config?.model || '未设置');
                console.log('  - 超时设置:', this.aiInterface?.config?.timeout || '默认');
                console.log('  - 最大Token数:', this.aiInterface?.config?.maxTokens || '默认');
                
                // 记录API请求开始时间
                const startTime = Date.now();
                console.log('⏰ [调试] API调用开始时间:', new Date(startTime).toLocaleString());
                
                // 执行API调用
                console.log('🚀 [调试] 正在发送API请求...');
                result = await this.aiInterface.generateFormulaWithEndpoint(requestData);
                
                // 计算调用耗时
                const endTime = Date.now();
                const duration = endTime - startTime;
                console.log('⏰ [调试] API调用结束时间:', new Date(endTime).toLocaleString());
                console.log('⏱️ [调试] API调用总耗时:', duration, 'ms');
                
                console.log('✅ [调试] AI接口调用成功');
                
                // 详细的API响应日志记录
                console.log('📡 [调试] API响应详细分析:');
                console.log('  📋 [调试] 响应数据类型:', typeof result);
                console.log('  📋 [调试] 响应是否为对象:', result && typeof result === 'object');
                
                if (result) {
                    console.log('  📋 [调试] 响应字段:', Object.keys(result));
                    
                    if (result.formulas && Array.isArray(result.formulas)) {
                        console.log('  🎯 [调试] 公式数量:', result.formulas.length);
                        result.formulas.forEach((formula, index) => {
                            console.log(`    公式${index + 1}:`);
                            console.log(`      - 标题: ${formula.title || '无标题'}`);
                            console.log(`      - 公式: ${formula.formula || '无公式'}`);
                            console.log(`      - 可信度: ${formula.confidence || 0}%`);
                            console.log(`      - 描述: ${formula.description || '无描述'}`);
                        });
                    }
                    
                    if (result.usage) {
                        console.log('  📊 [调试] API使用统计:');
                        console.log(`    - Prompt Token数: ${result.usage.prompt_tokens || 0}`);
                        console.log(`    - Completion Token数: ${result.usage.completion_tokens || 0}`);
                        console.log(`    - Total Token数: ${result.usage.total_tokens || 0}`);
                    }
                    
                    if (result.model) {
                        console.log(`  🤖 [调试] 使用的模型: ${result.model}`);
                    }
                    
                    if (result.response_time) {
                        console.log(`  ⏱️ [调试] 模型响应时间: ${result.response_time}ms`);
                    }
                } else {
                    console.warn('  ⚠️ [调试] API响应为空');
                }
                
            } catch (aiCallError) {
                // 计算错误发生时的耗时
                const errorTime = Date.now();
                console.log('⏰ [调试] API错误发生时间:', new Date(errorTime).toLocaleString());
                
                console.error('❌ [调试] AI接口调用失败:', aiCallError);
                console.error('📊 [调试] AI接口错误详情:', {
                    '错误消息': aiCallError.message,
                    '错误类型': aiCallError.constructor.name,
                    '错误堆栈': aiCallError.stack,
                    'API端点': this.aiInterface?.config?.apiEndpoint || '未设置',
                    'API密钥存在': !!this.aiInterface?.config?.apiKey,
                    'API密钥长度': this.aiInterface?.config?.apiKey?.length || 0,
                    '模型名称': this.aiInterface?.config?.model || '未设置',
                    '错误发生时间': new Date(errorTime).toLocaleString()
                });
                
                // 如果错误对象包含更多信息，记录下来
                if (aiCallError.response) {
                    console.error('📡 [调试] API响应错误详情:');
                    console.error('  - HTTP状态码:', aiCallError.response.status);
                    console.error('  - 状态文本:', aiCallError.response.statusText);
                    console.error('  - 响应头:', aiCallError.response.headers);
                }
                
                if (aiCallError.request) {
                    console.error('📡 [调试] API请求详情:');
                    console.error('  - 请求URL:', aiCallError.request?.url || '未知');
                    console.error('  - 请求方法:', aiCallError.request?.method || '未知');
                }
                
                // 重新包装错误信息
                const errorMessage = `AI接口调用失败 (${aiCallError.constructor.name}): ${aiCallError.message}`;
                console.error('📋 [调试] 包装后的错误信息:', errorMessage);
                throw new Error(errorMessage);
            }
            
            // ✅ [调试] 在界面上显示AI返回的公式建议
            if (result && result.formulas && result.formulas.length > 0) {
                console.log('🎯 [调试] 准备在界面显示公式建议:', result.formulas);
                this.showFormulaSuggestions(result.formulas);
                this.showNotification('AI公式建议生成成功！请选择要应用的公式。', 'success');
                
                // 记录显示的公式信息
                console.log('📋 [调试] 公式建议显示完成，包含', result.formulas.length, '个公式选项');
                result.formulas.forEach((formula, index) => {
                    console.log(`  公式${index + 1}: ${formula.title} - ${formula.formula} (可信度: ${formula.confidence}%)`);
                });
            } else {
                console.warn('⚠️ [调试] AI返回的响应中没有找到formulas数组:', result);
                this.showNotification('AI返回的响应格式异常，无法显示公式建议', 'error');
            }
            
            console.log('✅ [调试] generateFormula方法完成');
            return result;
            
        } catch (error) {
            console.error('❌ [调试] generateFormula方法出错:', error);
            console.error('📊 [调试] 错误堆栈:', error.stack);
            throw error;
        } finally {
            // 确保在任何情况下都重置生成状态
            this.isGenerating = false;
            console.log('🔄 [调试] 重置生成状态');
        }
    }
    
    /**
     * 获取当前工作簿的完整信息
     */
    getCurrentWorkbookInfo() {
        try {
            const activeWorkbook = window.Application.ActiveWorkbook;
            const activeSheet = window.Application.ActiveSheet;
            const selection = window.Application.Selection;
            
            // 获取当前工作簿信息
            const currentWorkbook = {
                name: activeWorkbook ? activeWorkbook.Name : '未知工作簿'
            };
            
            // 获取所有工作表信息
            const allWorksheets = [];
            if (activeWorkbook && activeWorkbook.Worksheets) {
                for (let i = 1; i <= activeWorkbook.Worksheets.Count; i++) {
                    try {
                        const ws = activeWorkbook.Worksheets.Item(i);
                        if (ws) {
                            const usedRange = ws.UsedRange;
                            allWorksheets.push({
                                name: ws.Name,
                                usedRange: usedRange ? {
                                    rows: usedRange.Rows.Count,
                                    columns: usedRange.Columns.Count
                                } : { rows: 0, columns: 0 }
                            });
                        }
                    } catch (error) {
                        console.warn(`处理工作表 ${i} 失败:`, error);
                    }
                }
            }
            
            // 获取当前工作表详细信息
            const currentWorksheet = {
                name: activeSheet ? activeSheet.Name : '未知工作表',
                usedRange: null
            };
            if (activeSheet && activeSheet.UsedRange) {
                const usedRange = activeSheet.UsedRange;
                currentWorksheet.usedRange = {
                    rows: usedRange.Rows.Count,
                    columns: usedRange.Columns.Count
                };
            }
            
            // 获取当前工作表的所有列标题
            let columnHeaders = [];
            if (activeSheet) {
                columnHeaders = this.extractHeaders(activeSheet);
            }
            
            return {
                currentWorkbook,
                currentWorksheet,
                allWorksheets,
                columnHeaders
            };
            
        } catch (error) {
            console.warn('获取当前工作簿信息失败:', error);
            return {
                currentWorkbook: { name: '获取失败' },
                currentWorksheet: { name: '获取失败', usedRange: { rows: 0, columns: 0 } },
                allWorksheets: [],
                columnHeaders: []
            };
        }
    }

    /**
     * 构建智能分析描述
     * 根据当前单元格信息、工作表信息等生成分析需求
     */
    buildIntelligentAnalysisDescription() {
        console.log('🤖 [buildIntelligentAnalysisDescription] 开始生成智能分析描述...');
        
        let description = '';
        
        // 优先检查用户输入框中的描述（可能仍有内容）
        const requirementInput = document.getElementById('formulaDescription');
        const userInputDescription = requirementInput ? requirementInput.value.trim() : '';
        
        if (userInputDescription && userInputDescription !== '') {
            // 如果用户输入框中有内容，优先使用用户描述
            description = userInputDescription;
            console.log('📝 [buildIntelligentAnalysisDescription] 使用用户输入框中的描述:', description);
            return description;
        }
        
        // 如果输入框为空，检查this.formulaDescription（状态管理中的值）
        if (this.formulaDescription && this.formulaDescription.trim() !== '') {
            description = this.formulaDescription.trim();
            console.log('💾 [buildIntelligentAnalysisDescription] 使用this.formulaDescription中的描述:', description);
            return description;
        }
        
        console.log('🤖 [buildIntelligentAnalysisDescription] 用户无描述，构建智能分析描述');
        
        // 智能分析需求 - 基于当前单元格和工作表信息
        const currentCell = this.getCurrentCellInfo();
        if (currentCell && currentCell.workSheetName) {
            description += `当前工作表：${currentCell.workSheetName}，单元格：${currentCell.cellAddress}，列名：${currentCell.columnName}`;
            console.log('📍 [buildIntelligentAnalysisDescription] 当前单元格信息:', currentCell);
        } else {
            description += `单元格分析需求`;
            console.log('⚠️ [buildIntelligentAnalysisDescription] 无法获取当前单元格信息');
        }
        
        // 获取工作表列标题信息
        const workbookInfo = this.getCurrentWorkbookInfo();
        if (workbookInfo && workbookInfo.columnHeaders && workbookInfo.columnHeaders.length > 0) {
            description += `\n\n当前工作表列标题（${workbookInfo.columnHeaders.length}个）：`;
            description += `\n${workbookInfo.columnHeaders.join(', ')}`;
            console.log('📊 [buildIntelligentAnalysisDescription] 当前工作表列标题:', workbookInfo.columnHeaders);
        }
        
        // 获取选中工作簿的详细信息
        const selectedWorkbooks = this.buildSelectedWorkbooksInfo();
        if (selectedWorkbooks && selectedWorkbooks.length > 0) {
            description += `\n\n可用的数据源（共${selectedWorkbooks.length}个工作簿）：`;
            
            selectedWorkbooks.forEach(workbook => {
                description += `\n- 工作簿：${workbook.workBookName}`;
                
                if (workbook.worksheets && workbook.worksheets.length > 0) {
                    workbook.worksheets.forEach(worksheet => {
                        const headersCount = worksheet.columnHeaders ? Object.keys(worksheet.columnHeaders).length : 0;
                        description += `\n  - 工作表：${worksheet.workSheetName} (${headersCount}个列标题)`;
                        
                        // 显示有列标题的工作表的详细信息
                        if (worksheet.columnHeaders && Object.keys(worksheet.columnHeaders).length > 0) {
                            const headerList = Object.values(worksheet.columnHeaders).join(', ');
                            description += `\n    列标题：${headerList}`;
                        }
                    });
                }
            });
            
            console.log('📁 [buildIntelligentAnalysisDescription] 选中工作簿信息:', selectedWorkbooks);
        }
        
        // 获取引用类型信息
        const referenceType = this.referenceType || 'current';
        description += `\n\n参考类型：${referenceType}`;
        
        // 根据当前单元格列名和表头信息生成分析需求
        if (currentCell && currentCell.columnName) {
            description += `\n\n🎯 核心任务 - 针对"${currentCell.columnName}"列的Excel公式建议：`;            
            description += `\n\n当前情况：`;
            description += `\n- 选中的单元格：${currentCell.cellAddress}`;
            description += `\n- 所在列名：${currentCell.columnName}`;
            description += `\n- 当前值：待计算`;
            description += `\n- 业务含义：${currentCell.columnName}应该是通过其他库存数据计算得出`;
            
            description += `\n\n具体需求：`;
            description += `\n1. 📊 提供计算"${currentCell.columnName}"的Excel公式`;
            description += `\n2. 🔄 公式应该基于同行的其他库存数据（如中央库库位库存、配套库存等）`;
            description += `\n3. 🎯 优先使用SUM、SUMIF、VLOOKUP等库存计算常用函数`;
            description += `\n4. 🛡️ 考虑数据验证和错误处理`;
            description += `\n5. 📈 确保公式适合向下或向右填充`;
            
            description += `\n\n请重点分析如何基于当前行的"中央库库位库存"、"配套库存"等数据来计算"${currentCell.columnName}"的值。`;
        }
        
        description += `\n\n🔍 分析提示：`;
        description += `\n- 这是一个库存管理系统的数据表`;
        description += `\n- "${currentCell?.columnName || '目标列'}"可能需要计算得出`;
        description += `\n- 建议的公式类型：汇总公式、查找公式、条件计算公式`;
        description += `\n- 输出格式：公式内容 + 详细解释 + 使用示例`;
        
        const finalDescription = description.trim();
        console.log('✅ [buildIntelligentAnalysisDescription] 智能描述生成完成:', finalDescription);
        
        return finalDescription;
    }
    
    /**
     * 获取当前单元格信息
     */
    getCurrentCellInfo() {
        try {
            if (window.Application && window.Application.ActiveSheet && window.Application.Selection) {
                const activeCell = window.Application.Selection;
                const activeWorkbook = window.Application.ActiveWorkbook;
                const activeSheet = window.Application.ActiveSheet;
                
                if (activeCell) {
                    // 获取单元格地址和列标题
                    const cellAddress = this.getCellAddress(activeCell.Row || 1, activeCell.Column || 1);
                    const columnName = this.getColumnHeaderFromCell(activeCell, activeSheet);
                    
                    return {
                        cellAddress: cellAddress,
                        columnName: columnName,
                        workSheetName: activeSheet ? activeSheet.Name : '未知工作表',
                        workBookName: activeWorkbook ? activeWorkbook.Name : '未知工作簿',
                        workBookPath: activeWorkbook ? activeWorkbook.Path : ''
                    };
                }
            }
        } catch (error) {
            console.warn('⚠️ [getCurrentCellInfo] 获取当前单元格信息失败:', error);
        }
        
        return null;
    }
    
    collectHeaders() {
        const headers = [];
        
        this.selectedWorksheets.forEach(workbook => {
            workbook.worksheets.forEach(worksheet => {
                try {
                    const ws = window.Application.Workbooks(workbook.name).Worksheets(worksheet.name);
                    if (ws) {
                        const headerRow = this.extractHeaders(ws);
                        headers.push({
                            workbook: workbook.name,
                            worksheet: worksheet.name,
                            headers: headerRow
                        });
                    }
                } catch (error) {
                    console.warn(`获取工作表 ${workbook.name} - ${worksheet.name} 的表头失败:`, error);
                }
            });
        });
        
        return headers;
    }
    
    extractHeaders(worksheet) {
        try {
            // 优先从第一行获取表头
            const firstRow = worksheet.Rows.Item(1);
            if (firstRow) {
                const headers = [];
                let emptyColumnCount = 0;
                const maxEmptyColumnsThreshold = 5; // 连续5列为空则停止
                
                console.log(`🔍 [extractHeaders] 开始提取表头...`);
                
                // 首先尝试获取实际使用范围来限制列数
                let maxColumnsToCheck = 20; // 默认限制为20列
                try {
                    const usedRange = worksheet.UsedRange;
                    if (usedRange && usedRange.Columns.Count) {
                        // 获取实际使用的列数，并加上一些缓冲
                        maxColumnsToCheck = Math.min(usedRange.Columns.Count + 5, 50);
                        console.log(`📊 [extractHeaders] 检测到实际使用范围：${usedRange.Columns.Count}列，限制检查范围为${maxColumnsToCheck}列`);
                    }
                } catch (rangeError) {
                    console.warn(`⚠️ [extractHeaders] 获取使用范围失败，使用默认限制:`, rangeError);
                }
                
                // 获取实际有数据的列标题，最多检查maxColumnsToCheck列
                for (let col = 1; col <= maxColumnsToCheck; col++) {
                    try {
                        const columnLetter = this.getColumnLetter(col);
                        const cellAddress = `${columnLetter}1`;
                        
                        // 使用WPS规范的Range方式获取单元格值
                        const cell = worksheet.Range(cellAddress);
                        
                        let cellValue = '';
                        
                        if (cell) {
                            try {
                                // 优先使用Value2（原始值），如果为空则尝试Text（显示文本）
                                let rawValue = cell.Value2;
                                
                                if (rawValue === null || rawValue === undefined) {
                                    console.log(`🔍 [extractHeaders] ${cellAddress} Value2为空，尝试使用Text`);
                                    rawValue = cell.Text;
                                }
                                
                                // 数据类型处理
                                if (rawValue === null || rawValue === undefined) {
                                    cellValue = '';
                                } else if (typeof rawValue === 'string') {
                                    cellValue = rawValue.trim();
                                } else if (typeof rawValue === 'number') {
                                    cellValue = String(rawValue);
                                } else if (rawValue instanceof Date) {
                                    cellValue = rawValue.toLocaleDateString();
                                } else if (typeof rawValue === 'boolean') {
                                    cellValue = rawValue ? 'TRUE' : 'FALSE';
                                } else if (typeof rawValue === 'object' && rawValue !== null) {
                                    // 如果是对象，尝试获取Text属性
                                    if (rawValue.Text && typeof rawValue.Text === 'string') {
                                        cellValue = rawValue.Text.trim();
                                    } else {
                                        // 最后尝试转为字符串
                                        const strValue = String(rawValue);
                                        if (!strValue.includes('function') && !strValue.includes('[native code]')) {
                                            cellValue = strValue.trim();
                                        }
                                    }
                                } else {
                                    // 其他类型转换为字符串
                                    cellValue = String(rawValue).trim();
                                }
                                
                            } catch (valueError) {
                                console.warn(`⚠️ [extractHeaders] 处理${cellAddress}单元格值时出错:`, valueError);
                                cellValue = '';
                            }
                        }
                        
                        // 检查是否为有效数据
                        const isValidData = cellValue && 
                                          cellValue !== '' && 
                                          !cellValue.includes('function') && 
                                          !cellValue.includes('[native code]') &&
                                          !cellValue.startsWith('列');
                        
                        if (isValidData) {
                            headers.push(cellValue);
                            emptyColumnCount = 0; // 重置空列计数器
                            console.log(`📝 [extractHeaders] 列${col}: "${cellValue}"`);
                        } else {
                            // 检查是否是空列（使用默认列名的列）
                            if (!cellValue || cellValue.startsWith('列')) {
                                emptyColumnCount++;
                                // 如果连续5列都是空的，停止获取更多列
                                if (emptyColumnCount >= maxEmptyColumnsThreshold) {
                                    console.log(`🛑 [extractHeaders] 连续${emptyColumnCount}列为空，停止获取更多列`);
                                    break;
                                }
                            }
                        }
                        
                    } catch (error) {
                        console.warn(`⚠️ [extractHeaders] 获取第${col}列表头失败:`, error);
                        emptyColumnCount++;
                        // 错误列也计入空列计数
                        if (emptyColumnCount >= maxEmptyColumnsThreshold) {
                            break;
                        }
                    }
                }
                
                console.log(`✅ [extractHeaders] 表头提取完成，获得${headers.length}个有效列标题`);
                
                // 如果没有获取到有效表头，尝试从数据范围查找
                if (headers.length === 0) {
                    console.log('🔍 [extractHeaders] 第一行无有效表头，尝试从数据范围查找');
                    return this.extractHeadersFromDataRange(worksheet);
                }
                
                return headers;
            }
        } catch (error) {
            console.error('❌ [extractHeaders] 提取表头失败:', error);
        }
        
        return [];
    }
    
    /**
     * 从数据范围中提取表头（当第一行为空时使用）
     */
    extractHeadersFromDataRange(worksheet) {
        try {
            const usedRange = worksheet.UsedRange;
            if (usedRange && usedRange.Rows.Count > 1) {
                console.log(`🔍 [extractHeadersFromDataRange] 开始从数据范围提取表头...`);
                
                // 获取实际使用的列数，并限制检查范围
                let maxColumnsToCheck = Math.min(usedRange.Columns.Count, 50);
                console.log(`📊 [extractHeadersFromDataRange] 检测到使用范围：${usedRange.Columns.Count}列`);
                
                // 尝试第一行到第五行，查找第一个非空行作为表头
                const maxRowToCheck = Math.min(5, usedRange.Rows.Count);
                
                for (let row = 1; row <= maxRowToCheck; row++) {
                    const headerRow = usedRange.Rows.Item(row);
                    let hasData = false;
                    const rowHeaders = [];
                    let emptyColumnCount = 0;
                    const maxEmptyColumnsThreshold = 5; // 连续5列为空则停止
                    
                    console.log(`🔍 [extractHeadersFromDataRange] 检查第${row}行...`);
                    
                    for (let col = 1; col <= maxColumnsToCheck; col++) {
                        try {
                            // 获取列字母并构建单元格地址
                            const columnLetter = this.getColumnLetter(col);
                            const cellAddress = columnLetter + row;
                            const cell = headerRow.Range(cellAddress);
                            let cellValue = '';
                            
                            // 优先使用 Value2 属性获取值
                            if (cell && cell.Value2 !== null && cell.Value2 !== undefined) {
                                const value = cell.Value2;
                                
                                // 类型安全处理
                                if (typeof value === 'string') {
                                    const trimmed = value.trim();
                                    if (!trimmed.includes('function') && !trimmed.includes('[native code]')) {
                                        cellValue = trimmed;
                                    }
                                } else if (typeof value === 'number' && !isNaN(value)) {
                                    cellValue = value.toString();
                                } else if (value instanceof Date) {
                                    cellValue = value.toLocaleDateString();
                                } else if (typeof value === 'boolean') {
                                    cellValue = value ? 'TRUE' : 'FALSE';
                                } else if (typeof value === 'object' && value !== null) {
                                    if (value.Text && typeof value.Text === 'string' && value.Text.trim()) {
                                        cellValue = value.Text.trim();
                                    } else if (value.hasOwnProperty('Text')) {
                                        const textValue = value.Text;
                                        if (typeof textValue === 'string' && textValue.trim()) {
                                            cellValue = textValue.trim();
                                        }
                                    }
                                }
                            }
                            // 如果 Value2 为空，尝试使用 Text 属性
                            else if (cell && cell.Text && typeof cell.Text === 'string' && cell.Text.trim()) {
                                cellValue = cell.Text.trim();
                            }
                            
                            // 检查是否为有效数据
                            const isValidData = cellValue && 
                                              cellValue !== '' && 
                                              !cellValue.includes('function') && 
                                              !cellValue.includes('[native code]') &&
                                              !cellValue.startsWith('列');
                            
                            if (isValidData) {
                                rowHeaders.push(cellValue);
                                emptyColumnCount = 0; // 重置空列计数器
                                hasData = true;
                                console.log(`📝 [extractHeadersFromDataRange] 行${row}列${col}: "${cellValue}"`);
                            } else {
                                // 检查是否是空列（使用默认列名的列）
                                if (!cellValue || cellValue.startsWith('列')) {
                                    emptyColumnCount++;
                                    // 如果连续5列都是空的，停止获取更多列
                                    if (emptyColumnCount >= maxEmptyColumnsThreshold) {
                                        console.log(`🛑 [extractHeadersFromDataRange] 连续${emptyColumnCount}列为空，停止获取更多列`);
                                        break;
                                    }
                                }
                            }
                        } catch (error) {
                            console.warn(`⚠️ [extractHeadersFromDataRange] 获取第${row}行第${col}列失败:`, error);
                            emptyColumnCount++;
                            // 错误列也计入空列计数
                            if (emptyColumnCount >= maxEmptyColumnsThreshold) {
                                break;
                            }
                        }
                    }
                    
                    if (hasData && rowHeaders.length > 0) {
                        console.log(`✅ [extractHeadersFromDataRange] 在第${row}行找到有效表头: ${rowHeaders.length}个列标题`);
                        return rowHeaders;
                    } else if (rowHeaders.length > 0) {
                        console.log(`⚠️ [extractHeadersFromDataRange] 第${row}行无有效数据，跳过`);
                    }
                }
                
                console.log('⚠️ [extractHeadersFromDataRange] 在前5行中未找到有效表头');
            }
        } catch (error) {
            console.warn('❌ [extractHeadersFromDataRange] 从数据范围提取表头失败:', error);
        }
        
        return [];
    }

    /**
     * 收集请求数据（按新格式要求修改）- 增强错误处理和堆栈跟踪
     */
    collectRequestData() {
        console.log('🔍 [collectRequestData] 开始收集请求数据...');
        console.trace('🔍 [collectRequestData] 调用堆栈:');
        
        const requirementInput = document.getElementById('formulaDescription');
        const referenceType = document.querySelector('input[name="referenceType"]:checked');
        
        // 获取当前单元格信息
        let currentCell = {
            cellAddress: '',
            columnName: '',
            workSheetName: '',
            workBookName: '',
            workBookPath: ''
        };
        
        try {
            // 检查Excel对象状态
            console.log('🔍 [collectRequestData] Excel对象状态检查...');
            console.log('  - window.Application存在:', !!window.Application);
            console.log('  - window.Application.ActiveSheet存在:', !!(window.Application && window.Application.ActiveSheet));
            console.log('  - window.Application.Selection存在:', !!(window.Application && window.Application.Selection));
            console.log('  - window.Application.ActiveWorkbook存在:', !!(window.Application && window.Application.ActiveWorkbook));
            
            if (window.Application && window.Application.ActiveSheet) {
                const activeCell = window.Application.Selection;
                const activeWorkbook = window.Application.ActiveWorkbook;
                const activeSheet = window.Application.ActiveSheet;
                
                if (activeCell) {
                    // 获取单元格地址和列标题
                    console.log('📍 [collectRequestData] 原始单元格信息:', {
                        row: activeCell.Row,
                        column: activeCell.Column,
                        value: activeCell.Value
                    });
                    
                    const cellAddress = this.getCellAddress(activeCell.Row || 1, activeCell.Column || 1);
                    const columnName = this.getColumnHeaderFromCell(activeCell, activeSheet); // 获取该列第一行的标题
                    
                    currentCell = {
                        cellAddress: cellAddress,
                        columnName: columnName || '未知列',
                        workSheetName: activeSheet ? activeSheet.Name : '未知工作表',
                        workBookName: activeWorkbook ? activeWorkbook.Name : '未知工作簿',
                        workBookPath: activeWorkbook ? activeWorkbook.Path : ''
                    };
                    
                    // 验证单元格信息完整性
                    console.log('✅ [collectRequestData] 单元格信息完整性检查:');
                    console.log('  - cellAddress:', currentCell.cellAddress);
                    console.log('  - columnName:', currentCell.columnName);
                    console.log('  - workSheetName:', currentCell.workSheetName);
                    console.log('  - workBookName:', currentCell.workBookName);
                    
                } else {
                    console.warn('⚠️ [collectRequestData] 未选择任何单元格');
                }
                
                console.log('📍 [collectRequestData] 当前单元格:', currentCell);
            } else {
                console.warn('⚠️ [collectRequestData] Excel应用程序或活动工作表不可用');
            }
        } catch (error) {
            console.error('❌ [collectRequestData] 获取当前单元格信息时发生错误:', error);
            console.error('  错误类型:', error.constructor.name);
            console.error('  错误消息:', error.message);
            console.error('  错误堆栈:', error.stack);
        }
        
        // 构建设选工作簿信息（按新格式）- 增强错误处理
        console.log('📚 [collectRequestData] 开始构建工作簿信息...');
        let selectedWorkbooks = [];
        
        try {
            console.log('🔄 [collectRequestData] 调用buildSelectedWorkbooksInfo...');
            selectedWorkbooks = this.buildSelectedWorkbooksInfo();
            console.log('✅ [collectRequestData] 工作簿信息构建完成，工作簿数量:', selectedWorkbooks.length);
        } catch (error) {
            console.error('❌ [collectRequestData] 构建工作簿信息时发生错误:', error);
            console.error('  错误类型:', error.constructor.name);
            console.error('  错误消息:', error.message);
            console.error('  错误堆栈:', error.stack);
            selectedWorkbooks = []; // 设置为空数组以避免后续错误
        }
        
        // 获取当前工作簿的完整信息，包括列标题和所有工作表
        console.log('📊 [collectRequestData] 开始获取当前工作簿信息...');
        let workbookInfo = {};
        
        try {
            console.log('🔄 [collectRequestData] 调用getCurrentWorkbookInfo...');
            workbookInfo = this.getCurrentWorkbookInfo();
            console.log('✅ [collectRequestData] 当前工作簿信息获取完成，列标题数量:', workbookInfo.columnHeaders ? workbookInfo.columnHeaders.length : 0);
        } catch (error) {
            console.error('❌ [collectRequestData] 获取当前工作簿信息时发生错误:', error);
            console.error('  错误类型:', error.constructor.name);
            console.error('  错误消息:', error.message);
            console.error('  错误堆栈:', error.stack);
            workbookInfo = { columnHeaders: [], allWorksheets: [] }; // 设置默认值
        }
        
        // 如果没有用户描述，使用智能分析描述 - 增强错误处理
        console.log('📝 [collectRequestData] 开始处理描述字段...');
        let description = this.formulaDescription || (requirementInput ? requirementInput.value : '');
        
        if (!description || description.trim() === '') {
            console.log('⚠️ [collectRequestData] 描述为空，开始生成智能分析描述...');
            try {
                console.log('🔄 [collectRequestData] 调用buildIntelligentAnalysisDescription...');
                description = this.buildIntelligentAnalysisDescription();
                console.log('✅ [collectRequestData] 智能分析描述生成成功，描述长度:', description.length);
                console.log('🤖 [collectRequestData] 智能分析描述内容:', description);
            } catch (error) {
                console.error('❌ [collectRequestData] 生成智能分析描述时发生错误:', error);
                console.error('  错误类型:', error.constructor.name);
                console.error('  错误消息:', error.message);
                console.error('  错误堆栈:', error.stack);
                description = '请分析当前工作表的数据并生成相应的公式建议。'; // 设置默认描述
                console.log('🔄 [collectRequestData] 使用默认描述:', description);
            }
        } else {
            console.log('✅ [collectRequestData] 使用用户输入描述，描述长度:', description.length);
        }
        
        const requestData = {
            description: description,
            referenceType: this.referenceType || (referenceType ? referenceType.value : 'current'),
            currentCell: currentCell,
            selectedWorkbooks: selectedWorkbooks,
            // 添加AI接口期望的字段
            columnHeaders: workbookInfo.columnHeaders || [],
            allWorksheets: workbookInfo.allWorksheets || []
        };
        
        console.log('✅ [collectRequestData] 数据收集完成');
        console.log('📊 [collectRequestData] 收集到的数据:', {
            description: requestData.description,
            referenceType: requestData.referenceType,
            currentCell: requestData.currentCell,
            selectedWorkbooksCount: requestData.selectedWorkbooks.length,
            columnHeadersCount: requestData.columnHeaders.length,
            allWorksheetsCount: requestData.allWorksheets.length
        });
        
        // 最终验证和完整性检查
        console.log('🔍 [collectRequestData] 进行最终数据验证...');
        
        try {
            // 验证关键字段
            console.log('🔍 [collectRequestData] 验证description字段...');
            if (!requestData.description || typeof requestData.description !== 'string') {
                throw new Error('description字段无效或缺失');
            }
            
            console.log('🔍 [collectRequestData] 验证referenceType字段...');
            if (!requestData.referenceType || typeof requestData.referenceType !== 'string') {
                console.warn('⚠️ [collectRequestData] referenceType字段无效，使用默认值 "current"');
                requestData.referenceType = 'current';
            }
            
            console.log('🔍 [collectRequestData] 验证currentCell字段...');
            if (!requestData.currentCell || typeof requestData.currentCell !== 'object') {
                console.warn('⚠️ [collectRequestData] currentCell字段无效，使用默认值');
                requestData.currentCell = {
                    cellAddress: '',
                    columnName: '',
                    workSheetName: '',
                    workBookName: '',
                    workBookPath: ''
                };
            }
            
            console.log('🔍 [collectRequestData] 验证selectedWorkbooks字段...');
            if (!Array.isArray(requestData.selectedWorkbooks)) {
                console.warn('⚠️ [collectRequestData] selectedWorkbooks字段无效，使用空数组');
                requestData.selectedWorkbooks = [];
            }
            
            console.log('🔍 [collectRequestData] 验证columnHeaders字段...');
            if (!Array.isArray(requestData.columnHeaders)) {
                console.warn('⚠️ [collectRequestData] columnHeaders字段无效，使用空数组');
                requestData.columnHeaders = [];
            }
            
            console.log('🔍 [collectRequestData] 验证allWorksheets字段...');
            if (!Array.isArray(requestData.allWorksheets)) {
                console.warn('⚠️ [collectRequestData] allWorksheets字段无效，使用空数组');
                requestData.allWorksheets = [];
            }
            
            // 输出最终验证结果
            console.log('✅ [collectRequestData] 数据验证完成，所有字段状态良好');
            console.log('📋 [collectRequestData] 最终数据结构:', {
                descriptionLength: requestData.description.length,
                referenceType: requestData.referenceType,
                hasCurrentCell: !!(requestData.currentCell && requestData.currentCell.cellAddress),
                selectedWorkbooksCount: requestData.selectedWorkbooks.length,
                columnHeadersCount: requestData.columnHeaders.length,
                allWorksheetsCount: requestData.allWorksheets.length
            });
            
        } catch (validationError) {
            console.error('❌ [collectRequestData] 数据验证失败:', validationError);
            console.error('  错误类型:', validationError.constructor.name);
            console.error('  错误消息:', validationError.message);
            console.error('  错误堆栈:', validationError.stack);
            
            // 如果验证失败，抛出错误以通知调用者
            throw new Error(`数据收集验证失败: ${validationError.message}`);
        }
        
        console.log('✅ [collectRequestData] 数据收集和验证完全成功');
        console.log('🔄 [collectRequestData] 返回最终请求数据');
        
        return requestData;
    }
    
    /**
     * 构建选中工作簿信息（按新格式要求）
     */
    buildSelectedWorkbooksInfo() {
        const selectedWorkbooks = [];
        
        try {
            // 如果有选中的工作簿信息，检查是否有有效数据
            if (this.selectedWorkbooks && this.selectedWorkbooks.length > 0) {
                console.log(`📊 [buildSelectedWorkbooksInfo] 使用已选择的工作簿信息，共 ${this.selectedWorkbooks.length} 个工作簿`);
                
                let hasValidData = false;
                const processedWorkbooks = [];
                
                this.selectedWorkbooks.forEach((workbook, wbIndex) => {
                    console.log(`📁 [buildSelectedWorkbooksInfo] 处理已选择的工作簿 ${wbIndex + 1}: ${workbook.name || workbook.workBookName || '未知'}`);
                    const workbookInfo = {
                        workBookName: workbook.name || workbook.workBookName || '',
                        workBookPath: workbook.path || workbook.workBookPath || '',
                        worksheets: []
                    };
                    
                    if (workbook.worksheets && workbook.worksheets.length > 0) {
                        console.log(`📋 [buildSelectedWorkbooksInfo] 工作簿 ${workbook.name} 有 ${workbook.worksheets.length} 个工作表`);
                        workbook.worksheets.forEach((worksheet, wsIndex) => {
                            console.log(`🔍 [buildSelectedWorkbooksInfo] 处理工作表 ${wsIndex + 1}: ${worksheet.name || worksheet.workSheetName || '未知'}`);
                            
                            // 尝试从已有数据获取列标题，如果没有则重新提取
                            let columnHeaders = worksheet.headers || worksheet.columnHeaders || [];
                            console.log(`📊 [buildSelectedWorkbooksInfo] 工作表 ${worksheet.name} 原始列标题:`, columnHeaders);
                            
                            // 如果没有现有列标题数据，尝试重新提取
                            if (!columnHeaders || columnHeaders.length === 0) {
                                console.log(`🔄 [buildSelectedWorkbooksInfo] 工作表 ${worksheet.name} 没有现有列标题，尝试重新提取...`);
                                try {
                                    // 尝试切换到该工作表进行提取
                                    const activeWorkbook = window.Application.ActiveWorkbook;
                                    if (activeWorkbook) {
                                        const targetWorksheet = activeWorkbook.Worksheets.Item(worksheet.name || worksheet.workSheetName);
                                        if (targetWorksheet) {
                                            columnHeaders = this.extractHeaders(targetWorksheet);
                                            console.log(`📋 [buildSelectedWorkbooksInfo] 工作表 ${worksheet.name} 重新提取结果:`, columnHeaders.length, '个列标题');
                                        }
                                    }
                                } catch (error) {
                                    console.warn(`⚠️ [buildSelectedWorkbooksInfo] 工作表 ${worksheet.name} 重新提取失败:`, error);
                                }
                            }
                            
                            // 检查是否有有效数据
                            if (columnHeaders && columnHeaders.length > 0) {
                                hasValidData = true;
                            }
                            
                            const worksheetInfo = {
                                workSheetName: worksheet.name || worksheet.workSheetName || '',
                                columnHeaders: this.convertHeadersToNewFormat(columnHeaders)
                            };
                            
                            console.log(`🔄 [buildSelectedWorkbooksInfo] 工作表 ${worksheet.name} 转换后的列标题:`, worksheetInfo.columnHeaders);
                            workbookInfo.worksheets.push(worksheetInfo);
                        });
                    }
                    
                    processedWorkbooks.push(workbookInfo);
                });
                
                // 如果已选择的工作簿有有效数据，使用它们；否则使用当前工作簿
                 if (hasValidData) {
                     console.log('✅ [buildSelectedWorkbooksInfo] 已选择的工作簿有有效数据，使用它们');
                     return processedWorkbooks;
                 } else {
                     console.log('⚠️ [buildSelectedWorkbooksInfo] 已选择的工作簿没有有效数据，使用当前活动工作簿');
                 }
            }
            
            // 如果没有选中的工作簿或已选择的工作簿没有有效数据，使用当前工作簿
            const activeWorkbook = window.Application.ActiveWorkbook;
            if (activeWorkbook) {
                const workbookInfo = {
                    workBookName: activeWorkbook.Name,
                    workBookPath: activeWorkbook.Path || '',
                    worksheets: []
                };
                
                // 获取所有工作表信息
                if (activeWorkbook.Worksheets) {
                    console.log(`📊 [buildSelectedWorkbooksInfo] 找到 ${activeWorkbook.Worksheets.Count} 个工作表`);
                    for (let i = 1; i <= activeWorkbook.Worksheets.Count; i++) {
                        try {
                            const ws = activeWorkbook.Worksheets.Item(i);
                            if (ws) {
                                console.log(`🔍 [buildSelectedWorkbooksInfo] 处理工作表 ${i}: ${ws.Name}`);
                                
                                // 获取列标题
                                const columnHeaders = this.extractHeaders(ws);
                                console.log(`📋 [buildSelectedWorkbooksInfo] 工作表 ${ws.Name} 列标题提取结果:`, columnHeaders.length, '个');
                                
                                const convertedHeaders = this.convertHeadersToNewFormat(columnHeaders);
                                console.log(`🔄 [buildSelectedWorkbooksInfo] 工作表 ${ws.Name} 转换后的列标题:`, convertedHeaders);
                                
                                const worksheetInfo = {
                                    workSheetName: ws.Name,
                                    columnHeaders: convertedHeaders
                                };
                                
                                workbookInfo.worksheets.push(worksheetInfo);
                            }
                        } catch (error) {
                            console.warn(`⚠️ [buildSelectedWorkbooksInfo] 处理工作表 ${i} 失败:`, error);
                        }
                    }
                }
                
                selectedWorkbooks.push(workbookInfo);
            }
        } catch (error) {
            console.warn('⚠️ [buildSelectedWorkbooksInfo] 构建工作簿信息失败:', error);
        }
        
        return selectedWorkbooks;
    }
    
    /**
     * 将列标题数组转换为新格式（键值对形式，键为列字母）
     */
    convertHeadersToNewFormat(headersArray) {
        const columnHeaders = {};
        
        if (Array.isArray(headersArray)) {
            headersArray.forEach((header, index) => {
                const columnLetter = this.getColumnLetter(index + 1); // 列索引从1开始
                
                if (header && typeof header === 'string' && header.trim()) {
                    columnHeaders[columnLetter] = header.trim();
                } else if (header && typeof header === 'object' && header.value) {
                    // 如果是对象格式，提取value字段
                    columnHeaders[columnLetter] = header.value.trim();
                }
            });
        } else if (typeof headersArray === 'object' && headersArray !== null) {
            // 如果已经是对象格式，直接返回
            return headersArray;
        }
        
        return columnHeaders;
    }
    
    /**
     * 获取列字母（如 A, B, C...）
     */
    getColumnLetter(columnNumber) {
        let result = '';
        let n = columnNumber;
        while (n > 0) {
            n--;
            result = String.fromCharCode(65 + (n % 26)) + result;
            n = Math.floor(n / 26);
        }
        return result;
    }
    
    /**
     * 获取单元格所在列的第一行标题（表头）
     * @param {Object} activeCell - Excel单元格对象
     * @param {Object} activeSheet - Excel工作表对象
     * @returns {string} 列标题
     */
    getColumnHeaderFromCell(activeCell, activeSheet) {
        try {
            if (!activeCell || !activeSheet) {
                console.warn('⚠️ [getColumnHeaderFromCell] 单元格或工作表对象为空');
                return '';
            }
            
            const columnIndex = activeCell.Column || 1;
            const columnLetter = this.getColumnLetter(columnIndex);
            const cellAddress = `${columnLetter}1`;
            
            console.log(`🔍 [getColumnHeaderFromCell] 尝试获取${cellAddress}单元格值...`);
            
            // 使用WPS规范的Range方式获取单元格值
            const headerRange = activeSheet.Range(cellAddress);
            
            if (!headerRange) {
                console.warn(`⚠️ [getColumnHeaderFromCell] 无法获取${cellAddress}单元格`);
                return '';
            }
            
            let headerValue = '';
            
            try {
                // 优先使用Value2（原始值），如果为空则尝试Text（显示文本）
                let rawValue = headerRange.Value2;
                
                if (rawValue === null || rawValue === undefined) {
                    console.log(`🔍 [getColumnHeaderFromCell] Value2为空，尝试使用Text`);
                    rawValue = headerRange.Text;
                }
                console.log('🔍 [getColumnHeaderFromCell] 原始值类型:', typeof rawValue, '值:', rawValue);
                
                if (rawValue === null || rawValue === undefined) {
                    console.warn(`⚠️ [getColumnHeaderFromCell] ${cellAddress}单元格值为空`);
                    headerValue = '';
                } else if (typeof rawValue === 'string') {
                    headerValue = rawValue.trim();
                } else if (typeof rawValue === 'number') {
                    headerValue = String(rawValue);
                } else if (rawValue instanceof Date) {
                    headerValue = rawValue.toLocaleDateString();
                } else if (typeof rawValue === 'boolean') {
                    headerValue = rawValue ? 'TRUE' : 'FALSE';
                } else if (typeof rawValue === 'object' && rawValue !== null) {
                    // 如果是对象，尝试获取Text属性
                    if (rawValue.Text && typeof rawValue.Text === 'string') {
                        headerValue = rawValue.Text.trim();
                    } else {
                        // 最后尝试转为字符串
                        const strValue = String(rawValue);
                        if (!strValue.includes('function') && !strValue.includes('[native code]')) {
                            headerValue = strValue.trim();
                        }
                    }
                } else {
                    // 其他类型转换为字符串
                    headerValue = String(rawValue).trim();
                }
                
            } catch (valueError) {
                console.warn('⚠️ [getColumnHeaderFromCell] 处理单元格值时出错:', valueError);
                headerValue = '';
            }
            
            // 过滤函数字符串和空值
            if (!headerValue || (typeof headerValue === 'string' && (headerValue.startsWith('=') || headerValue.includes('function') || headerValue.includes('[native code]')))) {
                console.log(`⚠️ [getColumnHeaderFromCell] 跳过无效值: "${headerValue}"`);
                return '';
            }
            
            console.log(`📋 [getColumnHeaderFromCell] 第${columnIndex}列标题: "${headerValue}"`);
            return headerValue;
            
        } catch (error) {
            console.warn(`⚠️ [getColumnHeaderFromCell] 获取列标题失败:`, error);
            return '';
        }
    }
    
    showFormulaSuggestions(formulas) {
        const container = document.getElementById('formulaSuggestions');
        const resultsArea = document.getElementById('aiResults');
        
        container.innerHTML = '';
        
        formulas.forEach((formula, index) => {
            const card = document.createElement('div');
            card.className = 'formula-card';
            card.innerHTML = `
                <div class="formula-header">
                    <div class="formula-title">${formula.title}</div>
                    <div class="formula-confidence">${formula.confidence}%</div>
                </div>
                <div class="formula-content">${formula.formula}</div>
                <div class="formula-explanation">${formula.explanation}</div>
            `;
            
            card.addEventListener('click', () => {
                this.selectFormula(card, formula);
            });
            
            container.appendChild(card);
        });
        
        resultsArea.style.display = 'block';
    }
    
    selectFormula(card, formula) {
        // 移除其他选中状态
        document.querySelectorAll('.formula-card').forEach(item => {
            item.classList.remove('selected');
        });
        
        // 添加选中状态
        card.classList.add('selected');
        this.selectedFormula = formula;
        
        // 启用应用按钮
        document.getElementById('applyFormula').disabled = false;
    }
    
    applySelectedFormula() {
        try {
            if (!this.selectedFormula) {
                this.showNotification('请先选择一个公式建议', 'error');
                return;
            }
            
            // 获取当前活动单元格
            const activeSheet = window.Application.ActiveSheet;
            const selection = window.Application.Selection;
            
            if (!selection || !activeSheet) {
                this.showNotification('无法获取当前活动单元格', 'error');
                return;
            }
            
            // 应用公式
            const formula = this.adjustFormulaForFill(this.selectedFormula.formula);
            selection.Formula = formula;
            
            // 如果需要填充
            if (this.fillRight || this.fillDown) {
                this.applyFill(selection, formula);
            }
            
            this.showNotification('公式应用成功！', 'success');
            this.clearAll();
            
        } catch (error) {
            console.error('应用公式失败:', error);
            this.showNotification('应用公式失败: ' + error.message, 'error');
        }
    }
    
    adjustFormulaForFill(formula) {
        // 根据填充方向调整公式中的相对引用
        if (this.fillRight) {
            // 水平填充：固定行号，列号相对引用
            formula = this.makeColumnsRelative(formula);
        }
        
        if (this.fillDown) {
            // 垂直填充：固定列号，行号相对引用
            formula = this.makeRowsRelative(formula);
        }
        
        return formula;
    }
    
    makeColumnsRelative(formula) {
        // 将列引用转换为相对引用（保留美元符号在行号前）
        return formula.replace(/\$([A-Z]+)/g, '$1');
    }
    
    makeRowsRelative(formula) {
        // 将行引用转换为相对引用（保留美元符号在列号前）
        return formula.replace(/(\$)([0-9]+)/g, '$2');
    }
    
    applyFill(selection, formula) {
        try {
            let targetRange;
            
            if (this.fillRight && this.fillDown) {
                // 右下角填充
                targetRange = selection.Resize(selection.Rows.Count + 1, selection.Columns.Count + 1);
            } else if (this.fillRight) {
                // 向右填充
                targetRange = selection.Resize(1, selection.Columns.Count + 1);
            } else if (this.fillDown) {
                // 向下填充
                targetRange = selection.Resize(selection.Rows.Count + 1, 1);
            }
            
            if (targetRange) {
                targetRange.Formula = formula;
            }
            
        } catch (error) {
            console.warn('填充公式失败:', error);
        }
    }
    
    clearAll() {
        // 清空表单（保留原始描述用于重新生成）
        const originalDescription = this.formulaDescription; // 保留原始描述
        document.getElementById('formulaDescription').value = '';
        document.getElementById('fillRight').checked = false;
        document.getElementById('fillDown').checked = false;
        
        // 重置状态（保留原始描述）
        this.formulaDescription = originalDescription; // 恢复原始描述
        this.fillRight = false;
        this.fillDown = false;
        this.selectedFormula = null;
        this.selectedWorksheets = [];
        
        // 清空结果
        document.getElementById('formulaSuggestions').innerHTML = '';
        document.getElementById('aiResults').style.display = 'none';
        
        // 重置按钮状态
        document.getElementById('applyFormula').disabled = true;
        document.getElementById('selectedSources').textContent = '0';
        
        // 隐藏工作表选择区域
        document.getElementById('worksheetSelection').style.display = 'none';
        
        // 重置为当前工作表模式
        document.querySelector('input[name="referenceType"][value="current"]').checked = true;
        this.referenceType = 'current';
        
        // 清空选中状态
        document.querySelectorAll('.worksheet-item').forEach(item => {
            item.classList.remove('selected');
        });
        
        this.updateStatus('已清空所有内容');
    }
    
    showLoading(show) {
        const loadingArea = document.getElementById('loadingArea');
        const generateBtn = document.getElementById('generateFormula');
        
        if (show) {
            loadingArea.style.display = 'flex';
            generateBtn.disabled = true;
            generateBtn.innerHTML = '<span class="btn-icon">⏳</span>正在生成...';
        } else {
            loadingArea.style.display = 'none';
            generateBtn.disabled = false;
            generateBtn.innerHTML = '<span class="btn-icon">🚀</span>生成公式建议';
        }
    }
    
    updateStatus(status) {
        document.getElementById('aiStatus').textContent = status;
    }
    
    showNotification(message, type = 'info') {
        const notification = document.getElementById('notification');
        const messageElement = notification.querySelector('.notification-message');
        const iconElement = notification.querySelector('.notification-icon');
        
        // 设置图标
        let icon;
        switch (type) {
            case 'success':
                icon = '✅';
                break;
            case 'error':
                icon = '❌';
                break;
            case 'warning':
                icon = '⚠️';
                break;
            default:
                icon = 'ℹ️';
        }
        
        iconElement.textContent = icon;
        messageElement.textContent = message;
        
        // 设置样式
        notification.className = `notification ${type}`;
        
        // 显示通知
        notification.style.display = 'block';
        
        // 3秒后自动隐藏
        setTimeout(() => {
            notification.style.display = 'none';
        }, 3000);
    }
    
    // 对话框相关方法
    openWorkbookModal() {
        document.getElementById('workbookModal').style.display = 'flex';
        this.loadWorkbookData();
    }
    
    closeWorkbookModal() {
        document.getElementById('workbookModal').style.display = 'none';
    }
    
    confirmWorkbookSelection() {
        const selectedWorkbooks = document.querySelectorAll('.workbook-grid-item.selected');
        if (selectedWorkbooks.length === 0) {
            this.showNotification('请至少选择一个工作簿', 'warning');
            return;
        }
        
        this.selectedWorkbooks = Array.from(selectedWorkbooks).map(item => ({
            name: item.dataset.workbookName,
            path: item.dataset.workbookPath
        }));
        
        this.updateSelectedSources();
        this.closeWorkbookModal();
    }
    
    filterWorkbooks(searchText) {
        const workbooks = document.querySelectorAll('.workbook-grid-item');
        workbooks.forEach(item => {
            const name = item.dataset.workbookName || '';
            const shouldShow = name.toLowerCase().includes(searchText.toLowerCase());
            item.style.display = shouldShow ? 'block' : 'none';
        });
    }
}

// 全局函数，供HTML调用
function closeWorkbookModal() {
    if (window.formulaGenerator) {
        window.formulaGenerator.closeWorkbookModal();
    }
}

function confirmWorkbookSelection() {
    if (window.formulaGenerator) {
        window.formulaGenerator.confirmWorkbookSelection();
    }
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', () => {
    window.formulaGenerator = new FormulaGenerator();
});