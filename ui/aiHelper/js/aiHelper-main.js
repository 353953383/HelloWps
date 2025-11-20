/**
 * 智能办公主入口文件
 * 负责加载和管理智能办公功能的所有模块
 */

class AIHelperMain {
    constructor() {
        this.isInitialized = false;
        this.modules = {};
        this.config = {};
        
        this.init();
    }
    
    init() {
        try {
            // 等待DOM加载完成
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', () => this.initializeModules());
            } else {
                this.initializeModules();
            }
        } catch (error) {
            console.error('AI Helper初始化失败:', error);
            this.showError('初始化失败，请刷新页面重试');
        }
    }
    
    async initializeModules() {
        try {
            this.showLoading();
            
            // 加载基础模块
            await this.loadModule('jsonSpec', '/js/aiHelper/jsonSpec.js');
            await this.loadModule('aiInterface', '/js/aiHelper/aiInterface.js');
            await this.loadModule('workbookSelector', '/js/aiHelper/workbookSelector.js');
            await this.loadModule('formulaGenerator', '/js/aiHelper/formulaGenerator.js');
            
            // 初始化各个模块
            this.initializeComponents();
            
            // 设置事件监听
            this.setupEventListeners();
            
            this.isInitialized = true;
            this.hideLoading();
            this.showSuccess('智能办公系统初始化完成');
            
        } catch (error) {
            console.error('模块加载失败:', error);
            this.hideLoading();
            this.showError('模块加载失败: ' + error.message);
        }
    }
    
    /**
     * 动态加载JavaScript模块
     */
    async loadModule(moduleName, scriptPath) {
        if (this.modules[moduleName]) {
            return this.modules[moduleName];
        }
        
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = scriptPath;
            script.async = true;
            
            script.onload = () => {
                resolve(this.modules[moduleName]);
            };
            
            script.onerror = () => {
                console.error(`模块 ${moduleName} 加载失败`);
                reject(new Error(`无法加载模块: ${moduleName}`));
            };
            
            document.head.appendChild(script);
        });
    }
    
    /**
     * 初始化各个组件
     */
    initializeComponents() {
        try {
            // 初始化公式生成器
            if (typeof FormulaGenerator !== 'undefined') {
                this.modules.formulaGenerator = new FormulaGenerator();
            }
            
            // 初始化工作簿选择器
            if (typeof WorkbookSelector !== 'undefined') {
                this.modules.workbookSelector = new WorkbookSelector();
            }
            
            // 初始化AI接口
            if (typeof AIInterface !== 'undefined') {
                this.modules.aiInterface = window.aiInterface;
            }
            
            // 初始化界面交互
            this.initUIEvents();
            
        } catch (error) {
            console.error('组件初始化失败:', error);
            throw error;
        }
    }
    
    /**
     * 设置事件监听器
     */
    setupEventListeners() {
        // 监听来自WPS的事件
        document.addEventListener('wps-ready', () => {
            this.refreshStatus();
        });
        
        // 监听配置更新事件
        document.addEventListener('ai-config-updated', (e) => {
            this.updateConfig(e.detail);
        });
        
        // 监听键盘快捷键
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'Enter') {
                e.preventDefault();
                this.handleQuickFormula();
            }
        });
    }
    
    /**
     * 初始化UI事件
     */
    initUIEvents() {
        // 公式需求输入
        const requirementInput = document.getElementById('requirementInput');
        if (requirementInput) {
            requirementInput.addEventListener('input', this.debounce((e) => {
                this.updateFormulaPreview(e.target.value);
            }, 500));
        }
        
        // 引用类型变化
        const referenceTypeRadios = document.querySelectorAll('input[name="referenceType"]');
        referenceTypeRadios.forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.handleReferenceTypeChange(e.target.value);
            });
        });
        
        // 填充选项变化
        const fillOptions = document.querySelectorAll('input[name="fillOption"]');
        fillOptions.forEach(option => {
            option.addEventListener('change', () => {
                this.updateFillOptions();
            });
        });
        
        // 生成公式按钮
        const generateBtn = document.getElementById('generateBtn');
        if (generateBtn) {
            generateBtn.addEventListener('click', () => {
                this.generateFormulas();
            });
        }
        
        // 应用公式按钮
        const applyBtn = document.getElementById('applyBtn');
        if (applyBtn) {
            applyBtn.addEventListener('click', () => {
                this.applySelectedFormula();
            });
        }
        
        // 设置按钮
        const settingsBtn = document.getElementById('settingsBtn');
        if (settingsBtn) {
            settingsBtn.addEventListener('click', () => {
                this.showSettings();
            });
        }
        
        // 工作簿选择按钮
        const selectWorkbookBtn = document.getElementById('selectWorkbookBtn');
        if (selectWorkbookBtn) {
            selectWorkbookBtn.addEventListener('click', () => {
                this.selectWorkbooks();
            });
        }
    }
    
    /**
     * 处理引用类型变化
     */
    handleReferenceTypeChange(type) {
        const workbookSelector = document.getElementById('workbookSelector');
        const currentWorksheetInfo = document.getElementById('currentWorksheetInfo');
        
        if (type === 'current') {
            if (workbookSelector) workbookSelector.style.display = 'none';
            if (currentWorksheetInfo) currentWorksheetInfo.style.display = 'block';
        } else {
            if (workbookSelector) workbookSelector.style.display = 'block';
            if (currentWorksheetInfo) currentWorksheetInfo.style.display = 'none';
        }
        
        // 更新工作簿选择器
        if (this.modules.workbookSelector) {
            this.modules.workbookSelector.updateReferenceType(type);
        }
    }
    
    /**
     * 生成公式建议
     */
    async generateFormulas() {
        try {
            this.showGenerating();
            
            // 获取用户输入
            const requestData = this.collectRequestData();
            
            // 验证输入数据
            const validation = AIJsonValidator.validateRequest(requestData);
            if (!validation.isValid) {
                throw new Error('输入数据验证失败: ' + validation.errors.join(', '));
            }
            
            // 调用AI接口生成公式
            const response = await this.modules.aiInterface.generateFormula(requestData);
            
            // 显示结果
            this.displayFormulaResults(response);
            
            this.hideGenerating();
            
        } catch (error) {
            console.error('公式生成失败:', error);
            this.hideGenerating();
            this.showError('公式生成失败: ' + error.message);
        }
    }
    
    /**
     * 收集请求数据
     */
    collectRequestData() {
        const requirementInput = document.getElementById('requirementInput');
        const referenceType = document.querySelector('input[name="referenceType"]:checked');
        const fillOptions = {
            right: document.getElementById('fillRight').checked,
            down: document.getElementById('fillDown').checked
        };
        
        // 获取当前单元格信息
        let currentCell = {};
        let workbookInfo = {};
        try {
            if (window.Application && window.Application.ActiveSheet) {
                const activeCell = window.Application.ActiveSheet.ActiveCell;
                const activeWorkbook = window.Application.ActiveWorkbook;
                const activeSheet = window.Application.ActiveSheet;
                
                currentCell = {
                    cellAddress: activeCell.Address,
                    row: activeCell.Row,
                    column: activeCell.Column,
                    worksheet: activeSheet.Name
                };
                
                // 获取完整的工作簿信息
                workbookInfo = this.getCurrentWorkbookInfo();
            }
        } catch (error) {
            console.warn('无法获取当前单元格信息:', error);
        }
        
        const requestData = {
            description: requirementInput ? requirementInput.value : '',
            referenceType: referenceType ? referenceType.value : 'current',
            currentCell: currentCell,
            selectedWorkbooks: this.modules.workbookSelector ? this.modules.workbookSelector.getSelectedWorkbooks() : [],
            selectedWorksheets: this.modules.workbookSelector ? this.modules.workbookSelector.getSelectedWorksheets() : [],
            fillOptions: fillOptions,
            headers: this.modules.workbookSelector ? this.modules.workbookSelector.getHeadersInfo() : [],
            // 新增完整的工作表信息
            currentWorkbook: workbookInfo.currentWorkbook,
            currentWorksheet: workbookInfo.currentWorksheet,
            allWorksheets: workbookInfo.allWorksheets,
            columnHeaders: workbookInfo.columnHeaders
        };
        
        return requestData;
    }
    
    /**
     * 获取当前工作簿的完整信息 (与 formulaGenerator.js 中的实现保持一致)
     */
    getCurrentWorkbookInfo() {
        console.log('🔍 开始获取当前工作簿信息...');
        
        try {
            // 检查Excel COM对象是否可用
            if (!window.Application) {
                console.warn('⚠️ window.Application不可用，可能在Web环境中');
                return this.getFallbackWorkbookInfo();
            }
            
            const activeWorkbook = window.Application.ActiveWorkbook;
            const activeSheet = window.Application.ActiveSheet;
            
            if (!activeWorkbook) {
                console.warn('⚠️ 无法获取ActiveWorkbook');
                return this.getFallbackWorkbookInfo();
            }
            
            if (!activeSheet) {
                console.warn('⚠️ 无法获取ActiveSheet');
                return this.getFallbackWorkbookInfo();
            }
            
            // 获取当前工作簿信息
            const currentWorkbook = {
                name: activeWorkbook ? activeWorkbook.Name : '未知工作簿'
            };
            console.log('📁 工作簿名称:', currentWorkbook.name);
            
            // 获取所有工作表信息
            const allWorksheets = [];
            if (activeWorkbook && activeWorkbook.Worksheets) {
                console.log(`📊 开始处理工作簿中的 ${activeWorkbook.Worksheets.Count} 个工作表...`);
                for (let i = 1; i <= activeWorkbook.Worksheets.Count; i++) {
                    try {
                        const ws = activeWorkbook.Worksheets.Item(i);
                        if (ws) {
                            const usedRange = ws.UsedRange;
                            const sheetInfo = {
                                name: ws.Name,
                                usedRange: usedRange ? {
                                    rows: usedRange.Rows.Count,
                                    columns: usedRange.Columns.Count
                                } : { rows: 0, columns: 0 }
                            };
                            allWorksheets.push(sheetInfo);
                            console.log(`  ✅ 工作表${i}: ${ws.Name} (${sheetInfo.usedRange.rows}x${sheetInfo.usedRange.columns})`);
                        }
                    } catch (error) {
                        console.warn(`⚠️ 处理工作表 ${i} 失败:`, error);
                    }
                }
            }
            console.log(`📋 成功获取 ${allWorksheets.length} 个工作表信息`);
            
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
                console.log(`📄 当前工作表: ${currentWorksheet.name} (${currentWorksheet.usedRange.rows}x${currentWorksheet.usedRange.columns})`);
            } else {
                console.log(`📄 当前工作表: ${currentWorksheet.name} (无法获取使用范围)`);
            }
            
            // 获取当前工作表的所有列标题
            let columnHeaders = [];
            if (activeSheet) {
                console.log('🔍 开始提取表头信息...');
                columnHeaders = this.extractHeaders(activeSheet);
                console.log(`📊 提取到 ${columnHeaders.length} 个表头:`, columnHeaders.slice(0, 5));
            }
            
            const result = {
                currentWorkbook,
                currentWorksheet,
                allWorksheets,
                columnHeaders
            };
            
            console.log('✅ 工作簿信息获取完成:', result);
            return result;
            
        } catch (error) {
            console.error('❌ 获取当前工作簿信息失败:', error);
            return this.getFallbackWorkbookInfo();
        }
    }
    
    /**
     * 获取备用工作簿信息 (当COM对象不可用时)
     */
    getFallbackWorkbookInfo() {
        console.log('🔄 使用备用工作簿信息获取策略...');
        
        try {
            // 尝试通过ActiveCell获取基本信息
            let basicInfo = {};
            if (window.Application && window.Application.ActiveSheet) {
                const activeCell = window.Application.ActiveSheet.ActiveCell;
                const activeSheet = window.Application.ActiveSheet;
                
                basicInfo = {
                    cellAddress: activeCell ? activeCell.Address : '未知',
                    worksheet: activeSheet ? activeSheet.Name : '未知工作表'
                };
                console.log('📍 基本单元格信息:', basicInfo);
            }
            
            // 构建备用信息结构
            const fallbackInfo = {
                currentWorkbook: {
                    name: window.Application?.ActiveWorkbook?.Name || '工作簿信息不可用'
                },
                currentWorksheet: {
                    name: basicInfo.worksheet || '工作表信息不可用',
                    usedRange: { rows: 0, columns: 0 }
                },
                allWorksheets: [], // 在Web环境中无法获取所有工作表
                columnHeaders: this.getFallbackHeaders(basicInfo.worksheet)
            };
            
            console.log('✅ 备用信息构建完成:', fallbackInfo);
            return fallbackInfo;
            
        } catch (error) {
            console.error('❌ 备用信息获取也失败:', error);
            return {
                currentWorkbook: { name: '信息获取失败' },
                currentWorksheet: { name: '信息获取失败', usedRange: { rows: 0, columns: 0 } },
                allWorksheets: [],
                columnHeaders: []
            };
        }
    }
    
    /**
     * 获取备用表头信息
     */
    getFallbackHeaders(worksheetName) {
        console.log(`🔍 为工作表"${worksheetName}"生成备用表头...`);
        
        // 根据工作表名称推测可能的表头
        if (worksheetName && worksheetName.includes('库存')) {
            console.log('📦 检测到库存相关工作表，生成预设表头');
            return [
                '模块编号', '模块名称', '库存数量', '安全库存', '库存金额',
                '供应商', '入库日期', '出库日期', '库存状态', '备注'
            ];
        }
        
        if (worksheetName && worksheetName.includes('模块说明')) {
            console.log('📋 检测到模块说明工作表，生成预设表头');
            return [
                '模块编号', '模块名称', '模块类型', '功能描述', '参数说明',
                '安装位置', '维护周期', '技术规格', '供应商信息', '备注'
            ];
        }
        
        // 默认表头
        console.log('📊 生成默认表头');
        return [
            '列1', '列2', '列3', '列4', '列5',
            '列6', '列7', '列8', '列9', '列10'
        ];
    }
    
    /**
     * 提取工作表表头信息 (与 formulaGenerator.js 中的实现保持一致)
     */
    extractHeaders(worksheet) {
        try {
            // 优先从第一行获取表头
            const firstRow = worksheet.Rows.Item(1);
            if (firstRow) {
                const usedColumns = firstRow.Columns.Count;
                const headers = [];
                
                // 获取所有非空列的标题
                for (let col = 1; col <= usedColumns; col++) {
                    try {
                        const cell = firstRow.Cells.Item(1, col);
                        let cellValue = '';
                        
                        if (cell && cell.Value !== null && cell.Value !== undefined) {
                            // 处理不同的数据类型
                            if (typeof cell.Value === 'string') {
                                cellValue = cell.Value.trim();
                            } else if (typeof cell.Value === 'number') {
                                cellValue = cell.Value.toString();
                            } else if (cell.Value instanceof Date) {
                                cellValue = cell.Value.toLocaleDateString();
                            } else {
                                cellValue = cell.Value.toString();
                            }
                        }
                        
                        headers.push(cellValue || `列${col}`);
                    } catch (error) {
                        console.warn(`获取第${col}列表头失败:`, error);
                        headers.push(`列${col}`);
                    }
                }
                
                // 如果第一行都是空值，尝试查找实际的数据行
                const hasValidHeaders = headers.some(header => header !== '' && header !== `列1`);
                if (!hasValidHeaders) {
                    return this.extractHeadersFromDataRange(worksheet);
                }
                
                return headers;
            }
        } catch (error) {
            console.warn('提取表头失败:', error);
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
                // 尝试第一行到第五行，查找第一个非空行作为表头
                const maxRowToCheck = Math.min(5, usedRange.Rows.Count);
                const headers = [];
                
                for (let row = 1; row <= maxRowToCheck; row++) {
                    const headerRow = usedRange.Rows.Item(row);
                    let hasData = false;
                    const rowHeaders = [];
                    
                    for (let col = 1; col <= headerRow.Columns.Count; col++) {
                        try {
                            const cell = headerRow.Cells.Item(1, col);
                            let cellValue = '';
                            
                            if (cell && cell.Value !== null && cell.Value !== undefined) {
                                if (typeof cell.Value === 'string') {
                                    cellValue = cell.Value.trim();
                                } else {
                                    cellValue = cell.Value.toString();
                                }
                            }
                            
                            rowHeaders.push(cellValue || `列${col}`);
                            if (cellValue !== '') hasData = true;
                        } catch (error) {
                            rowHeaders.push(`列${col}`);
                        }
                    }
                    
                    if (hasData) {
                        return rowHeaders;
                    }
                }
            }
        } catch (error) {
            console.warn('从数据范围提取表头失败:', error);
        }
        
        return [];
    }
    
    /**
     * 显示公式结果
     */
    displayFormulaResults(response) {
        const resultsContainer = document.getElementById('formulaResults');
        const alternativeContainer = document.getElementById('alternativeResults');
        
        if (resultsContainer) {
            resultsContainer.innerHTML = '';
            
            if (response.formulas && response.formulas.length > 0) {
                response.formulas.forEach((formula, index) => {
                    const formulaElement = this.createFormulaElement(formula, index);
                    resultsContainer.appendChild(formulaElement);
                });
            } else {
                resultsContainer.innerHTML = '<p class="no-results">未找到合适的公式建议</p>';
            }
        }
        
        if (alternativeContainer) {
            alternativeContainer.innerHTML = '';
            
            if (response.alternative_formulas && response.alternative_formulas.length > 0) {
                response.alternative_formulas.forEach((altFormula, index) => {
                    const altElement = this.createAlternativeElement(altFormula, index);
                    alternativeContainer.appendChild(altElement);
                });
            }
        }
        
        // 显示数据分析结果
        if (response.data_analysis) {
            this.displayDataAnalysis(response.data_analysis);
        }
        
        // 显示使用统计
        if (response.metadata) {
            this.displayMetadata(response.metadata);
        }
    }
    
    /**
     * 创建公式元素
     */
    createFormulaElement(formula, index) {
        const element = document.createElement('div');
        element.className = 'formula-item';
        element.innerHTML = `
            <div class="formula-header">
                <h4>${formula.title}</h4>
                <div class="confidence-badge confidence-${Math.floor(formula.confidence / 20)}">
                    置信度: ${formula.confidence}%
                </div>
            </div>
            <div class="formula-content">
                <div class="formula-text">${formula.formula}</div>
                <div class="formula-explanation">${formula.explanation}</div>
                <div class="formula-meta">
                    <span class="functions">函数: ${formula.required_functions.join(', ')}</span>
                    <span class="applicable-ranges">适用: ${formula.applicable_ranges.join(', ')}</span>
                </div>
                <button class="select-formula-btn" data-index="${index}">选择此公式</button>
            </div>
        `;
        
        // 添加选择事件
        const selectBtn = element.querySelector('.select-formula-btn');
        selectBtn.addEventListener('click', () => {
            this.selectFormula(formula, index);
        });
        
        return element;
    }
    
    /**
     * 创建替代方案元素
     */
    createAlternativeElement(altFormula, index) {
        const element = document.createElement('div');
        element.className = 'alternative-item';
        element.innerHTML = `
            <div class="alternative-header">
                <h4>${altFormula.description}</h4>
            </div>
            <div class="alternative-content">
                <div class="alternative-formula">${altFormula.formula}</div>
                <div class="pros-cons">
                    <div class="pros">
                        <strong>优点:</strong>
                        <ul>${altFormula.pros.map(pro => `<li>${pro}</li>`).join('')}</ul>
                    </div>
                    <div class="cons">
                        <strong>缺点:</strong>
                        <ul>${altFormula.cons.map(con => `<li>${con}</li>`).join('')}</ul>
                    </div>
                </div>
            </div>
        `;
        
        return element;
    }
    
    /**
     * 选择公式
     */
    selectFormula(formula, index) {
        // 移除之前选中的样式
        document.querySelectorAll('.formula-item').forEach(item => {
            item.classList.remove('selected');
        });
        
        // 添加当前选中的样式
        const selectedElement = document.querySelector(`[data-index="${index}"]`).closest('.formula-item');
        selectedElement.classList.add('selected');
        
        // 保存选中的公式
        this.selectedFormula = formula;
        
        // 更新应用按钮状态
        const applyBtn = document.getElementById('applyBtn');
        if (applyBtn) {
            applyBtn.disabled = false;
            applyBtn.textContent = '应用此公式';
        }
        
        this.showSuccess(`已选择公式: ${formula.title}`);
    }
    
    /**
     * 应用选中的公式
     */
    applySelectedFormula() {
        if (!this.selectedFormula) {
            this.showError('请先选择一个公式');
            return;
        }
        
        try {
            // 获取当前选中的单元格或范围
            let targetRange = null;
            try {
                if (window.Application && window.Application.ActiveSheet) {
                    targetRange = window.Application.ActiveSheet.Selection;
                }
            } catch (error) {
                console.error('无法获取目标范围:', error);
            }
            
            if (!targetRange) {
                // 如果没有选中范围，获取当前活动单元格
                try {
                    if (window.Application && window.Application.ActiveSheet) {
                        targetRange = window.Application.ActiveSheet.ActiveCell;
                    }
                } catch (error) {
                    console.error('无法获取活动单元格:', error);
                }
            }
            
            if (!targetRange) {
                throw new Error('无法确定目标单元格位置');
            }
            
            // 应用公式
            this.applyFormulaToRange(this.selectedFormula.formula, targetRange);
            
            // 如果需要填充，处理填充逻辑
            this.handleFillOperations(targetRange);
            
            this.showSuccess('公式应用成功！');
            
        } catch (error) {
            console.error('应用公式失败:', error);
            this.showError('应用公式失败: ' + error.message);
        }
    }
    
    /**
     * 将公式应用到指定范围
     */
    applyFormulaToRange(formula, range) {
        try {
            // 设置公式
            range.Formula = formula;
            
            // 如果有多个单元格，应用后进行格式设置
            if (range.Cells.Count > 1) {
                // 可以在这里添加格式设置逻辑
            }
            
        } catch (error) {
            console.error('设置公式失败:', error);
            throw new Error('无法设置公式到选中范围');
        }
    }
    
    /**
     * 处理填充操作
     */
    handleFillOperations(targetRange) {
        try {
            const fillRight = document.getElementById('fillRight').checked;
            const fillDown = document.getElementById('fillDown').checked;
            
            if (fillRight || fillDown) {
                // 计算填充范围
                let fillRange = targetRange;
                
                if (fillRight) {
                    // 向右填充
                    // 这里需要根据公式的具体内容来调整填充逻辑
                }
                
                if (fillDown) {
                    // 向下填充
                    // 这里需要根据公式的具体内容来调整填充逻辑
                }
            }
        } catch (error) {
            console.warn('填充操作失败:', error);
            // 填充失败不影响主功能，继续执行
        }
    }
    
    /**
     * 快速公式生成（Ctrl+Enter快捷键）
     */
    handleQuickFormula() {
        if (!this.isInitialized) {
            this.showError('系统尚未初始化完成，请稍候');
            return;
        }
        
        this.generateFormulas();
    }
    
    /**
     * 刷新状态
     */
    refreshStatus() {
        // 更新当前工作表信息
        try {
            if (window.Application && window.Application.ActiveSheet) {
                const activeSheet = window.Application.ActiveSheet;
                const infoElement = document.getElementById('currentWorksheetInfo');
                
                if (infoElement) {
                    infoElement.innerHTML = `
                        <div class="current-info">
                            <p><strong>当前工作表:</strong> ${activeSheet.Name}</p>
                            <p><strong>使用范围:</strong> ${activeSheet.UsedRange.Rows.Count} 行 x ${activeSheet.UsedRange.Columns.Count} 列</p>
                        </div>
                    `;
                }
            }
        } catch (error) {
            console.warn('刷新工作表信息失败:', error);
        }
    }
    
    /**
     * 显示设置面板
     */
    showSettings() {
        // TODO: 实现设置面板
        this.showInfo('设置功能正在开发中...');
    }
    
    /**
     * 选择工作簿
     */
    selectWorkbooks() {
        if (this.modules.workbookSelector) {
            this.modules.workbookSelector.openSelector();
        }
    }
    
    /**
     * 显示加载状态
     */
    showLoading() {
        const loadingElement = document.getElementById('loadingIndicator');
        if (loadingElement) {
            loadingElement.style.display = 'block';
        }
    }
    
    /**
     * 隐藏加载状态
     */
    hideLoading() {
        const loadingElement = document.getElementById('loadingIndicator');
        if (loadingElement) {
            loadingElement.style.display = 'none';
        }
    }
    
    /**
     * 显示生成中状态
     */
    showGenerating() {
        const generatingElement = document.getElementById('generatingIndicator');
        if (generatingElement) {
            generatingElement.style.display = 'block';
        }
        
        const generateBtn = document.getElementById('generateBtn');
        if (generateBtn) {
            generateBtn.disabled = true;
            generateBtn.textContent = '生成中...';
        }
    }
    
    /**
     * 隐藏生成中状态
     */
    hideGenerating() {
        const generatingElement = document.getElementById('generatingIndicator');
        if (generatingElement) {
            generatingElement.style.display = 'none';
        }
        
        const generateBtn = document.getElementById('generateBtn');
        if (generateBtn) {
            generateBtn.disabled = false;
            generateBtn.textContent = '生成公式建议';
        }
    }
    
    /**
     * 显示成功消息
     */
    showSuccess(message) {
        this.showNotification(message, 'success');
    }
    
    /**
     * 显示错误消息
     */
    showError(message) {
        this.showNotification(message, 'error');
    }
    
    /**
     * 显示信息消息
     */
    showInfo(message) {
        this.showNotification(message, 'info');
    }
    
    /**
     * 显示通知
     */
    showNotification(message, type = 'info') {
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <div class="notification-content">
                <span class="notification-icon">${this.getNotificationIcon(type)}</span>
                <span class="notification-message">${message}</span>
                <button class="notification-close" onclick="this.parentElement.parentElement.remove()">×</button>
            </div>
        `;
        
        document.body.appendChild(notification);
        
        // 自动隐藏
        setTimeout(() => {
            if (notification.parentNode) {
                notification.remove();
            }
        }, 5000);
    }
    
    /**
     * 获取通知图标
     */
    getNotificationIcon(type) {
        const icons = {
            success: '✓',
            error: '✗',
            info: 'ℹ',
            warning: '⚠'
        };
        return icons[type] || icons.info;
    }
    
    /**
     * 防抖函数
     */
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }
    
    /**
     * 获取系统状态
     */
    getSystemStatus() {
        return {
            initialized: this.isInitialized,
            modules: Object.keys(this.modules),
            config: this.config,
            timestamp: new Date().toISOString()
        };
    }
}

// 创建全局实例
window.AIHelperMain = AIHelperMain;

// 页面加载完成后自动初始化
let aiHelperInstance = null;

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        aiHelperInstance = new AIHelperMain();
    });
} else {
    aiHelperInstance = new AIHelperMain();
}

// 导出到全局作用域
window.aiHelper = aiHelperInstance;