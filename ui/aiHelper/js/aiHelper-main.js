/**
 * 智能办公主入口文件 - 简化版
 * 负责加载和管理智能办公功能的所有模块
 */

var AIHelperMain = (function() {
    'use strict';
    
    function AIHelperMain() {
        this.isInitialized = false;
        this.modules = {};
        this.config = {};
        this.currentFormulas = []; // 存储当前生成的公式
        
        this.init();
    }
    
    AIHelperMain.prototype.init = function() {
        try {
            // 等待DOM加载完成
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', this.initializeModules.bind(this));
            } else {
                this.initializeModules();
            }
        } catch (error) {
            console.error('AI Helper初始化失败:', error);
        }
    };
    
    AIHelperMain.prototype.initializeModules = function() {
        try {
            // 直接使用已加载的全局模块，不进行异步加载
            this.initializeComponents();
            
            this.isInitialized = true;
            
            // 更新状态显示
            var statusElement = document.getElementById('aiStatus');
            if (statusElement) {
                statusElement.textContent = '准备就绪';
                statusElement.className = 'status-indicator success';
            }
            
        } catch (error) {
            console.error('模块初始化失败:', error);
            
            var statusElement = document.getElementById('aiStatus');
            if (statusElement) {
                statusElement.textContent = '初始化失败';
                statusElement.className = 'status-indicator error';
            }
        }
    };
    
    /**
     * 直接获取已加载的JavaScript模块
     */
    AIHelperMain.prototype.loadModule = function(moduleName) {
        try {
            switch(moduleName) {
                case 'formulaGenerator':
                    if (typeof FormulaGenerator !== 'undefined') {
                        this.modules[moduleName] = new FormulaGenerator();
                    }
                    break;
                case 'workbookSelector':
                    if (typeof WorkbookSelector !== 'undefined') {
                        this.modules[moduleName] = new WorkbookSelector();
                    }
                    break;
                case 'aiInterface':
                    if (typeof window.aiInterface !== 'undefined') {
                        this.modules[moduleName] = window.aiInterface;
                    }
                    break;
                case 'jsonSpec':
                    if (typeof window.jsonSpec !== 'undefined') {
                        this.modules[moduleName] = window.jsonSpec;
                    }
                    break;
            }
        } catch (error) {
            console.warn('模块 ' + moduleName + ' 获取失败:', error);
        }
        
        return this.modules[moduleName];
    };
    
    /**
     * 初始化各个组件
     */
    AIHelperMain.prototype.initializeComponents = function() {
        try {
            // 加载模块但不依赖它们
            this.modules.formulaGenerator = this.loadModule('formulaGenerator');
            this.modules.workbookSelector = this.loadModule('workbookSelector'); 
            this.modules.aiInterface = this.loadModule('aiInterface');
            this.modules.jsonSpec = this.loadModule('jsonSpec');
            
            // 只初始化基本的UI事件
            this.initBasicUIEvents();
            
        } catch (error) {
            console.error('组件初始化失败:', error);
            // 不抛出错误，避免卡死
        }
    };
    
    /**
     * 只初始化基本的UI事件，移除可能导致卡死的复杂逻辑
     */
    AIHelperMain.prototype.initBasicUIEvents = function() {
        var self = this;
        try {
            // 生成公式按钮
            var generateBtn = document.getElementById('generateFormula');
            if (generateBtn) {
                generateBtn.addEventListener('click', function() {
                    self.handleGenerateClick();
                });
            }
            
            // 刷新工作簿按钮
            var refreshBtn = document.getElementById('refreshWorkbooks');
            if (refreshBtn) {
                refreshBtn.addEventListener('click', function() {
                    self.handleRefreshClick();
                });
            }
            
            // 清空所有按钮
            var clearBtn = document.getElementById('clearAll');
            if (clearBtn) {
                clearBtn.addEventListener('click', function() {
                    self.handleClearClick();
                });
            }
            
        } catch (error) {
            console.error('UI事件初始化失败:', error);
        }
    };
    
    /**
     * 处理生成按钮点击
     */
    AIHelperMain.prototype.handleGenerateClick = function() {
        var self = this;
        try {
            var description = document.getElementById('formulaDescription');
            var value = description ? description.value : '';
            
            // 获取填充方向设置
            var fillDirection = 'none';
            var fillDirectionRadios = document.querySelectorAll('input[name="fillDirection"]');
            for (var i = 0; i < fillDirectionRadios.length; i++) {
                if (fillDirectionRadios[i].checked) {
                    fillDirection = fillDirectionRadios[i].value;
                    break;
                }
            }
            
            // 转换填充方向设置为旧格式
            var fillOptions = {
                right: fillDirection === 'right' || fillDirection === 'both',
                down: fillDirection === 'down' || fillDirection === 'both'
            };
            
            if (!value.trim()) {
                // 获取当前单元格信息
                this.getCurrentCellInfo().then(function(currentCellInfo) {
                    // 获取所有工作簿信息
                    var workbookInfo = self.getAllWorkbookInfo();
                    
                    // 构建完整的请求数据
                    var requestData = {
                        description: "", // 空描述，让AI根据单元格信息自行推测需求
                        referenceType: "current",
                        currentCell: currentCellInfo,
                        selectedWorkbooks: workbookInfo.selectedWorkbooks || [],
                        selectedWorksheets: workbookInfo.selectedWorksheets || [],
                        fillOptions: fillOptions
                    };
                    
                    // 发送到API
                    return self.sendFormulaRequest(requestData);
                }).catch(function(error) {
                    console.error('获取单元格信息失败:', error);
                    self.showNotification('获取单元格信息失败：' + error.message, 'error');
                });
                
            } else {
                // 使用用户输入的描述
                // 获取当前单元格信息
                this.getCurrentCellInfo().then(function(currentCellInfo) {
                    // 获取所有工作簿信息
                    var workbookInfo = self.getAllWorkbookInfo();
                    
                    // 构建完整的请求数据
                    var requestData = {
                        description: value.trim(),
                        referenceType: "current",
                        currentCell: currentCellInfo,
                        selectedWorkbooks: workbookInfo.selectedWorkbooks || [],
                        selectedWorksheets: workbookInfo.selectedWorksheets || [],
                        fillOptions: fillOptions
                    };
                    
                    // 发送到API
                    return self.sendFormulaRequest(requestData);
                }).catch(function(error) {
                    console.error('获取单元格信息失败:', error);
                    self.showNotification('获取单元格信息失败：' + error.message, 'error');
                });
            }
            
        } catch (error) {
            console.error('生成处理失败:', error);
            this.showNotification('生成失败：' + error.message, 'error');
        }
    };
    
    /**
     * 处理刷新按钮点击
     */
    AIHelperMain.prototype.handleRefreshClick = function() {
        try {
            this.showNotification('刷新完成', 'success');
        } catch (error) {
            console.error('刷新处理失败:', error);
        }
    };
    
    /**
     * 处理清空按钮点击
     */
    AIHelperMain.prototype.handleClearClick = function() {
        try {
            var description = document.getElementById('formulaDescription');
            if (description) {
                description.value = '';
            }
            
            // 清空公式结果
            this.clearFormulaResults();
            
            this.showNotification('已清空', 'success');
        } catch (error) {
            console.error('清空处理失败:', error);
        }
    };
    
    /**
     * 清空公式结果
     */
    AIHelperMain.prototype.clearFormulaResults = function() {
        try {
            // 隐藏AI结果区域
            var aiResults = document.getElementById('aiResults');
            if (aiResults) {
                aiResults.style.display = 'none';
            }
            
            // 清空公式建议
            var formulaSuggestions = document.getElementById('formulaSuggestions');
            if (formulaSuggestions) {
                formulaSuggestions.innerHTML = '';
            }
            
            // 隐藏应用公式区域
            var applyFormulaSection = document.getElementById('applyFormulaSection');
            if (applyFormulaSection) {
                applyFormulaSection.style.display = 'none';
            }
            
            // 清空当前公式
            this.currentFormulas = [];
        } catch (error) {
            console.error('清空公式结果失败:', error);
        }
    };
    
    /**
     * 获取当前单元格信息
     */
    AIHelperMain.prototype.getCurrentCellInfo = function() {
        var self = this;
        return new Promise(function(resolve, reject) {
            try {
                // 优先使用WPS JSA环境
                if (window.Application && window.Application.ActiveCell) {
                    var cell = window.Application.ActiveCell;
                    // 正确获取单元格地址
                    var address = '';
                    try {
                        // 尝试多种方式获取地址
                        if (typeof cell.Address === 'function') {
                            address = cell.Address();
                        } else if (typeof cell.Address === 'string') {
                            address = cell.Address;
                        } else {
                            address = 'A1'; // 默认值
                        }
                    } catch (addrError) {
                        address = 'A1';
                    }
                    
                    // 获取工作表和工作簿信息
                    var worksheet = window.Application.ActiveSheet;
                    var workbook = window.Application.ActiveWorkbook;
                    
                    var cellInfo = {
                        workbook: workbook ? workbook.Name : '未知工作簿',
                        worksheet: worksheet ? worksheet.Name : '未知工作表',
                        row: cell.Row || 1,
                        col: cell.Column || 1,
                        cellAddress: address,
                        value: self.extractCellValue(cell),
                        formula: cell.Formula || '',
                        numberFormat: cell.NumberFormat || '',
                        columnHeader: self.getColumnHeader(worksheet, cell.Column || 1),
                        timestamp: new Date().toISOString()
                    };
                    
                    resolve(cellInfo);
                    return;
                }
                
                // 尝试使用Office.js环境
                if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
                    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function(result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            var cellInfo = {
                                address: result.value || 'A1',
                                columnName: self.extractColumnName(result.value || 'A1'),
                                value: result.value || '',
                                rowNumber: self.extractRowNumber(result.value || 'A1'),
                                timestamp: new Date().toISOString()
                            };
                            
                            resolve(cellInfo);
                        } else {
                            reject(new Error('无法获取单元格信息'));
                        }
                    });
                    return;
                }
                
                // 开发环境回退
                resolve({
                    workbook: '示例工作簿.xlsx',
                    worksheet: 'Sheet1',
                    address: 'A1',
                    columnName: 'A',
                    value: '',
                    rowNumber: 1,
                    row: 1,
                    col: 1,
                    isDevelopmentMode: true,
                    timestamp: new Date().toISOString()
                });
                
            } catch (error) {
                console.error('❌ 获取当前单元格信息失败:', error);
                resolve({
                    workbook: '未知工作簿',
                    worksheet: '未知工作表',
                    address: 'A1',
                    columnName: 'A',
                    value: '',
                    rowNumber: 1,
                    row: 1,
                    col: 1,
                    isErrorMode: true,
                    timestamp: new Date().toISOString()
                });
            }
        });
    };

    /**
     * 获取所有工作簿信息
     */
    AIHelperMain.prototype.getAllWorkbookInfo = function() {
        try {
            // 尝试获取工作簿选择器中的工作簿信息
            if (this.modules.workbookSelector && typeof this.modules.workbookSelector.getAllWorkbooks === 'function') {
                // 获取所有工作簿（不仅仅是选中的）
                var allWorkbooks = this.modules.workbookSelector.getAllWorkbooks();
                
                // 格式化为AI接口需要的格式
                var formattedWorkbooks = allWorkbooks.map(workbook => {
                    return {
                        workBookName: workbook.name,
                        workBookPath: workbook.path,
                        worksheets: workbook.worksheets.map(worksheet => {
                            // 确保列标题格式正确
                            let columnHeaders = {};
                            if (worksheet.headers && Array.isArray(worksheet.headers)) {
                                worksheet.headers.forEach((header, index) => {
                                    const columnLetter = this.getColumnLetter(index + 1);
                                    if (header && typeof header === 'object' && header.value) {
                                        columnHeaders[columnLetter] = header.value;
                                    } else if (typeof header === 'string') {
                                        columnHeaders[columnLetter] = header;
                                    }
                                });
                            }
                            
                            return {
                                workSheetName: worksheet.name,
                                columnHeaders: columnHeaders
                            };
                        })
                    };
                });
                
                return {
                    selectedWorkbooks: formattedWorkbooks,
                    selectedWorksheets: []
                };
            }
            
            // 如果没有工作簿选择器，则尝试直接从WPS获取信息
            if (window.Application && window.Application.Workbooks) {
                var workbooks = [];
                for (var i = 1; i <= window.Application.Workbooks.Count; i++) {
                    var wb = window.Application.Workbooks.Item(i);
                    var worksheets = [];
                    
                    if (wb.Worksheets) {
                        for (var j = 1; j <= wb.Worksheets.Count; j++) {
                            var ws = wb.Worksheets.Item(j);
                            // 获取表头信息
                            var headers = this.extractWorksheetHeaders(ws);
                            
                            // 格式化列标题
                            let columnHeaders = {};
                            if (headers && Array.isArray(headers)) {
                                headers.forEach((header, index) => {
                                    const columnLetter = this.getColumnLetter(index + 1);
                                    columnHeaders[columnLetter] = header;
                                });
                            }
                            
                            worksheets.push({
                                workSheetName: ws.Name,
                                columnHeaders: columnHeaders
                            });
                        }
                    }
                    
                    workbooks.push({
                        workBookName: wb.Name,
                        workBookPath: wb.Path || '',
                        worksheets: worksheets
                    });
                }
                
                return {
                    selectedWorkbooks: workbooks,
                    selectedWorksheets: []
                };
            }
            
            return {
                selectedWorkbooks: [],
                selectedWorksheets: []
            };
        } catch (error) {
            console.error('❌ 获取工作簿信息失败:', error);
            return {
                selectedWorkbooks: [],
                selectedWorksheets: []
            };
        }
    };

    /**
     * 提取工作表表头
     */
    AIHelperMain.prototype.extractWorksheetHeaders = function(worksheet) {
        try {
            if (!worksheet || !worksheet.UsedRange) {
                return [];
            }
            
            var usedRange = worksheet.UsedRange;
            if (usedRange.Rows.Count < 1) {
                return [];
            }
            
            // 获取第一行作为表头
            var headerRow = usedRange.Rows.Item(1);
            var headers = [];
            
            for (var col = 1; col <= usedRange.Columns.Count; col++) {
                try {
                    var cell = worksheet.Cells.Item(usedRange.Row, usedRange.Column + col - 1);
                    var value = this.extractCellValue(cell);
                    headers.push(value || '列' + this.getColumnLetter(col));
                } catch (e) {
                    headers.push('列' + this.getColumnLetter(col));
                }
            }
            
            return headers;
        } catch (error) {
            return [];
        }
    };

    /**
     * 提取单元格值（兼容WPS JSA环境）
     */
    AIHelperMain.prototype.extractCellValue = function(cell) {
        try {
            if (cell.Value2 !== null && cell.Value2 !== undefined) {
                return cell.Value2;
            }
            if (cell.Value && typeof cell.Value === 'function') {
                return cell.Value();
            }
            if (cell.Value && typeof cell.Value === 'string') {
                return cell.Value;
            }
            if (cell.Text && typeof cell.Text === 'string') {
                return cell.Text;
            }
            return null;
        } catch (error) {
            return null;
        }
    };

    /**
     * 获取列标题
     */
    AIHelperMain.prototype.getColumnHeader = function(worksheet, column) {
        try {
            if (worksheet && worksheet.Cells) {
                var headerCell = worksheet.Cells.Item(1, column); // 第一行是标题行
                var headerValue = this.extractCellValue(headerCell);
                return headerValue || '列' + this.getColumnLetter(column);
            }
            return '未知列';
        } catch (error) {
            return '未知列';
        }
    };

    /**
     * 获取列号对应的字母表示 (1 -> A, 2 -> B, ..., 26 -> Z, 27 -> AA)
     */
    AIHelperMain.prototype.getColumnLetter = function(columnNumber) {
        let result = '';
        while (columnNumber > 0) {
            columnNumber--;
            result = String.fromCharCode(65 + (columnNumber % 26)) + result;
            columnNumber = Math.floor(columnNumber / 26);
        }
        return result;
    };

    /**
     * 从单元格地址提取列名
     */
    AIHelperMain.prototype.extractColumnName = function(address) {
        if (!address) return 'A';
        return address.replace(/[0-9]/g, '').toUpperCase();
    };

    /**
     * 从单元格地址提取行号
     */
    AIHelperMain.prototype.extractRowNumber = function(address) {
        if (!address) return 1;
        var match = address.match(/[0-9]+/);
        return match ? parseInt(match[0]) : 1;
    };

    /**
     * 发送公式生成请求到API
     */
    AIHelperMain.prototype.sendFormulaRequest = function(requestData) {
        var self = this;
        return new Promise(function(resolve, reject) {
            try {
                self.showNotification('正在生成公式...', 'info');
                
                // 打印发送给AI的原始数据
                console.log('📤 发送给AI的原始数据:', JSON.stringify(requestData, null, 2));
                
                // 使用增强AI接口
                if (window.enhancedAIInterface) {
                    window.enhancedAIInterface.generateFormulaRequest(requestData).then(function(result) {
                        // 打印AI响应的原始数据
                        console.log('📥 AI响应的原始数据:', JSON.stringify(result, null, 2));
                        
                        if (result.success && result.formulas && result.formulas.length > 0) {
                            // 保存当前公式
                            self.currentFormulas = result.formulas;
                            
                            // 显示公式结果
                            self.showFormulaResults(result);
                            
                            self.showNotification('公式生成成功！', 'success');
                            resolve(result);
                        } else {
                            var error = new Error('API返回结果格式错误');
                            console.error('❌ API返回结果格式错误:', result);
                            reject(error);
                        }
                    }).catch(function(error) {
                        console.error('❌ API请求失败:', error);
                        self.showNotification('API请求失败：' + error.message, 'error');
                        reject(error);
                    });
                } else {
                    // 如果没有增强AI接口，使用简单模拟
                    var error = new Error('AI接口未初始化');
                    console.error('❌ AI接口未初始化');
                    reject(error);
                }
                
            } catch (error) {
                console.error('❌ API请求异常:', error);
                self.showNotification('API请求异常：' + error.message, 'error');
                reject(error);
            }
        });
    };

    /**
     * 显示公式结果
     */
    AIHelperMain.prototype.showFormulaResults = function(result) {
        try {
            // 显示AI结果区域
            var aiResults = document.getElementById('aiResults');
            if (aiResults) {
                aiResults.style.display = 'block';
            }
            
            // 显示公式建议
            var formulaSuggestions = document.getElementById('formulaSuggestions');
            if (formulaSuggestions) {
                formulaSuggestions.innerHTML = '';
                
                result.formulas.forEach(function(formula, index) {
                    var formulaItem = document.createElement('div');
                    formulaItem.className = 'formula-item';
                    formulaItem.innerHTML = `
                        <div class="formula-header">
                            <h4>${formula.title || '推荐公式'}</h4>
                            <span class="confidence">置信度: ${formula.confidence || 0}%</span>
                        </div>
                        <div class="formula-content">
                            <div class="formula-code">${formula.formula || '无公式'}</div>
                            <div class="formula-explanation">${formula.explanation || '无说明'}</div>
                            ${formula.applicable_ranges ? `<div class="formula-ranges">适用范围: ${formula.applicable_ranges.join(', ')}</div>` : ''}
                            ${formula.required_functions ? `<div class="formula-functions">所需函数: ${formula.required_functions.join(', ')}</div>` : ''}
                            ${formula.example ? `<div class="formula-example">示例: ${formula.example}</div>` : ''}
                        </div>
                    `;
                    formulaSuggestions.appendChild(formulaItem);
                });
                
                // 显示数据分析信息
                if (result.data_analysis) {
                    var analysisDiv = document.createElement('div');
                    analysisDiv.className = 'data-analysis';
                    analysisDiv.innerHTML = `
                        <h4>📊 数据分析</h4>
                        ${result.data_analysis.smart_analysis ? `<div class="smart-analysis">${result.data_analysis.smart_analysis}</div>` : ''}
                        ${result.data_analysis.recommendations ? `
                            <div class="recommendations">
                                <h5>建议:</h5>
                                <ul>
                                    ${result.data_analysis.recommendations.map(rec => `<li>${rec}</li>`).join('')}
                                </ul>
                            </div>
                        ` : ''}
                        ${result.data_analysis.headers_found ? `
                            <details class="headers-details">
                                <summary>发现的表头 (${result.data_analysis.headers_found.length} 个)</summary>
                                <div class="headers-list">
                                    ${result.data_analysis.headers_found.map(header => `<span class="header-item">${header}</span>`).join('')}
                                </div>
                            </details>
                        ` : ''}
                    `;
                    formulaSuggestions.appendChild(analysisDiv);
                }
            }
            
            // 显示替代公式
            if (result.alternative_formulas && result.alternative_formulas.length > 0) {
                var alternativesDiv = document.createElement('div');
                alternativesDiv.className = 'alternative-formulas';
                alternativesDiv.innerHTML = `
                    <h4>🔄 替代方案</h4>
                    ${result.alternative_formulas.map((alt, index) => `
                        <div class="alternative-item">
                            <div class="alternative-header">
                                <h5>${alt.description}</h5>
                            </div>
                            <div class="alternative-content">
                                <div class="alternative-formula">${alt.formula}</div>
                                ${alt.pros ? `
                                    <div class="alternative-pros">
                                        <strong>优点:</strong>
                                        <ul>
                                            ${alt.pros.map(pro => `<li>${pro}</li>`).join('')}
                                        </ul>
                                    </div>
                                ` : ''}
                                ${alt.cons ? `
                                    <div class="alternative-cons">
                                        <strong>缺点:</strong>
                                        <ul>
                                            ${alt.cons.map(con => `<li>${con}</li>`).join('')}
                                        </ul>
                                    </div>
                                ` : ''}
                            </div>
                        </div>
                    `).join('')}
                `;
                formulaSuggestions.appendChild(alternativesDiv);
            }
            
            // 显示应用公式区域
            this.showApplyFormulaOptions(result.formulas);
            
        } catch (error) {
            console.error('显示公式结果失败:', error);
        }
    };

    /**
     * 显示应用公式选项
     */
    AIHelperMain.prototype.showApplyFormulaOptions = function(formulas) {
        try {
            var applyFormulaSection = document.getElementById('applyFormulaSection');
            var formulaApplyOptions = document.getElementById('formulaApplyOptions');
            
            if (applyFormulaSection && formulaApplyOptions) {
                // 显示区域
                applyFormulaSection.style.display = 'block';
                formulaApplyOptions.innerHTML = '';
                
                // 按置信度排序（从高到低）
                var sortedFormulas = formulas.slice().sort(function(a, b) {
                    return (b.confidence || 0) - (a.confidence || 0);
                });
                
                // 创建应用选项
                sortedFormulas.forEach(function(formula, index) {
                    var optionButton = document.createElement('button');
                    optionButton.className = 'btn-formula-option';
                    optionButton.innerHTML = `
                        <div class="option-header">
                            <span class="option-title">${formula.title || '推荐公式'}</span>
                            <span class="option-confidence">${formula.confidence || 0}%</span>
                        </div>
                        <div class="option-formula">${formula.formula || '无公式'}</div>
                    `;
                    optionButton.onclick = function() {
                        this.applyFormula(formula.formula);
                    }.bind(this);
                    
                    formulaApplyOptions.appendChild(optionButton);
                }.bind(this));
            }
        } catch (error) {
            console.error('显示应用公式选项失败:', error);
        }
    };

    /**
     * 应用公式到当前单元格
     */
    AIHelperMain.prototype.applyFormula = function(formula) {
        try {
            if (window.Application && window.Application.Selection) {
                window.Application.Selection.Formula = formula;
                this.showNotification('公式已应用到当前单元格', 'success');
            } else {
                this.showNotification('无法访问Excel对象模型', 'warning');
            }
            
        } catch (error) {
            console.error('应用公式失败:', error);
            this.showNotification('应用公式失败：' + error.message, 'error');
        }
    };

    /**
     * 处理应用按钮点击
     */
    AIHelperMain.prototype.handleApplyClick = function() {
        try {
            this.showNotification('应用功能开发中...', 'info');
        } catch (error) {
            console.error('应用处理失败:', error);
        }
    };
    
    /**
     * 显示通知消息
     */
    AIHelperMain.prototype.showNotification = function(message, type) {
        try {
            // 创建简单的通知元素
            var notification = document.createElement('div');
            notification.className = 'notification notification-' + (type || 'info');
            notification.style.cssText = 
                'position: fixed; ' +
                'top: 20px; ' +
                'right: 20px; ' +
                'padding: 10px 20px; ' +
                'background: ' + (type === 'success' ? '#d4edda' : type === 'error' ? '#f8d7da' : '#d1ecf1') + '; ' +
                'color: ' + (type === 'success' ? '#155724' : type === 'error' ? '#721c24' : '#0c5460') + '; ' +
                'border: 1px solid ' + (type === 'success' ? '#c3e6cb' : type === 'error' ? '#f5c6cb' : '#bee5eb') + '; ' +
                'border-radius: 5px; ' +
                'z-index: 10000; ' +
                'font-size: 14px; ' +
                'max-width: 300px; ' +
                'word-wrap: break-word;';
            notification.textContent = message;
            
            document.body.appendChild(notification);
            
            // 3秒后自动隐藏
            setTimeout(function() {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 3000);
            
        } catch (error) {
            console.error('显示通知失败:', error);
        }
    };
    
    /**
     * 获取系统状态
     */
    AIHelperMain.prototype.getSystemStatus = function() {
        return {
            initialized: this.isInitialized,
            modules: Object.keys(this.modules),
            timestamp: new Date().toISOString()
        };
    };
    
    return AIHelperMain;
})();

// 创建全局实例
window.AIHelperMain = AIHelperMain;

// 页面加载完成后自动初始化
var aiHelperInstance = null;

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function() {
        aiHelperInstance = new AIHelperMain();
        // 设置全局实例引用，供按钮点击事件使用
        window.aiHelperMainInstance = aiHelperInstance;
    });
} else {
    aiHelperInstance = new AIHelperMain();
    // 设置全局实例引用，供按钮点击事件使用
    window.aiHelperMainInstance = aiHelperInstance;
}

// 导出到全局作用域
window.aiHelper = aiHelperInstance;