/**
 * 智能公式生成器 - Web兼容版本
 * 负责处理用户交互、公式生成和数据处理
 * 专门为Web环境设计，避免Excel COM对象调用
 */

// 使用立即执行函数包装，确保兼容WPS JSA环境
var FormulaGenerator = (function() {
    'use strict';
    
    function FormulaGenerator() {
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
        
        // 环境检测
        this.isExcelEnvironment = this.detectExcelEnvironment();
        
        if (!this.isExcelEnvironment) {
            console.log('🌐 [FormulaGenerator] 检测到Web环境，使用兼容模式');
        }
        
        // 优先使用标准AI接口（严格遵循AIapi.txt规范）
        this.aiInterface = this.getPreferredAIInterface();
        this.standardApi = this.getStandardApi();
        
        // 简化日志输出
        /*
        const standardApiName = this.standardApi ? this.standardApi.constructor.name : '未加载';
        const enhancedApiName = this.aiInterface ? this.aiInterface.constructor.name : 'null';
        if (standardApiName !== '未加载' || enhancedApiName !== 'null') {
            console.log('✅ AI接口初始化成功');
        } else {
            console.log('❌ AI接口初始化失败');
        }
        */
        
        this.init();
    }
    
    /**
     * 检测是否在Excel环境中运行
     */
    FormulaGenerator.prototype.detectExcelEnvironment = function() {
        var hasExcelObjects = 
            window.Application && 
            window.Application.Workbooks &&
            typeof window.Application.Workbooks.Count === 'number';
            
        if (hasExcelObjects) {
            return true;
        } else {
            return false;
        }
    };
    
    /**
     * 获取首选的AI接口（优先使用增强版）
     */
    FormulaGenerator.prototype.getPreferredAIInterface = function() {
        // 1. 优先使用增强AI接口
        if (window.enhancedAIInterface) {
            return window.enhancedAIInterface;
        }
        
        // 2. 回退到标准AI接口
        if (window.aiInterface) {
            return window.aiInterface;
        }
        
        return null;
    };
    
    /**
     * 获取标准AI接口（遵循AIapi.txt规范）
     */
    FormulaGenerator.prototype.getStandardApi = function() {
        // 1. 优先使用aiapiStandard（严格按AIapi.txt规范）
        if (window.aiapiStandard && window.CURRENT_AI_CONFIG) {
            // 使用当前配置创建新的AI API实例
            try {
                var api = new window.aiapiStandard(window.CURRENT_AI_CONFIG);
                return api;
            } catch (error) {
                console.warn('标准AI接口初始化失败:', error.message);
            }
        }
        
        // 2. 回退到传统AI接口
        if (window.aiInterface) {
            return window.aiInterface;
        }
        
        return null;
    };
    
    FormulaGenerator.prototype.init = function() {
        this.bindEvents();
        this.updateCurrentCell();
        
        if (this.isExcelEnvironment) {
            this.loadWorkbookData();
        } else {
            this.loadMockData();
        }
    };
    
    /**
     * 加载模拟数据（Web环境）
     */
    FormulaGenerator.prototype.loadMockData = function() {
        console.log('📋 [loadMockData] 加载模拟数据');
        
        var mockWorkbookData = [
            {
                name: '示例工作簿.xlsx',
                path: '',
                worksheets: [
                    {
                        name: 'Sheet1',
                        usedRange: { rows: 100, columns: 5 },
                        headers: ['日期', '产品', '数量', '单价', '总计'],
                        sampleData: [
                            ['2024-01-01', '产品A', 10, 100, 1000],
                            ['2024-01-02', '产品B', 20, 200, 4000]
                        ],
                        dataStructure: {
                            type: 'medium',
                            description: '中等数据量（≤100行）',
                            rowCount: 100,
                            colCount: 5,
                            dataDensity: 0.95,
                            hasHeaders: true
                        }
                    },
                    {
                        name: 'Sheet2',
                        usedRange: { rows: 50, columns: 3 },
                        headers: ['姓名', '部门', '薪资'],
                        sampleData: [
                            ['张三', '销售部', 5000],
                            ['李四', '技术部', 6000]
                        ],
                        dataStructure: {
                            type: 'small',
                            description: '小数据量（≤10行）',
                            rowCount: 50,
                            colCount: 3,
                            dataDensity: 0.98,
                            hasHeaders: true
                        }
                    }
                ]
            }
        ];
        
        this.updateWorkbookList(mockWorkbookData);
        this.updateStatus('Web环境模式：使用示例数据');
        console.log('📊 [loadMockData] 模拟数据加载完成:', mockWorkbookData);
    };
    
    FormulaGenerator.prototype.bindEvents = function() {
        var self = this;
        
        // 引用类型切换
        var referenceTypeRadios = document.querySelectorAll('input[name="referenceType"]');
        for (var i = 0; i < referenceTypeRadios.length; i++) {
            referenceTypeRadios[i].addEventListener('change', function(e) {
                self.referenceType = e.target.value;
                self.toggleReferenceSelection();
                // 当引用类型改变时，触发相应操作
                if (e.target.value === 'worksheet') {
                    // 跨工作表选择 - 显示工作表选择区域
                    self.showWorksheetSelection();
                } else if (e.target.value === 'workbook') {
                    // 跨工作簿选择 - 显示工作簿选择对话框
                    self.showWorkbookSelection();
                } else {
                    // 隐藏工作表选择区域
                    var worksheetSelection = document.getElementById('worksheetSelection');
                    if (worksheetSelection) {
                        worksheetSelection.style.display = 'none';
                    }
                }
            });
        }
        
        // 添加对跨工作簿选项的点击事件监听，支持重新选择
        var workbookRadio = document.querySelector('input[name="referenceType"][value="workbook"]');
        if (workbookRadio) {
            workbookRadio.addEventListener('click', function(e) {
                // 如果已经是选中状态，再次点击则重新打开选择对话框
                if (this.checked) {
                    self.showWorkbookSelection();
                }
            });
        }
        
        // 添加对跨工作表选项的点击事件监听，支持重新选择
        var worksheetRadio = document.querySelector('input[name="referenceType"][value="worksheet"]');
        if (worksheetRadio) {
            worksheetRadio.addEventListener('click', function(e) {
                // 如果已经是选中状态，再次点击则重新打开选择对话框
                if (this.checked) {
                    self.showWorksheetSelection();
                }
            });
        }
        
        // 填充方向设置
        var fillRightElement = document.getElementById('fillRight');
        if (fillRightElement) {
            fillRightElement.addEventListener('change', function(e) {
                self.fillRight = e.target.checked;
            });
        }
        
        var fillDownElement = document.getElementById('fillDown');
        if (fillDownElement) {
            fillDownElement.addEventListener('change', function(e) {
                self.fillDown = e.target.checked;
            });
        }
        
        // 公式描述输入
        var formulaDescriptionElement = document.getElementById('formulaDescription');
        if (formulaDescriptionElement) {
            formulaDescriptionElement.addEventListener('input', function(e) {
                self.formulaDescription = e.target.value;
            });
        }
        
        // 生成按钮
        var generateFormulaElement = document.getElementById('generateFormula');
        if (generateFormulaElement) {
            generateFormulaElement.addEventListener('click', function() {
                self.generateFormula();
            });
        }
        
        // 刷新工作簿按钮
        var refreshWorkbooksElement = document.getElementById('refreshWorkbooks');
        if (refreshWorkbooksElement) {
            refreshWorkbooksElement.addEventListener('click', function() {
                if (self.isExcelEnvironment) {
                    self.loadWorkbookData();
                } else {
                    self.loadMockData();
                }
            });
        }
        
        // 应用公式按钮
        var applyFormulaElement = document.getElementById('applyFormula');
        if (applyFormulaElement) {
            applyFormulaElement.addEventListener('click', function() {
                self.applySelectedFormula();
            });
        }
        
        // 清空按钮
        var clearAllElement = document.getElementById('clearAll');
        if (clearAllElement) {
            clearAllElement.addEventListener('click', function() {
                self.clearAll();
            });
        }
        
        // 工作簿搜索
        var searchInputElement = document.getElementById('workbookSearch');
        if (searchInputElement) {
            searchInputElement.addEventListener('input', function(e) {
                self.filterWorkbooks(e.target.value);
            });
        }
    };
    
    /**
     * 显示工作表选择区域（跨工作表模式）
     */
    FormulaGenerator.prototype.showWorksheetSelection = function() {
        try {
            // 显示工作表选择区域
            var worksheetSelection = document.getElementById('worksheetSelection');
            if (worksheetSelection) {
                worksheetSelection.style.display = 'block';
            }
            
            // 加载当前工作簿的工作表列表
            this.loadCurrentWorkbookWorksheets();
        } catch (error) {
            console.error('显示工作表选择区域失败:', error);
        }
    };
    
    /**
     * 显示工作簿选择对话框（跨工作簿模式）
     */
    FormulaGenerator.prototype.showWorkbookSelection = function() {
        try {
            // 触发全局事件来显示工作簿选择对话框
            if (window.aiHelperMainInstance && typeof window.aiHelperMainInstance.showWorkbookSelector === 'function') {
                window.aiHelperMainInstance.showWorkbookSelector();
            } else {
                // 备用方案：直接显示工作簿模态框
                var workbookModal = document.getElementById('workbookModal');
                if (workbookModal) {
                    workbookModal.style.display = 'flex';
                }
            }
        } catch (error) {
            console.error('显示工作簿选择对话框失败:', error);
        }
    };
    
    /**
     * 加载当前工作簿的工作表列表（跨工作表模式）
     */
    FormulaGenerator.prototype.loadCurrentWorkbookWorksheets = function() {
        try {
            if (!this.isExcelEnvironment || !window.Application) {
                return;
            }
            
            var workbookList = document.getElementById('workbookList');
            if (!workbookList) return;
            
            // 清空之前的内容
            workbookList.innerHTML = '';
            
            // 获取当前工作簿
            var activeWorkbook = window.Application.ActiveWorkbook;
            if (!activeWorkbook) return;
            
            var html = '';
            var workbookName = activeWorkbook.Name;
            
            html += '<div class="workbook-item" data-workbook="' + workbookName + '">' +
                    '<div class="workbook-header">' +
                    '<h4>' + workbookName + '</h4>' +
                    '</div>' +
                    '<div class="worksheet-list">';
            
            // 获取当前工作簿的所有工作表
            if (activeWorkbook.Worksheets) {
                for (var i = 1; i <= activeWorkbook.Worksheets.Count; i++) {
                    var worksheet = activeWorkbook.Worksheets.Item(i);
                    var worksheetName = worksheet.Name;
                    
                    html += '<div class="worksheet-item">' +
                            '<input type="checkbox" id="ws_' + workbookName + '_' + worksheetName + '" ' +
                            'name="worksheets" value="' + worksheetName + '" data-workbook="' + workbookName + '">' +
                            '<label for="ws_' + workbookName + '_' + worksheetName + '">' +
                            worksheetName +
                            '</label>' +
                            '</div>';
                }
            }
            
            html += '</div></div>';
            
            workbookList.innerHTML = html;
            
            // 绑定工作表选择事件
            var worksheetCheckboxes = document.querySelectorAll('input[name="worksheets"]');
            var self = this;
            for (var j = 0; j < worksheetCheckboxes.length; j++) {
                worksheetCheckboxes[j].addEventListener('change', function(e) {
                    self.handleWorksheetSelection(e.target);
                });
            }
            
        } catch (error) {
            console.error('加载当前工作簿工作表列表失败:', error);
        }
    };
    
    FormulaGenerator.prototype.updateCurrentCell = function() {
        try {
            // 获取当前选中的单元格信息
            if (this.isExcelEnvironment && window.Application && window.Application.ActiveSheet) {
                var activeSheet = window.Application.ActiveSheet;
                var selection = window.Application.Selection;
                
                if (selection) {
                    this.currentCell = {
                        workbook: window.Application.ActiveWorkbook ? window.Application.ActiveWorkbook.Name : '',
                        worksheet: activeSheet ? activeSheet.Name : '',
                        row: selection.Row || 1,
                        col: selection.Column || 1,
                        cellAddress: this.getCellAddress(selection.Row || 1, selection.Column || 1)
                    };
                    
                    this.updateCurrentCellDisplay();
                    return;
                }
            }
            
            // Web环境或无法获取Excel信息时的默认值
            this.currentCell = {
                workbook: '未知',
                worksheet: '当前工作表',
                row: 1,
                col: 1,
                cellAddress: 'A1'
            };
            
            this.updateCurrentCellDisplay();
            
        } catch (error) {
            console.error('获取当前单元格信息失败:', error);
            this.currentCell = {
                workbook: '未知',
                worksheet: '当前工作表',
                row: 1,
                col: 1,
                cellAddress: 'A1'
            };
            this.updateCurrentCellDisplay();
        }
    };
    
    /**
     * 更新当前单元格显示
     */
    FormulaGenerator.prototype.updateCurrentCellDisplay = function() {
        var currentCellElement = document.getElementById('currentCell');
        if (currentCellElement) {
            currentCellElement.textContent = 
                this.currentCell.cellAddress + ' (' + this.currentCell.worksheet + ')';
        }
    };
    
    FormulaGenerator.prototype.getCellAddress = function(row, col) {
        var columnLetters = this.getColumnLetter(col);
        return columnLetters + row;
    };
    
    FormulaGenerator.prototype.getColumnLetter = function(col) {
        var temp = '';
        var columnNumber = col;
        
        while (columnNumber > 0) {
            var remainder = (columnNumber - 1) % 26;
            temp = String.fromCharCode(65 + remainder) + temp;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        
        return temp;
    };
    
    FormulaGenerator.prototype.loadWorkbookData = function() {
        if (!this.isExcelEnvironment) {
            this.loadMockData();
            return;
        }
        
        try {
            this.updateStatus('正在加载工作簿...');
            
            // 检查Excel对象
            if (!window.Application || !window.Application.Workbooks) {
                throw new Error('Excel应用程序不可用');
            }
            
            var workbooks = window.Application.Workbooks;
            var workbookData = [];
            
            for (var i = 1; i <= workbooks.Count; i++) {
                var wb = workbooks.Item(i);
                var worksheets = [];
                
                if (wb.Worksheets) {
                    for (var j = 1; j <= wb.Worksheets.Count; j++) {
                        var ws = wb.Worksheets.Item(j);
                        worksheets.push({
                            name: ws.Name,
                            usedRange: this.getUsedRangeInfo(ws),
                            headers: this.extractWorksheetHeaders(ws),
                            sampleData: this.extractSampleData(ws),
                            dataStructure: this.analyzeDataStructure(ws)
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
            
        } catch (error) {
            console.error('加载Excel工作簿数据失败:', error);
            this.updateStatus('加载工作簿失败，使用模拟数据');
            this.loadMockData();
            this.showNotification('Excel环境不可用，已切换到演示模式', 'warning');
        }
    };
    
    /**
     * 获取工作表使用范围信息（仅Excel环境）
     */
    FormulaGenerator.prototype.getUsedRangeInfo = function(worksheet) {
        try {
            if (!this.isExcelEnvironment) {
                return { rows: 0, columns: 0 };
            }
            
            var usedRange = worksheet.UsedRange;
            if (usedRange) {
                return {
                    rows: usedRange.Rows.Count,
                    columns: usedRange.Columns.Count
                };
            }
        } catch (error) {
            console.warn('获取使用范围信息失败:', error);
        }
        return { rows: 0, columns: 0 };
    };
    
    /**
     * 提取工作表表头（仅Excel环境）
     */
    FormulaGenerator.prototype.extractWorksheetHeaders = function(worksheet) {
        try {
            if (!this.isExcelEnvironment) {
                return [];
            }
            
            var usedRange = worksheet.UsedRange;
            if (!usedRange) return [];
            
            var firstRow = usedRange.Row;
            var maxCol = usedRange.Column + usedRange.Columns.Count - 1;
            var headers = [];
            
            // 按照WPS JSA规范，使用Cells.Item方式访问单元格
            for (var col = usedRange.Column; col <= maxCol; col++) {
                try {
                    var cell = worksheet.Cells.Item(firstRow, col);
                    // WPS JSA环境中的单元格值获取
                    var value = '';
                    if (cell) {
                        // 尝试多种方式获取单元格值
                        if (typeof cell.Value === 'function') {
                            value = String(cell.Value()).trim();
                        } else if (cell.Value !== null && cell.Value !== undefined) {
                            value = String(cell.Value).trim();
                        } else {
                            value = '列' + this.getColumnLetter(col);
                        }
                    }
                    headers.push(value || '列' + this.getColumnLetter(col));
                } catch (cellError) {
                    headers.push('列' + this.getColumnLetter(col));
                }
            }
            
            return headers;
        } catch (error) {
            return [];
        }
    };
    
    /**
     * 提取示例数据（仅Excel环境）
     */
    FormulaGenerator.prototype.extractSampleData = function(worksheet) {
        try {
            if (!this.isExcelEnvironment) {
                return [];
            }
            
            var samples = [];
            var usedRange = worksheet.UsedRange;
            
            if (!usedRange || usedRange.Rows.Count < 2) return samples;
            
            var maxRows = Math.min(usedRange.Rows.Count - 1, 5);
            var maxCol = usedRange.Column + usedRange.Columns.Count - 1;
            
            for (var row = usedRange.Row + 1; row <= usedRange.Row + maxRows; row++) {
                var rowData = [];
                for (var col = usedRange.Column; col <= maxCol; col++) {
                    try {
                        // 按照WPS JSA规范，使用Cells.Item方式访问单元格
                        var cell = worksheet.Cells.Item(row, col);
                        // WPS JSA环境中的单元格值获取
                        var value = '';
                        if (cell) {
                            // 尝试多种方式获取单元格值
                            if (typeof cell.Value === 'function') {
                                value = String(cell.Value());
                            } else if (cell.Value !== null && cell.Value !== undefined) {
                                value = String(cell.Value);
                            }
                        }
                        rowData.push(value);
                    } catch (cellError) {
                        rowData.push('');
                    }
                }
                samples.push(rowData);
            }
            
            return samples;
        } catch (error) {
            return [];
        }
    };
    
    /**
     * 分析数据结构（仅Excel环境）
     */
    FormulaGenerator.prototype.analyzeDataStructure = function(worksheet) {
        try {
            if (!this.isExcelEnvironment) {
                return { type: 'web', description: 'Web环境' };
            }
            
            var usedRange = worksheet.UsedRange;
            if (!usedRange) return { type: 'empty', description: '空工作表' };
            
            var rowCount = usedRange.Rows.Count;
            var colCount = usedRange.Columns.Count;
            
            var dataType = 'unknown';
            var description = '';
            
            if (rowCount === 0) {
                dataType = 'empty';
                description = '空工作表';
            } else if (rowCount <= 10) {
                dataType = 'small';
                description = '小数据量（≤10行）';
            } else if (rowCount <= 100) {
                dataType = 'medium';
                description = '中等数据量（≤100行）';
            } else {
                dataType = 'large';
                description = '大数据量（>100行）';
            }
            
            return {
                type: dataType,
                description: description,
                rowCount: rowCount,
                colCount: colCount,
                hasHeaders: this.hasHeaderRow(worksheet)
            };
        } catch (error) {
            console.warn('分析数据结构失败:', error);
            return { type: 'error', description: '分析失败' };
        }
    };
    
    /**
     * 检查是否有表头行（仅Excel环境）
     */
    FormulaGenerator.prototype.hasHeaderRow = function(worksheet) {
        try {
            if (!this.isExcelEnvironment) return true;
            
            var usedRange = worksheet.UsedRange;
            if (!usedRange || usedRange.Rows.Count < 1) return false;
            
            var firstRow = usedRange.Row;
            var maxCol = usedRange.Column + usedRange.Columns.Count - 1;
            
            var nonEmptyCount = 0;
            for (var col = usedRange.Column; col <= maxCol; col++) {
                try {
                    // 按照WPS JSA规范，使用Cells.Item方式访问单元格
                    var cell = worksheet.Cells.Item(firstRow, col);
                    // WPS JSA环境中的单元格值获取
                    var cellValue = '';
                    if (cell) {
                        if (typeof cell.Value === 'function') {
                            cellValue = String(cell.Value()).trim();
                        } else if (cell.Value !== null && cell.Value !== undefined) {
                            cellValue = String(cell.Value).trim();
                        }
                    }
                    if (cellValue !== '') {
                        nonEmptyCount++;
                    }
                } catch (e) {
                    // 忽略单元格访问错误
                }
            }
            
            return nonEmptyCount > 0;
        } catch (error) {
            console.warn('检查表头行失败:', error);
            return true;
        }
    };
    
    /**
     * 更新工作簿列表显示
     */
    FormulaGenerator.prototype.updateWorkbookList = function(workbookData) {
        var workbookList = document.getElementById('workbookList');
        if (!workbookList) return;
        
        var html = '';
        for (var i = 0; i < workbookData.length; i++) {
            var workbook = workbookData[i];
            var worksheetCount = workbook.worksheets ? workbook.worksheets.length : 0;
            html += '<div class="workbook-item" data-workbook="' + workbook.name + '">' +
                    '<div class="workbook-header">' +
                    '<h4>' + workbook.name + '</h4>' +
                    '<span class="workbook-info">' + worksheetCount + '个工作表</span>' +
                    '</div>' +
                    '<div class="worksheet-list">';
            
            if (workbook.worksheets) {
                for (var j = 0; j < workbook.worksheets.length; j++) {
                    var ws = workbook.worksheets[j];
                    html += '<div class="worksheet-item">' +
                            '<input type="checkbox" id="ws_' + workbook.name + '_' + ws.name + '" ' +
                            'name="worksheets" value="' + ws.name + '" data-workbook="' + workbook.name + '">' +
                            '<label for="ws_' + workbook.name + '_' + ws.name + '">' +
                            ws.name + ' ' +
                            '<small>(' + ws.usedRange.rows + '行 x ' + ws.usedRange.columns + '列)</small>' +
                            '</label>' +
                            '</div>';
                }
            } else {
                html += '<div class="no-worksheets">无工作表</div>';
            }
            
            html += '</div></div>';
        }
        
        workbookList.innerHTML = html;
        
        // 重新绑定工作表选择事件
        var worksheetCheckboxes = document.querySelectorAll('input[name="worksheets"]');
        var self = this;
        for (var k = 0; k < worksheetCheckboxes.length; k++) {
            worksheetCheckboxes[k].addEventListener('change', function(e) {
                self.handleWorksheetSelection(e.target);
            });
        }
    };
    
    /**
     * 处理工作表选择
     */
    FormulaGenerator.prototype.handleWorksheetSelection = function(checkbox) {
        var workbookName = checkbox.dataset.workbook; // 使用dataset获取工作簿名称
        var worksheetName = checkbox.value;
        
        if (checkbox.checked) {
            // 添加到选中的工作表列表
            var found = false;
            for (var i = 0; i < this.selectedWorksheets.length; i++) {
                if (this.selectedWorksheets[i].workbook === workbookName && 
                    this.selectedWorksheets[i].worksheet === worksheetName) {
                    found = true;
                    break;
                }
            }
            if (!found) {
                this.selectedWorksheets.push({ workbook: workbookName, worksheet: worksheetName });
            }
        } else {
            // 从选中列表中移除
            var newSelectedWorksheets = [];
            for (var j = 0; j < this.selectedWorksheets.length; j++) {
                if (!(this.selectedWorksheets[j].workbook === workbookName && 
                      this.selectedWorksheets[j].worksheet === worksheetName)) {
                    newSelectedWorksheets.push(this.selectedWorksheets[j]);
                }
            }
            this.selectedWorksheets = newSelectedWorksheets;
        }
        
        console.log('📋 [handleWorksheetSelection] 当前选中的工作表:', this.selectedWorksheets);
    };
    
    /**
     * 切换引用选择区域
     */
    FormulaGenerator.prototype.toggleReferenceSelection = function() {
        var referenceSection = document.getElementById('referenceSection');
        if (referenceSection) {
            referenceSection.style.display = this.referenceType === 'other' ? 'block' : 'none';
        }
        
        // 根据引用类型显示相应选择区域
        var worksheetSelection = document.getElementById('worksheetSelection');
        if (worksheetSelection) {
            worksheetSelection.style.display = (this.referenceType === 'worksheet' || this.referenceType === 'workbook') ? 'block' : 'none';
        }
    };
    
    /**
     * 过滤工作簿
     */
    FormulaGenerator.prototype.filterWorkbooks = function(searchTerm) {
        var items = document.querySelectorAll('.workbook-item');
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var text = item.textContent.toLowerCase();
            item.style.display = text.includes(searchTerm.toLowerCase()) ? 'block' : 'none';
        }
    };
    
    /**
     * 更新状态
     */
    FormulaGenerator.prototype.updateStatus = function(status) {
        var statusElement = document.getElementById('status');
        if (statusElement) {
            statusElement.textContent = status;
        }
    };
    
    /**
     * 显示通知
     */
    FormulaGenerator.prototype.showNotification = function(message, type) {
        type = type || 'info';
        // console.log('📢 [通知] ' + type + ': ' + message);
        
        // 简单的通知实现
        var notification = document.createElement('div');
        notification.className = 'notification notification-' + type;
        notification.textContent = message;
        notification.style.cssText = 
            'position: fixed; ' +
            'top: 10px; ' +
            'right: 10px; ' +
            'padding: 10px 15px; ' +
            'background: ' + (type === 'error' ? '#f44336' : type === 'warning' ? '#ff9800' : '#4caf50') + '; ' +
            'color: white; ' +
            'border-radius: 4px; ' +
            'z-index: 9999; ' +
            'font-size: 14px; ' +
            'max-width: 300px;';
        
        document.body.appendChild(notification);
        
        // 3秒后自动移除
        setTimeout(function() {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 3000);
    };
    
    /**
     * 生成公式
     */
    FormulaGenerator.prototype.generateFormula = function() {
        if (this.isGenerating) {
            console.warn('⚠️ 公式正在生成中，请等待');
            return;
        }
        
        this.isGenerating = true;
        this.updateStatus('正在生成公式...');
        
        var self = this;
        setTimeout(function() {  // 使用setTimeout模拟异步操作
            try {
                // 验证输入
                // 允许空描述，系统会进行智能分析
                if (!self.formulaDescription.trim()) {
                    console.log('ℹ️ [generateFormula] 空描述，将进行智能分析');
                }
                
                // 准备上下文信息
                var contextInfo = {
                    currentCell: self.currentCell,
                    selectedWorksheets: self.selectedWorksheets,
                    referenceType: self.referenceType,
                    workbookData: self.getSelectedWorkbookData()
                };
                
                console.log('🔧 [generateFormula] 开始生成公式，描述:', self.formulaDescription);
                console.log('📋 [generateFormula] 上下文信息:', contextInfo);
                
                // 优先使用标准API进行公式生成
                if (self.standardApi && self.standardApi.chat) {
                    console.log('🤖 [generateFormula] 使用标准API生成公式');
                    
                    var prompt = self.buildPrompt(self.formulaDescription, {
                        currentCell: self.currentCell,
                        selectedWorksheets: self.selectedWorksheets,
                        referenceType: self.referenceType
                    });
                    
                    // 使用标准API调用 - 严格遵循AIapi.txt格式
                    self.standardApi.chat(prompt)
                        .then(function(response) {
                            if (response && response.choices && response.choices[0] && response.choices[0].message) {
                                var formula = response.choices[0].message.content;
                                console.log('✅ [generateFormula] AI生成结果:', formula);
                                
                                // 显示生成的公式
                                self.displayGeneratedFormula(formula);
                                self.updateStatus('公式生成完成');
                                self.showNotification('公式生成成功！', 'success');
                            } else {
                                throw new Error('AI接口返回格式异常');
                            }
                        })
                        .catch(function(error) {
                            throw error;
                        });
                } else if (self.aiInterface && self.aiInterface.chat) {
                    // 回退到增强AI接口
                    console.log('🤖 [generateFormula] 使用增强AI接口生成公式');
                    
                    var prompt = self.buildPrompt(self.formulaDescription, {
                        currentCell: self.currentCell,
                        selectedWorksheets: self.selectedWorksheets,
                        referenceType: self.referenceType
                    });
                    
                    self.aiInterface.chat({
                        messages: [
                            { role: 'system', content: '你是一个Excel公式专家。' },
                            { role: 'user', content: prompt }
                        ]
                    })
                    .then(function(response) {
                        if (response && response.choices && response.choices[0] && response.choices[0].message) {
                            var formula = response.choices[0].message.content;
                            console.log('✅ [generateFormula] AI生成结果:', formula);
                            
                            // 显示生成的公式
                            self.displayGeneratedFormula(formula);
                            self.updateStatus('公式生成完成');
                            self.showNotification('公式生成成功！', 'success');
                        } else {
                            throw new Error('AI接口返回格式异常');
                        }
                    })
                    .catch(function(error) {
                        throw error;
                    });
                } else {
                    // 模拟公式生成（演示模式）
                    // console.log('🎭 [generateFormula] 使用模拟公式生成');
                    
                    var mockFormulas = [
                        '=IF(' + self.currentCell.cellAddress + '<>"",' + self.currentCell.cellAddress + ',"无数据")',
                        '=SUMIF(Sheet1!A:A,"条件",Sheet1!C:C)',
                        '=VLOOKUP(' + self.currentCell.cellAddress + ',Sheet1!A:E,5,FALSE)',
                        '=IFERROR(' + self.currentCell.cellAddress + '/100,"错误")',
                        '=COUNTIF(Sheet1!B:B,">0")'
                    ];
                    
                    var randomFormula = mockFormulas[Math.floor(Math.random() * mockFormulas.length)];
                    self.displayGeneratedFormula(randomFormula);
                    self.updateStatus('模拟公式生成完成');
                    // self.showNotification('演示模式：生成模拟公式', 'info');
                }
                
            } catch (error) {
                console.error('生成公式失败:', error);
                self.updateStatus('生成公式失败');
                self.showNotification('生成公式失败: ' + error.message, 'error');
            } finally {
                self.isGenerating = false;
            }
        }, 100); // 延迟执行以模拟异步操作
    };
    
    /**
     * 构建提示词
     */
    FormulaGenerator.prototype.buildPrompt = function(description, contextInfo) {
        var context = 
            '当前单元格: ' + contextInfo.currentCell.cellAddress + ' (工作表: ' + contextInfo.currentCell.worksheet + ')\n' +
            '引用类型: ' + contextInfo.referenceType + '\n' +
            '选中的工作表数量: ' + contextInfo.selectedWorksheets.length + '\n';
        
        var prompt = 
            '请为以下场景生成一个Excel公式：\n\n' +
            '需求描述: ' + description + '\n\n' +
            context + '\n' +
            '请提供：\n' +
            '1. 公式本身\n' +
            '2. 公式说明\n' +
            '3. 使用注意事项\n\n' +
            '注意：返回格式应该清晰易懂。\n';
        
        return prompt;
    };
    
    /**
     * 显示生成的公式
     */
    FormulaGenerator.prototype.displayGeneratedFormula = function(formula) {
        var formulaPreview = document.getElementById('formulaPreview');
        if (formulaPreview) {
            formulaPreview.textContent = formula;
            formulaPreview.style.display = 'block';
        }
        
        this.selectedFormula = formula;
        
        // 启用应用按钮
        var applyBtn = document.getElementById('applyFormula');
        if (applyBtn) {
            applyBtn.disabled = false;
        }
    };
    
    /**
     * 应用选中的公式
     */
    FormulaGenerator.prototype.applySelectedFormula = function() {
        if (!this.selectedFormula) {
            this.showNotification('请先生成公式', 'warning');
            return;
        }
        
        try {
            if (this.isExcelEnvironment && window.Application) {
                // 在Excel环境中应用公式
                var selection = window.Application.Selection;
                if (selection) {
                    selection.Formula = this.selectedFormula;
                    this.showNotification('公式已应用到选中的单元格', 'success');
                } else {
                    throw new Error('未选中任何单元格');
                }
            } else {
                // Web环境模拟
                console.log('🎭 [applySelectedFormula] 演示模式：模拟应用公式', this.selectedFormula);
                this.showNotification('演示模式：公式 "' + this.selectedFormula + '" 已应用', 'info');
            }
            
        } catch (error) {
            console.error('应用公式失败:', error);
            this.showNotification('应用公式失败: ' + error.message, 'error');
        }
    };
    
    /**
     * 清空所有内容
     */
    FormulaGenerator.prototype.clearAll = function() {
        // 清空公式描述
        var descriptionInput = document.getElementById('formulaDescription');
        if (descriptionInput) {
            descriptionInput.value = '';
        }
        
        // 清空公式预览
        var formulaPreview = document.getElementById('formulaPreview');
        if (formulaPreview) {
            formulaPreview.textContent = '';
            formulaPreview.style.display = 'none';
        }
        
        // 取消所有工作表选择
        var worksheetCheckboxes = document.querySelectorAll('input[name="worksheets"]');
        for (var i = 0; i < worksheetCheckboxes.length; i++) {
            worksheetCheckboxes[i].checked = false;
        }
        
        // 重置选项
        var fillRightElement = document.getElementById('fillRight');
        if (fillRightElement) {
            fillRightElement.checked = false;
        }
        var fillDownElement = document.getElementById('fillDown');
        if (fillDownElement) {
            fillDownElement.checked = false;
        }
        
        // 重置状态
        this.selectedWorksheets = [];
        this.selectedFormula = null;
        this.formulaDescription = '';
        this.fillRight = false;
        this.fillDown = false;
        
        // 禁用应用按钮
        var applyBtn = document.getElementById('applyFormula');
        if (applyBtn) {
            applyBtn.disabled = true;
        }
        
        this.updateStatus('已清空所有内容');
        console.log('🧹 [clearAll] 所有内容已清空');
        
        // 更新已选择数据源显示
        if (window.aiHelperMainInstance && typeof window.aiHelperMainInstance.updateSelectedSourcesDisplay === 'function') {
            window.aiHelperMainInstance.updateSelectedSourcesDisplay();
        }
    };
    
    /**
     * 获取选中的工作簿数据
     */
    FormulaGenerator.prototype.getSelectedWorkbookData = function() {
        try {
            // 如果有明确选择的工作表，则返回这些工作表的信息
            if (this.selectedWorksheets.length > 0) {
                return this.selectedWorksheets.map(item => ({
                    workbook: item.workbook,
                    worksheet: item.worksheet
                }));
            }
            
            // 根据引用类型返回相应的默认数据
            switch (this.referenceType) {
                case 'current':
                    // 当前工作表模式，返回当前工作表
                    return [{
                        workbook: this.currentCell.workbook,
                        worksheet: this.currentCell.worksheet
                    }];
                    
                case 'worksheet':
                    // 跨工作表模式，如果没有明确选择，则返回当前工作簿的所有工作表
                    if (this.isExcelEnvironment && window.Application && window.Application.ActiveWorkbook) {
                        var activeWorkbook = window.Application.ActiveWorkbook;
                        var worksheets = [];
                        if (activeWorkbook.Worksheets) {
                            for (var i = 1; i <= activeWorkbook.Worksheets.Count; i++) {
                                var ws = activeWorkbook.Worksheets.Item(i);
                                worksheets.push({
                                    workbook: activeWorkbook.Name,
                                    worksheet: ws.Name
                                });
                            }
                        }
                        return worksheets;
                    }
                    break;
                    
                case 'workbook':
                    // 跨工作簿模式，返回选中的工作簿中的工作表
                    if (this.selectedWorkbooks && this.selectedWorkbooks.length > 0) {
                        var allWorksheets = [];
                        this.selectedWorkbooks.forEach(workbook => {
                            if (workbook.worksheets) {
                                workbook.worksheets.forEach(worksheet => {
                                    allWorksheets.push({
                                        workbook: workbook.name || workbook.workBookName,
                                        worksheet: worksheet.name || worksheet.workSheetName
                                    });
                                });
                            }
                        });
                        return allWorksheets;
                    }
                    break;
            }
            
            // 默认返回当前工作表
            return [{
                workbook: this.currentCell.workbook,
                worksheet: this.currentCell.worksheet
            }];
        } catch (error) {
            console.error('获取选中的工作簿数据失败:', error);
            // 出错时返回当前工作表
            return [{
                workbook: this.currentCell.workbook,
                worksheet: this.currentCell.worksheet
            }];
        }
    };
    
    return FormulaGenerator;
})();

// 页面加载完成后初始化
if (typeof document !== 'undefined') {
    var initFormulaGenerator = function() {
        try {
            window.formulaGenerator = new FormulaGenerator();
        } catch (error) {
            console.error('❌ [FormulaGenerator] 初始化失败:', error);
        }
    };
    
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initFormulaGenerator);
    } else {
        initFormulaGenerator();
    }
}