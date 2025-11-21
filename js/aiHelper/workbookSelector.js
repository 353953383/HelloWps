/**
 * 工作簿选择器模块
 * 负责处理工作簿和跨工作簿选择逻辑
 */

class WorkbookSelector {
    constructor(formulaGenerator) {
        this.formulaGenerator = formulaGenerator;
        this.allWorkbooks = [];
        this.selectedWorkbooks = [];
        this.init();
    }
    
    init() {
        this.bindEvents();
        this.initSearchBox();
    }
    
    bindEvents() {
        // 模态框事件
        const modal = document.getElementById('workbookModal');
        
        // 点击模态框背景关闭
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.closeModal();
            }
        });
        
        // ESC键关闭模态框
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this.closeModal();
            }
        });
    }
    
    /**
     * 获取所有工作簿信息
     */
    getAllWorkbooks() {
        try {
            const workbooks = [];
            
            if (window.Application && window.Application.Workbooks) {
                for (let i = 1; i <= window.Application.Workbooks.Count; i++) {
                    const wb = window.Application.Workbooks.Item(i);
                    const workbookInfo = {
                        name: wb.Name,
                        path: wb.Path || '',
                        isActive: window.Application.ActiveWorkbook === wb,
                        worksheets: this.getWorkbookWorksheets(wb)
                    };
                    workbooks.push(workbookInfo);
                }
            }
            
            this.allWorkbooks = workbooks;
            return workbooks;
            
        } catch (error) {
            console.error('获取工作簿信息失败:', error);
            throw error;
        }
    }
    
    /**
     * 获取工作簿中的所有工作表
     */
    getWorkbookWorksheets(workbook) {
        try {
            const worksheets = [];
            
            if (workbook.Worksheets) {
                for (let j = 1; j <= workbook.Worksheets.Count; j++) {
                    const ws = workbook.Worksheets.Item(j);
                    const worksheetInfo = {
                        name: ws.Name,
                        usedRange: this.getWorksheetUsedRange(ws),
                        headers: this.extractWorksheetHeaders(ws)
                    };
                    worksheets.push(worksheetInfo);
                }
            }
            
            return worksheets;
            
        } catch (error) {
            console.warn(`获取工作簿 ${workbook.Name} 的工作表失败:`, error);
            return [];
        }
    }
    
    /**
     * 获取工作表的使用范围
     */
    getWorksheetUsedRange(worksheet) {
        try {
            const usedRange = worksheet.UsedRange;
            if (usedRange) {
                return {
                    rows: usedRange.Rows.Count,
                    columns: usedRange.Columns.Count,
                    startRow: usedRange.Row,
                    startCol: usedRange.Column,
                    endRow: usedRange.Row + usedRange.Rows.Count - 1,
                    endCol: usedRange.Column + usedRange.Columns.Count - 1,
                    address: usedRange.Address
                };
            }
        } catch (error) {
            console.warn('获取工作表使用范围失败:', error);
        }
        
        return {
            rows: 0,
            columns: 0,
            startRow: 1,
            startCol: 1,
            endRow: 1,
            endCol: 1,
            address: 'A1'
        };
    }
    
    /**
     * 提取工作表的表头信息
     */
    extractWorksheetHeaders(worksheet) {
        try {
            const usedRange = worksheet.UsedRange;
            if (usedRange && usedRange.Rows.Count > 0) {
                const headerRow = usedRange.Rows.Item(1);
                const headers = [];
                
                for (let col = 1; col <= headerRow.Columns.Count; col++) {
                    // 获取列字母并构建单元格地址
                    const columnLetter = this.getColumnLetter(col);
                    const cellAddress = columnLetter + '1';
                    const cell = headerRow.Range(cellAddress);
                    
                    // 优先使用 Value2 属性获取值
                    let cellValue = '';
                    if (cell && cell.Value2 !== null && cell.Value2 !== undefined) {
                        const value = cell.Value2;
                        if (typeof value === 'string') {
                            cellValue = value.trim();
                        } else if (typeof value === 'number' && !isNaN(value)) {
                            cellValue = value.toString();
                        } else if (value instanceof Date) {
                            cellValue = value.toLocaleDateString();
                        } else if (typeof value === 'boolean') {
                            cellValue = value ? 'TRUE' : 'FALSE';
                        } else if (typeof value === 'object' && value !== null) {
                            if (value.Text && typeof value.Text === 'string') {
                                cellValue = value.Text.trim();
                            }
                        }
                    }
                    // 如果 Value2 为空，尝试使用 Text 属性
                    else if (cell && cell.Text && typeof cell.Text === 'string') {
                        cellValue = cell.Text.trim();
                    }
                    
                    headers.push({
                        column: columnLetter,
                        columnIndex: col,
                        value: cellValue || `列${col}`,
                        dataType: this.detectDataType(cell)
                    });
                }
                
                return headers;
            }
        } catch (error) {
            console.warn('提取工作表表头失败:', error);
        }
        
        return [];
    }
    
    /**
     * 检测列的数据类型
     */
    detectDataType(cell) {
        try {
            const value = cell.Value;
            if (value === null || value === undefined || value === '') {
                return 'empty';
            }
            
            if (typeof value === 'number') {
                return 'number';
            }
            
            if (typeof value === 'string') {
                // 检查是否为日期
                if (value instanceof Date || this.isDateString(value)) {
                    return 'date';
                }
                
                // 检查是否为货币
                if (this.isCurrencyString(value)) {
                    return 'currency';
                }
                
                return 'text';
            }
            
            return 'unknown';
            
        } catch (error) {
            return 'unknown';
        }
    }
    
    /**
     * 检测字符串是否为日期格式
     */
    isDateString(str) {
        const datePatterns = [
            /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/, // YYYY-MM-DD
            /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/, // MM-DD-YYYY
            /^\d{4}年\d{1,2}月\d{1,2}日$/,   // 中文日期
            /^\d{1,2}月\d{1,2}日$/           // 中文日期（无年份）
        ];
        
        return datePatterns.some(pattern => pattern.test(str));
    }
    
    /**
     * 检测字符串是否为货币格式
     */
    isCurrencyString(str) {
        const currencyPatterns = [
            /^[¥$€£]\s*\d+(\.\d{2})?$/, // ¥100, $50.00
            /^\d+(\.\d{2})?\s*[¥$€£]$/, // 100¥, 50.00$
            /^[\d,]+\.\d{2}\s*[元]$/     // 1,000.00元
        ];
        
        return currencyPatterns.some(pattern => pattern.test(str));
    }
    
    /**
     * 获取列字母
     */
    getColumnLetter(columnNumber) {
        let temp = '';
        let num = columnNumber;
        
        while (num > 0) {
            let remainder = (num - 1) % 26;
            temp = String.fromCharCode(65 + remainder) + temp;
            num = Math.floor((num - 1) / 26);
        }
        
        return temp;
    }
    
    /**
     * 打开工作簿选择模态框
     */
    openWorkbookSelector() {
        try {
            this.allWorkbooks = this.getAllWorkbooks();
            this.renderWorkbookGrid();
            // 每次打开选择器时，重置UI选择状态以匹配当前选中的工作簿
            this.updateWorkbookGridSelection();
            document.getElementById('workbookModal').style.display = 'flex';
        } catch (error) {
            console.error('打开工作簿选择器失败:', error);
            // 使用更安全的方式调用通知
            if (this.formulaGenerator && typeof this.formulaGenerator.showNotification === 'function') {
                this.formulaGenerator.showNotification('打开工作簿选择器失败: ' + error.message, 'error');
            } else {
                // 备用方案：在控制台输出信息
                console.error('打开工作簿选择器失败:', error);
            }
        }
    }
    
    /**
     * 同步selectedWorkbooks与UI选择状态
     * 确保selectedWorkbooks只包含当前在UI中选中的工作簿
     */
    syncSelectedWorkbooksWithUI() {
        // 创建一个新的数组来存储当前选中的工作簿
        const currentlySelected = [];
        
        // 遍历所有工作簿，检查哪些在UI中被选中
        this.allWorkbooks.forEach(workbook => {
            const item = document.querySelector(`.workbook-grid-item[data-workbook-name="${workbook.name}"]`);
            if (item && item.classList.contains('selected')) {
                // 如果在UI中被选中，则添加到selectedWorkbooks中
                currentlySelected.push(workbook);
            }
        });
        
        // 更新selectedWorkbooks数组
        this.selectedWorkbooks = currentlySelected;
    }
    
    /**
     * 更新工作簿网格的选择状态以匹配当前选中的工作簿
     */
    updateWorkbookGridSelection() {
        const gridItems = document.querySelectorAll('.workbook-grid-item');
        gridItems.forEach(item => {
            const workbookName = item.dataset.workbookName;
            const isSelected = this.selectedWorkbooks.some(wb => wb.name === workbookName);
            if (isSelected) {
                item.classList.add('selected');
            } else {
                item.classList.remove('selected');
            }
        });
    }
    
    /**
     * 渲染工作簿网格
     */
    renderWorkbookGrid() {
        const grid = document.getElementById('workbookGrid');
        grid.innerHTML = '';
        
        this.allWorkbooks.forEach(workbook => {
            const item = this.createWorkbookGridItem(workbook);
            grid.appendChild(item);
        });
        
        // 渲染完网格后更新选择状态
        this.updateWorkbookGridSelection();
    }
    
    /**
     * 创建工作簿网格项
     */
    createWorkbookGridItem(workbook) {
        const item = document.createElement('div');
        item.className = 'workbook-grid-item';
        item.dataset.workbookName = workbook.name;
        item.dataset.workbookPath = workbook.path;
        
        item.innerHTML = `
            <div class="workbook-icon">📊</div>
            <div class="workbook-name">${workbook.name}</div>
            <div class="workbook-info">
                <div class="info-item">${workbook.worksheets.length} 个工作表</div>
                ${workbook.isActive ? '<div class="active-badge">当前</div>' : ''}
            </div>
        `;
        
        item.addEventListener('click', () => {
            this.toggleWorkbookSelection(item);
        });
        
        return item;
    }
    
    /**
     * 切换工作簿选择状态
     */
    toggleWorkbookSelection(item) {
        item.classList.toggle('selected');
        
        const workbookName = item.dataset.workbookName;
        
        if (item.classList.contains('selected')) {
            // 检查工作簿是否已经在selectedWorkbooks中
            const existingIndex = this.selectedWorkbooks.findIndex(wb => wb.name === workbookName);
            if (existingIndex === -1) {
                // 如果不在selectedWorkbooks中，则添加
                const workbook = this.allWorkbooks.find(wb => wb.name === workbookName);
                if (workbook) {
                    this.selectedWorkbooks.push(workbook);
                }
            }
        } else {
            // 从selectedWorkbooks中移除
            this.selectedWorkbooks = this.selectedWorkbooks.filter(wb => wb.name !== workbookName);
        }
    }
    
    /**
     * 确认选择工作簿
     */
    confirmSelection() {
        try {
            // 注意：根据需求规范，不应该默认选择所有工作簿
            // 如果没有选择任何工作簿，则保持为空
            
            // 为每个选中的工作簿加载详细的工作表信息（包括表头）
            this.selectedWorkbooks.forEach(workbook => {
                if (!workbook.worksheets || workbook.worksheets.length === 0) {
                    workbook.worksheets = this.getWorkbookWorksheets(workbook);
                }
            });
            
            // 更新公式生成器的选择
            // 修复：使用正确的属性名 selectedWorkbooks（而不是selectedWorkbook）
            if (this.formulaGenerator) {
                this.formulaGenerator.selectedWorkbooks = [...this.selectedWorkbooks];
                this.formulaGenerator.updateSelectedSources();
            }
            
            // 使用全局函数关闭模态框
            if (typeof closeWorkbookModal === 'function') {
                closeWorkbookModal();
            } else {
                // 备用方案：直接操作DOM
                const workbookModal = document.getElementById('workbookModal');
                if (workbookModal) {
                    workbookModal.style.display = 'none';
                }
            }
            
            // 显示工作表选择区域（如果是跨工作簿模式）
            const worksheetSelection = document.getElementById('worksheetSelection');
            if (worksheetSelection) {
                // 检查是否为跨工作簿模式
                const isWorkbookMode = document.querySelector('input[name="referenceType"][value="workbook"]')?.checked;
                if (isWorkbookMode && this.selectedWorkbooks.length > 0) {
                    worksheetSelection.style.display = 'block';
                    // 加载选中工作簿的工作表列表
                    this.loadSelectedWorkbookWorksheets();
                } else {
                    worksheetSelection.style.display = 'none';
                }
            }
            
            // 修复：使用更安全的方式调用通知
            if (this.formulaGenerator && typeof this.formulaGenerator.showNotification === 'function') {
                this.formulaGenerator.showNotification(`已选择 ${this.selectedWorkbooks.length} 个工作簿`, 'success');
            } else {
                // 备用方案：在控制台输出信息
                console.log(`已选择 ${this.selectedWorkbooks.length} 个工作簿`);
            }
            
            // 触发更新已选择数据源显示
            if (window.aiHelperMainInstance && typeof window.aiHelperMainInstance.updateSelectedSourcesDisplay === 'function') {
                window.aiHelperMainInstance.updateSelectedSourcesDisplay();
            }
            
            return true;
        } catch (error) {
            console.error('确认工作簿选择失败:', error);
            // 修复：使用更安全的方式调用通知
            if (this.formulaGenerator && typeof this.formulaGenerator.showNotification === 'function') {
                this.formulaGenerator.showNotification('确认工作簿选择失败: ' + error.message, 'error');
            }
            return false;
        }
    }

    /**
     * 加载选中工作簿的工作表列表（跨工作簿模式）
     */
    loadSelectedWorkbookWorksheets() {
        try {
            var workbookList = document.getElementById('workbookList');
            if (!workbookList) return;
            
            // 清空之前的内容
            workbookList.innerHTML = '';
            
            var html = '';
            
            // 为每个选中的工作簿生成工作表列表
            this.selectedWorkbooks.forEach(workbook => {
                html += '<div class="workbook-item" data-workbook="' + workbook.name + '">' +
                        '<div class="workbook-header">' +
                        '<h4>' + workbook.name + '</h4>' +
                        '</div>' +
                        '<div class="worksheet-list">';
                
                // 获取工作簿的所有工作表
                if (workbook.worksheets) {
                    workbook.worksheets.forEach(worksheet => {
                        // 检查该工作表是否已经被选中
                        const isSelected = this.formulaGenerator && this.formulaGenerator.selectedWorksheets && 
                            this.formulaGenerator.selectedWorksheets.some(item => 
                                item.workbook === workbook.name && item.worksheet === worksheet.name
                            );
                        
                        html += '<div class="worksheet-item">' +
                                '<input type="checkbox" id="ws_' + workbook.name + '_' + worksheet.name + '" ' +
                                'name="worksheets" value="' + worksheet.name + '" data-workbook="' + workbook.name + '"' + 
                                (isSelected ? ' checked' : '') + '>' +
                                '<label for="ws_' + workbook.name + '_' + worksheet.name + '">' +
                                worksheet.name +
                                '</label>' +
                                '</div>';
                    });
                }
                
                html += '</div></div>';
            });
            
            workbookList.innerHTML = html;
            
            // 绑定工作表选择事件
            var worksheetCheckboxes = document.querySelectorAll('input[name="worksheets"]');
            var self = this;
            for (var j = 0; j < worksheetCheckboxes.length; j++) {
                worksheetCheckboxes[j].addEventListener('change', function(e) {
                    // 直接在workbookSelector中处理工作表选择
                    self.handleWorksheetSelectionInWorkbookMode(e.target);
                });
            }
            
        } catch (error) {
            console.error('加载选中工作簿工作表列表失败:', error);
        }
    }
    
    /**
     * 在跨工作簿模式下处理工作表选择
     */
    handleWorksheetSelectionInWorkbookMode(checkbox) {
        var workbookName = checkbox.dataset.workbook;
        var worksheetName = checkbox.value;
        
        // 初始化selectedWorksheets数组（如果不存在）
        if (!this.formulaGenerator) {
            this.formulaGenerator = {};
        }
        if (!this.formulaGenerator.selectedWorksheets) {
            this.formulaGenerator.selectedWorksheets = [];
        }
        
        if (checkbox.checked) {
            // 添加到选中的工作表列表
            var found = false;
            for (var i = 0; i < this.formulaGenerator.selectedWorksheets.length; i++) {
                if (this.formulaGenerator.selectedWorksheets[i].workbook === workbookName && 
                    this.formulaGenerator.selectedWorksheets[i].worksheet === worksheetName) {
                    found = true;
                    break;
                }
            }
            if (!found) {
                this.formulaGenerator.selectedWorksheets.push({ workbook: workbookName, worksheet: worksheetName });
            }
        } else {
            // 从选中列表中移除
            var newSelectedWorksheets = [];
            for (var j = 0; j < this.formulaGenerator.selectedWorksheets.length; j++) {
                if (!(this.formulaGenerator.selectedWorksheets[j].workbook === workbookName && 
                      this.formulaGenerator.selectedWorksheets[j].worksheet === worksheetName)) {
                    newSelectedWorksheets.push(this.formulaGenerator.selectedWorksheets[j]);
                }
            }
            this.formulaGenerator.selectedWorksheets = newSelectedWorksheets;
        }
        
        console.log('📋 [handleWorksheetSelectionInWorkbookMode] 当前选中的工作表:', this.formulaGenerator.selectedWorksheets);
    }
    
    /**
     * 显示工作表选择对话框
     */
    showWorksheetSelection(workbook) {
        // 这里可以实现工作表选择的逻辑
        // 目前我们先简单地选择所有工作表
        if (!workbook.worksheets || workbook.worksheets.length === 0) {
            workbook.worksheets = this.getWorkbookWorksheets({name: workbook.name, path: workbook.path});
        }
    }

    /**
     * 过滤工作簿列表
     */
    filterWorkbooks(searchTerm) {
        const items = document.querySelectorAll('.workbook-grid-item');
        
        items.forEach(item => {
            const name = item.dataset.workbookName || '';
            const shouldShow = name.toLowerCase().includes(searchTerm.toLowerCase());
            item.style.display = shouldShow ? 'block' : 'none';
        });
    }
    
    /**
     * 初始化搜索框事件监听
     */
    initSearchBox() {
        const searchBox = document.getElementById('workbookSearch');
        if (searchBox) {
            searchBox.addEventListener('input', (e) => {
                this.filterWorkbooks(e.target.value);
            });
        }
    }
    
    /**
     * 获取选择的工作簿信息摘要
     */
    getSelectionSummary() {
        if (this.selectedWorkbooks.length === 0) {
            return '未选择工作簿';
        }
        
        const totalWorksheets = this.selectedWorkbooks.reduce((sum, wb) => {
            return sum + wb.worksheets.length;
        }, 0);
        
        return `${this.selectedWorkbooks.length} 个工作簿，${totalWorksheets} 个工作表`;
    }
    
    /**
     * 获取选择的工作簿列表
     * 返回符合AIInterface格式的工作簿信息
     */
    getSelectedWorkbooks() {
        try {
            // 如果没有选择任何工作簿，返回空数组
            if (this.selectedWorkbooks.length === 0) {
                console.log('📋 [workbookSelector.getSelectedWorkbooks] 没有选择工作簿，返回空数组');
                return [];
            }
            
            // 转换为AIInterface需要的格式
            const formattedWorkbooks = this.selectedWorkbooks.map(workbook => {
                // 检查是否有特定选择的工作表
                let worksheetsToSend = workbook.worksheets;
                
                // 如果在formulaGenerator中有特定选择的工作表，则只发送这些工作表
                if (this.formulaGenerator && this.formulaGenerator.selectedWorksheets && 
                    this.formulaGenerator.selectedWorksheets.length > 0) {
                    const selectedWorksheetsForThisWorkbook = this.formulaGenerator.selectedWorksheets
                        .filter(item => item.workbook === workbook.name)
                        .map(item => item.worksheet);
                    
                    // 如果有特定选择的工作表，则只发送这些工作表
                    if (selectedWorksheetsForThisWorkbook.length > 0) {
                        worksheetsToSend = workbook.worksheets.filter(worksheet => 
                            selectedWorksheetsForThisWorkbook.includes(worksheet.name)
                        );
                    }
                }
                
                return {
                    workBookName: workbook.name,
                    workBookPath: workbook.path,
                    worksheets: worksheetsToSend.map(worksheet => {
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
            
            console.log('📊 [workbookSelector.getSelectedWorkbooks] 返回格式化的选中工作簿:', formattedWorkbooks);
            return formattedWorkbooks;
            
        } catch (error) {
            console.error('⚠️ [workbookSelector.getSelectedWorkbooks] 获取选中工作簿失败:', error);
            return [];
        }
    }
    
    /**
     * 获取特定工作表的详细信息
     */
    getWorksheetDetails(workbookName, worksheetName) {
        try {
            const workbook = this.allWorkbooks.find(wb => wb.name === workbookName);
            if (!workbook) {
                return null;
            }
            
            const worksheet = workbook.worksheets.find(ws => ws.name === worksheetName);
            return worksheet || null;
            
        } catch (error) {
            console.warn(`获取工作表详细信息失败: ${workbookName} - ${worksheetName}`, error);
            return null;
        }
    }
    
    /**
     * 获取工作表数据范围
     */
    getWorksheetDataRange(workbookName, worksheetName) {
        try {
            const worksheetDetails = this.getWorksheetDetails(workbookName, worksheetName);
            if (!worksheetDetails) {
                return null;
            }
            
            const { usedRange } = worksheetDetails;
            return {
                workbookName: workbookName,
                worksheetName: worksheetName,
                rangeAddress: usedRange.address,
                startCell: `${this.getColumnLetter(usedRange.startCol)}${usedRange.startRow}`,
                endCell: `${this.getColumnLetter(usedRange.endCol)}${usedRange.endRow}`,
                dimensions: {
                    rows: usedRange.rows,
                    columns: usedRange.columns
                }
            };
            
        } catch (error) {
            console.warn(`获取工作表数据范围失败: ${workbookName} - ${worksheetName}`, error);
            return null;
        }
    }
    
    /**
     * 验证数据源完整性
     */
    validateDataSources() {
        const errors = [];
        
        // 检查是否有选中的工作簿
        if (this.selectedWorkbooks.length === 0) {
            errors.push('请至少选择一个工作簿');
        }
        
        // 检查每个工作簿是否有数据
        this.selectedWorkbooks.forEach(workbook => {
            if (workbook.worksheets.length === 0) {
                errors.push(`工作簿 "${workbook.name}" 没有工作表`);
            }
            
            workbook.worksheets.forEach(worksheet => {
                if (worksheet.usedRange.rows === 0 || worksheet.usedRange.columns === 0) {
                    errors.push(`工作表 "${workbook.name} - ${worksheet.name}" 没有数据`);
                }
            });
        });
        
        return {
            isValid: errors.length === 0,
            errors: errors
        };
    }
    
    /**
     * 清空选择
     */
    clearSelection() {
        this.selectedWorkbooks = [];
        
        // 更新UI
        document.querySelectorAll('.workbook-grid-item').forEach(item => {
            item.classList.remove('selected');
        });
        
        this.formulaGenerator.updateSelectedSources();
    }
    
    /**
     * 获取所有可用工作表的完整信息
     */
    getAllWorksheetInfo() {
        const allInfo = [];
        
        this.allWorkbooks.forEach(workbook => {
            workbook.worksheets.forEach(worksheet => {
                allInfo.push({
                    workbookName: workbook.name,
                    workbookPath: workbook.path,
                    worksheetName: worksheet.name,
                    usedRange: worksheet.usedRange,
                    headers: worksheet.headers,
                    dataType: this.analyzeWorksheetDataTypes(worksheet)
                });
            });
        });
        
        return allInfo;
    }
    
    /**
     * 分析工作表数据类型
     */
    analyzeWorksheetDataTypes(worksheet) {
        const types = {
            number: 0,
            text: 0,
            date: 0,
            currency: 0,
            empty: 0
        };
        
        worksheet.headers.forEach(header => {
            types[header.dataType] = (types[header.dataType] || 0) + 1;
        });
        
        return types;
    }
}

// 导出供其他模块使用
window.WorkbookSelector = WorkbookSelector;