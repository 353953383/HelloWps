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
            this.loadAllWorkbooks();
            this.renderWorkbookGrid();
            document.getElementById('workbookModal').style.display = 'flex';
        } catch (error) {
            console.error('打开工作簿选择器失败:', error);
            this.formulaGenerator.showNotification('打开工作簿选择器失败: ' + error.message, 'error');
        }
    }
    
    /**
     * 关闭工作簿选择模态框
     */
    closeModal() {
        document.getElementById('workbookModal').style.display = 'none';
    }
    
    /**
     * 加载所有工作簿数据
     */
    loadAllWorkbooks() {
        this.allWorkbooks = this.getAllWorkbooks();
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
            if (!this.selectedWorkbooks.find(wb => wb.name === workbookName)) {
                this.selectedWorkbooks.push(this.allWorkbooks.find(wb => wb.name === workbookName));
            }
        } else {
            this.selectedWorkbooks = this.selectedWorkbooks.filter(wb => wb.name !== workbookName);
        }
    }
    
    /**
     * 确认选择工作簿
     */
    confirmSelection() {
        if (this.selectedWorkbooks.length === 0) {
            this.formulaGenerator.showNotification('请至少选择一个工作簿', 'warning');
            return false;
        }
        
        // 更新公式生成器的选择
        this.formulaGenerator.selectedWorkbooks = [...this.selectedWorkbooks];
        this.formulaGenerator.updateSelectedSources();
        
        this.closeModal();
        this.formulaGenerator.showNotification(`已选择 ${this.selectedWorkbooks.length} 个工作簿`, 'success');
        
        return true;
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