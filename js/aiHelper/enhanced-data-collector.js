/**
 * 增强的数据收集器
 * 负责收集完整的工作表数据、跨工作表/工作簿信息，以及丰富的上下文数据
 * 用于向AI提供完整准确的数据上下文
 */
class EnhancedDataCollector {
    constructor() {
        this.debugMode = true; // 开启调试模式
        this.maxDataRows = 100; // 限制收集的数据行数，避免请求过大
        this.maxDataColumns = 50; // 限制收集的列数
    }

    /**
     * 收集完整的工作表数据上下文
     * @param {Object} options - 收集选项
     * @returns {Object} 完整的数据上下文
     */
    async collectComprehensiveData(options = {}) {
        try {
            console.log('🔄 [EnhancedDataCollector] 开始收集综合数据上下文...');
            
            const context = {
                timestamp: new Date().toISOString(),
                collectionVersion: '1.0.0',
                sessionId: this.generateSessionId(),
                collectionOptions: options,
                currentContext: {},
                workbookContext: {},
                selectionContext: {},
                dataContext: {},
                referenceContext: {}
            };

            // 收集当前上下文信息
            console.log('📍 [EnhancedDataCollector] 收集当前上下文...');
            context.currentContext = await this.collectCurrentContext();
            
            // 收集工作簿上下文
            console.log('📊 [EnhancedDataCollector] 收集工作簿上下文...');
            context.workbookContext = await this.collectWorkbookContext();
            
            // 收集选择上下文
            console.log('🎯 [EnhancedDataCollector] 收集选择上下文...');
            context.selectionContext = await this.collectSelectionContext();
            
            // 收集数据上下文
            console.log('📈 [EnhancedDataCollector] 收集数据上下文...');
            context.dataContext = await this.collectDataContext();
            
            // 收集引用上下文
            console.log('🔗 [EnhancedDataCollector] 收集引用上下文...');
            context.referenceContext = await this.collectReferenceContext(options);

            console.log('✅ [EnhancedDataCollector] 数据收集完成');
            console.log('📊 [EnhancedDataCollector] 收集统计:', {
                currentSheets: context.currentContext.sheets?.length || 0,
                selectedWorkbooks: context.workbookContext.selectedWorkbooks?.length || 0,
                dataRows: context.dataContext.sampleData?.length || 0,
                references: context.referenceContext.references?.length || 0
            });

            return context;

        } catch (error) {
            console.error('❌ [EnhancedDataCollector] 数据收集失败:', error);
            throw new Error(`数据收集失败: ${error.message}`);
        }
    }

    /**
     * 收集当前上下文信息
     * @returns {Object} 当前上下文
     */
    async collectCurrentContext() {
        const context = {
            timestamp: new Date().toISOString(),
            application: await this.getApplicationInfo(),
            activeWorkbook: await this.getActiveWorkbookInfo(),
            activeWorksheet: await this.getActiveWorksheetInfo(),
            activeCell: await this.getActiveCellInfo(),
            sheets: []
        };

        try {
            // 收集所有工作表信息
            if (window.Application && window.Application.ActiveWorkbook) {
                const workbook = window.Application.ActiveWorkbook;
                for (let i = 1; i <= workbook.Worksheets.Count; i++) {
                    const sheet = workbook.Worksheets(i);
                    context.sheets.push({
                        name: sheet.Name,
                        index: i,
                        visible: sheet.Visible,
                        codeName: sheet.CodeName || sheet.Name
                    });
                }
            }
        } catch (error) {
            console.warn('⚠️ [EnhancedDataCollector] 收集工作表信息失败:', error);
        }

        return context;
    }

    /**
     * 收集工作簿上下文
     * @returns {Object} 工作簿上下文
     */
    async collectWorkbookContext() {
        const context = {
            selectedWorkbooks: [],
            availableWorkbooks: [],
            workbookRelationships: [],
            globalNames: []
        };

        try {
            if (window.Application) {
                // 收集已打开的工作簿
                for (let i = 1; i <= window.Application.Workbooks.Count; i++) {
                    const workbook = window.Application.Workbooks(i);
                    
                    const workbookInfo = {
                        name: workbook.Name,
                        fullPath: workbook.FullName || '',
                        path: workbook.Path || '',
                        saved: workbook.Saved,
                        isActive: workbook === window.Application.ActiveWorkbook,
                        readOnly: workbook.ReadOnly || false,
                        sheets: []
                    };

                    // 收集工作簿中的工作表信息
                    for (let j = 1; j <= workbook.Worksheets.Count; j++) {
                        const sheet = workbook.Worksheets(j);
                        workbookInfo.sheets.push({
                            name: sheet.Name,
                            index: j,
                            visible: sheet.Visible
                        });
                    }

                    if (workbookInfo.isActive) {
                        context.selectedWorkbooks.push(workbookInfo);
                    } else {
                        context.availableWorkbooks.push(workbookInfo);
                    }
                }

                // 收集全局名称（如果支持）
                if (window.Application.ActiveWorkbook && window.Application.ActiveWorkbook.Names) {
                    const names = window.Application.ActiveWorkbook.Names;
                    for (let i = 1; i <= names.Count; i++) {
                        try {
                            const name = names.Item(i);
                            context.globalNames.push({
                                name: name.Name,
                                refersTo: name.RefersTo,
                                comment: name.Comment || ''
                            });
                        } catch (nameError) {
                            // 跳过无法访问的名称
                        }
                    }
                }
            }
        } catch (error) {
            console.warn('⚠️ [EnhancedDataCollector] 收集工作簿信息失败:', error);
        }

        return context;
    }

    /**
     * 收集选择上下文
     * @returns {Object} 选择上下文
     */
    async collectSelectionContext() {
        const context = {
            selectionRange: {},
            selectedCells: [],
            selectionBounds: {},
            contextCells: []
        };

        try {
            if (window.Application && window.Application.Selection) {
                const selection = window.Application.Selection;
                
                // 基本选择信息
                context.selectionRange = {
                    address: selection.Address(),
                    row: selection.Row,
                    col: selection.Column,
                    rowCount: selection.Rows.Count,
                    colCount: selection.Columns.Count,
                    firstCell: selection.Item(1).Address,
                    lastCell: selection.Item(selection.Rows.Count, selection.Columns.Count).Address
                };

                // 收集选择的单元格数据
                const maxCells = Math.min(selection.Cells.Count, 100); // 限制选择数量
                for (let i = 1; i <= maxCells; i++) {
                    try {
                        const cell = selection.Item(i);
                        context.selectedCells.push({
                            address: cell.Address,
                            row: cell.Row,
                            col: cell.Column,
                            value: cell.Value,
                            formula: cell.Formula,
                            numberFormat: cell.NumberFormat
                        });
                    } catch (cellError) {
                        // 跳过无法访问的单元格
                    }
                }

                // 计算选择边界
                context.selectionBounds = {
                    topRow: Math.min(...context.selectedCells.map(c => c.row)),
                    leftCol: Math.min(...context.selectedCells.map(c => c.col)),
                    bottomRow: Math.max(...context.selectedCells.map(c => c.row)),
                    rightCol: Math.max(...context.selectedCells.map(c => c.col))
                };

                // 收集上下文单元格（选择周围的单元格）
                context.contextCells = await this.collectContextCells(context.selectionBounds);
            }
        } catch (error) {
            console.warn('⚠️ [EnhancedDataCollector] 收集选择信息失败:', error);
        }

        return context;
    }

    /**
     * 收集数据上下文
     * @returns {Object} 数据上下文
     */
    async collectDataContext() {
        const context = {
            headers: [],
            sampleData: [],
            dataStructure: {},
            columnTypes: {},
            dataQuality: {}
        };

        try {
            if (window.Application && window.Application.ActiveSheet) {
                const activeSheet = window.Application.ActiveSheet;
                
                // 收集表头信息
                context.headers = await this.collectHeaders(activeSheet);
                
                // 收集样本数据
                context.sampleData = await this.collectSampleData(activeSheet, context.headers);
                
                // 分析数据结构
                context.dataStructure = this.analyzeDataStructure(context.sampleData, context.headers);
                
                // 分析列类型
                context.columnTypes = this.analyzeColumnTypes(context.sampleData, context.headers);
                
                // 评估数据质量
                context.dataQuality = this.assessDataQuality(context.sampleData, context.headers);
            }
        } catch (error) {
            console.warn('⚠️ [EnhancedDataCollector] 收集数据上下文失败:', error);
        }

        return context;
    }

    /**
     * 收集引用上下文
     * @param {Object} options - 收集选项
     * @returns {Object} 引用上下文
     */
    async collectReferenceContext(options = {}) {
        const context = {
            references: [],
            crossWorkbookRefs: [],
            externalConnections: [],
            formulaDependencies: []
        };

        try {
            // 如果指定了跨工作表或跨工作簿引用
            if (options.referenceType === 'cross_workbook' || options.referenceType === 'cross_worksheet') {
                context.references = await this.collectFormulaReferences(options);
            }

            // 检查外部连接
            if (window.Application && window.Application.ActiveWorkbook) {
                context.externalConnections = await this.collectExternalConnections();
            }
        } catch (error) {
            console.warn('⚠️ [EnhancedDataCollector] 收集引用上下文失败:', error);
        }

        return context;
    }

    // 辅助方法实现
    async getApplicationInfo() {
        try {
            return {
                name: window.Application ? window.Application.Name : 'Unknown',
                version: window.Application ? window.Application.Version : 'Unknown',
                build: window.Application ? window.Application.Build : 'Unknown'
            };
        } catch (error) {
            return {
                name: 'Unknown',
                version: 'Unknown',
                build: 'Unknown'
            };
        }
    }

    async getActiveWorkbookInfo() {
        try {
            if (window.Application && window.Application.ActiveWorkbook) {
                const wb = window.Application.ActiveWorkbook;
                return {
                    name: wb.Name,
                    fullPath: wb.FullName || '',
                    path: wb.Path || '',
                    saved: wb.Saved,
                    readOnly: wb.ReadOnly || false,
                    sheetsCount: wb.Worksheets.Count
                };
            }
            return null;
        } catch (error) {
            console.warn('获取活动工作簿信息失败:', error);
            return null;
        }
    }

    async getActiveWorksheetInfo() {
        try {
            if (window.Application && window.Application.ActiveSheet) {
                const ws = window.Application.ActiveSheet;
                return {
                    name: ws.Name,
                    index: ws.Index,
                    visible: ws.Visible,
                    codeName: ws.CodeName || ws.Name,
                    cellsCount: ws.Cells.Count,
                    usedRange: ws.UsedRange ? ws.UsedRange.Address : 'A1'
                };
            }
            return null;
        } catch (error) {
            console.warn('获取活动工作表信息失败:', error);
            return null;
        }
    }

    async getActiveCellInfo() {
        try {
            if (window.Application && window.Application.ActiveCell) {
                const cell = window.Application.ActiveCell;
                return {
                    address: cell.Address,
                    row: cell.Row,
                    col: cell.Column,
                    value: cell.Value,
                    formula: cell.Formula,
                    numberFormat: cell.NumberFormat
                };
            }
            return null;
        } catch (error) {
            console.warn('获取活动单元格信息失败:', error);
            return null;
        }
    }

    async collectHeaders(activeSheet) {
        const headers = [];
        try {
            // 尝试获取第一行作为表头
            const headerRange = activeSheet.Range('1:1');
            const headerValues = headerRange.Value;
            
            if (headerValues && headerValues[0]) {
                for (let i = 0; i < headerValues[0].length; i++) {
                    if (headerValues[0][i] !== null && headerValues[0][i] !== undefined) {
                        headers.push({
                            column: this.indexToColumnLetter(i + 1),
                            index: i + 1,
                            name: String(headerValues[0][i]).trim(),
                            type: 'unknown'
                        });
                    }
                }
            }
        } catch (error) {
            console.warn('收集表头信息失败:', error);
        }
        return headers;
    }

    async collectSampleData(activeSheet, headers) {
        const sampleData = [];
        try {
            const maxRows = Math.min(this.maxDataRows, 50); // 默认收集50行样本数据
            const maxCols = Math.min(headers.length, this.maxDataColumns);
            
            if (headers.length > 0) {
                const dataRange = activeSheet.Range(`A1:${this.indexToColumnLetter(maxCols)}${maxRows}`);
                const dataValues = dataRange.Value;
                
                if (dataValues) {
                    for (let row = 1; row < dataValues.length && row < maxRows; row++) {
                        if (dataValues[row]) {
                            const rowData = {};
                            for (let col = 0; col < Math.min(maxCols, dataValues[row].length); col++) {
                                const columnLetter = this.indexToColumnLetter(col + 1);
                                rowData[columnLetter] = dataValues[row][col];
                            }
                            sampleData.push(rowData);
                        }
                    }
                }
            }
        } catch (error) {
            console.warn('收集样本数据失败:', error);
        }
        return sampleData;
    }

    async collectContextCells(bounds) {
        const contextCells = [];
        try {
            // 收集选择周围5x5区域的单元格
            const startRow = Math.max(1, bounds.topRow - 5);
            const endRow = bounds.bottomRow + 5;
            const startCol = Math.max(1, bounds.leftCol - 5);
            const endCol = bounds.rightCol + 5;
            
            const contextRange = window.Application.ActiveSheet.Range(
                `${this.indexToColumnLetter(startCol)}${startRow}:${this.indexToColumnLetter(endCol)}${endRow}`
            );
            
            const contextValues = contextRange.Value;
            if (contextValues) {
                for (let row = 0; row < contextValues.length; row++) {
                    if (contextValues[row]) {
                        for (let col = 0; col < contextValues[row].length; col++) {
                            const cellValue = contextValues[row][col];
                            if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                                contextCells.push({
                                    address: `${this.indexToColumnLetter(startCol + col)}${startRow + row}`,
                                    row: startRow + row,
                                    col: startCol + col,
                                    value: cellValue
                                });
                            }
                        }
                    }
                }
            }
        } catch (error) {
            console.warn('收集上下文单元格失败:', error);
        }
        return contextCells;
    }

    analyzeDataStructure(sampleData, headers) {
        const structure = {
            rowCount: sampleData.length,
            columnCount: headers.length,
            hasHeaders: headers.length > 0,
            dataTypes: {},
            emptyRows: 0,
            filledRows: 0
        };

        try {
            sampleData.forEach(row => {
                const hasData = Object.values(row).some(val => val !== null && val !== undefined && val !== '');
                if (hasData) {
                    structure.filledRows++;
                } else {
                    structure.emptyRows++;
                }
            });
        } catch (error) {
            console.warn('分析数据结构失败:', error);
        }

        return structure;
    }

    analyzeColumnTypes(sampleData, headers) {
        const columnTypes = {};
        
        headers.forEach(header => {
            const column = header.column;
            const values = sampleData.map(row => row[column]).filter(val => val !== null && val !== undefined && val !== '');
            
            if (values.length === 0) {
                columnTypes[column] = 'empty';
            } else {
                // 简单的类型检测
                const numericValues = values.filter(val => !isNaN(Number(val)));
                const dateValues = values.filter(val => !isNaN(Date.parse(val)));
                
                if (numericValues.length === values.length) {
                    columnTypes[column] = 'number';
                } else if (dateValues.length > values.length * 0.5) {
                    columnTypes[column] = 'date';
                } else {
                    columnTypes[column] = 'text';
                }
            }
        });
        
        return columnTypes;
    }

    assessDataQuality(sampleData, headers) {
        const quality = {
            completeness: 0,
            consistency: 0,
            accuracy: 0,
            issues: []
        };

        try {
            if (sampleData.length === 0) {
                quality.issues.push('无样本数据');
                return quality;
            }

            const totalCells = sampleData.length * headers.length;
            const filledCells = sampleData.reduce((count, row) => {
                return count + Object.values(row).filter(val => val !== null && val !== undefined && val !== '').length;
            }, 0);

            quality.completeness = Math.round((filledCells / totalCells) * 100);

            // 检查数据类型一致性
            const columnTypes = this.analyzeColumnTypes(sampleData, headers);
            const consistentColumns = Object.values(columnTypes).filter(type => type !== 'unknown').length;
            quality.consistency = Math.round((consistentColumns / headers.length) * 100);

            // 整体准确性评估（简单实现）
            quality.accuracy = Math.min(quality.completeness, quality.consistency);

        } catch (error) {
            console.warn('评估数据质量失败:', error);
            quality.issues.push('数据质量评估失败');
        }

        return quality;
    }

    async collectFormulaReferences(options) {
        const references = [];
        // 这里可以实现公式引用分析逻辑
        return references;
    }

    async collectExternalConnections() {
        const connections = [];
        // 这里可以实现外部连接收集逻辑
        return connections;
    }

    indexToColumnLetter(index) {
        let temp = '';
        let columnNumber = index;
        
        while (columnNumber > 0) {
            let remainder = (columnNumber - 1) % 26;
            temp = String.fromCharCode(65 + remainder) + temp;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        
        return temp;
    }

    generateSessionId() {
        return 'session_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    /**
     * 将收集的数据转换为AI友好的格式
     * @param {Object} context - 收集的上下文数据
     * @returns {Object} AI格式的数据
     */
    convertToAIFormat(context) {
        const aiFormat = {
            session_info: {
                session_id: context.sessionId,
                timestamp: context.timestamp,
                collection_version: context.collectionVersion
            },
            
            current_worksheet: {
                name: context.currentContext.activeWorksheet?.name || '未知工作表',
                headers: context.dataContext.headers.map(h => h.name),
                structure: context.dataContext.dataStructure,
                sample_data: context.dataContext.sampleData.slice(0, 10) // 只保留前10行
            },
            
            selection_info: {
                cell_address: context.currentContext.activeCell?.address || '',
                selection_range: context.selectionContext.selectionRange.address || '',
                context_cells: context.selectionContext.contextCells.slice(0, 20) // 只保留前20个上下文单元格
            },
            
            workbook_info: {
                workbook_name: context.workbookContext.selectedWorkbooks[0]?.name || '',
                available_workbooks: context.workbookContext.availableWorkbooks.map(wb => wb.name),
                global_names: context.workbookContext.globalNames.map(n => n.name)
            },
            
            data_analysis: {
                column_types: context.dataContext.columnTypes,
                data_quality: context.dataContext.dataQuality
            },
            
            reference_info: {
                references: context.referenceContext.references,
                external_connections: context.referenceContext.externalConnections
            }
        };

        return aiFormat;
    }

    /**
     * 生成详细的调试报告
     * @param {Object} context - 收集的上下文数据
     * @returns {string} 调试报告
     */
    generateDebugReport(context) {
        let report = `🔍 EnhancedDataCollector 调试报告\n`;
        report += `==========================================\n\n`;
        
        report += `会话信息:\n`;
        report += `- 会话ID: ${context.sessionId}\n`;
        report += `- 收集时间: ${context.timestamp}\n`;
        report += `- 收集版本: ${context.collectionVersion}\n\n`;
        
        report += `当前工作表:\n`;
        report += `- 名称: ${context.currentContext.activeWorksheet?.name || '未知'}\n`;
        report += `- 表头数量: ${context.dataContext.headers.length}\n`;
        report += `- 样本数据行数: ${context.dataContext.sampleData.length}\n\n`;
        
        report += `选择信息:\n`;
        report += `- 当前单元格: ${context.currentContext.activeCell?.address || '未知'}\n`;
        report += `- 选择范围: ${context.selectionContext.selectionRange.address || '未知'}\n`;
        report += `- 上下文单元格: ${context.selectionContext.contextCells.length}\n\n`;
        
        report += `工作簿信息:\n`;
        report += `- 已选工作簿: ${context.workbookContext.selectedWorkbooks.length}\n`;
        report += `- 可用工作簿: ${context.workbookContext.availableWorkbooks.length}\n`;
        report += `- 全局名称: ${context.workbookContext.globalNames.length}\n\n`;
        
        report += `数据质量:\n`;
        report += `- 完整性: ${context.dataContext.dataQuality.completeness}%\n`;
        report += `- 一致性: ${context.dataContext.dataQuality.consistency}%\n`;
        report += `- 准确性: ${context.dataContext.dataQuality.accuracy}%\n\n`;
        
        report += `参考信息:\n`;
        report += `- 引用数量: ${context.referenceContext.references.length}\n`;
        report += `- 外部连接: ${context.referenceContext.externalConnections.length}\n`;
        
        return report;
    }
}

// 导出数据收集器类
if (typeof module !== 'undefined' && module.exports) {
    module.exports = EnhancedDataCollector;
} else {
    window.EnhancedDataCollector = EnhancedDataCollector;
}