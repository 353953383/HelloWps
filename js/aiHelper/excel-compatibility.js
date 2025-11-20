/**
 * Excel对象兼容性测试
 * 测试WPS Office Excel对象模型的不同访问方式
 */

// 单元格访问方式测试
function testCellAccessMethods(worksheet) {
    const results = {
        methods: {},
        errors: []
    };
    
    // 方法1: worksheet.Cells.Item(row, col) - 按照WPS JSA规范
    try {
        const cell1 = worksheet.Cells.Item(1, 1);
        results.methods.cells = {
            success: true,
            result: cell1,
            type: typeof cell1
        };
    } catch (error) {
        results.methods.cells = {
            success: false,
            error: error.message
        };
        results.errors.push(`worksheet.Cells: ${error.message}`);
    }
    
    // 方法2: worksheet.Range.Cells(row, col)
    try {
        const cell2 = worksheet.Range.Cells(1, 1);
        results.methods.rangeCells = {
            success: true,
            result: cell2,
            type: typeof cell2
        };
    } catch (error) {
        results.methods.rangeCells = {
            success: false,
            error: error.message
        };
        results.errors.push(`worksheet.Range.Cells: ${error.message}`);
    }
    
    // 方法3: worksheet.getCell(row, col)
    try {
        const cell3 = worksheet.getCell(1, 1);
        results.methods.getCell = {
            success: true,
            result: cell3,
            type: typeof cell3
        };
    } catch (error) {
        results.methods.getCell = {
            success: false,
            error: error.message
        };
        results.errors.push(`worksheet.getCell: ${error.message}`);
    }
    
    // 方法4: worksheet.Range("A1")
    try {
        const cell4 = worksheet.Range("A1");
        results.methods.rangeA1 = {
            success: true,
            result: cell4,
            type: typeof cell4
        };
    } catch (error) {
        results.methods.rangeA1 = {
            success: false,
            error: error.message
        };
        results.errors.push(`worksheet.Range("A1"): ${error.message}`);
    }
    
    // 方法5: 使用Values数组（如果支持）
    try {
        if (worksheet.Values) {
            const values = worksheet.Values;
            results.methods.values = {
                success: true,
                result: values,
                type: typeof values,
                length: values ? values.length : 0
            };
        } else {
            results.methods.values = {
                success: false,
                error: "worksheet.Values not available"
            };
        }
    } catch (error) {
        results.methods.values = {
            success: false,
            error: error.message
        };
        results.errors.push(`worksheet.Values: ${error.message}`);
    }
    
    return results;
}

// 获取单元格值的通用方法
function getCellValueUniversal(worksheet, row, col) {
    const methods = [
        // 方法1: worksheet.Cells
        () => {
            try {
                const cell = worksheet.Cells.Item(row, col);
                return cell ? cell.Value : null;
            } catch (e) {
                return null;
            }
        },
        
        // 方法2: worksheet.Range.Cells
        () => {
            try {
                const cell = worksheet.Range.Cells(row, col);
                return cell ? cell.Value : null;
            } catch (e) {
                return null;
            }
        },
        
        // 方法3: worksheet.getCell
        () => {
            try {
                const cell = worksheet.getCell(row, col);
                return cell ? cell.Value : null;
            } catch (e) {
                return null;
            }
        },
        
        // 方法4: worksheet.Range
        () => {
            try {
                const colLetter = getColumnLetter(col);
                const cell = worksheet.Range(`${colLetter}${row}`);
                return cell ? cell.Value : null;
            } catch (e) {
                return null;
            }
        },
        
        // 方法5: 使用索引器（如果支持）
        () => {
            try {
                if (worksheet.Cells && worksheet.Cells.Item) {
                    const cell = worksheet.Cells.Item(row, col);
                    return cell ? cell.Value : null;
                }
            } catch (e) {
                return null;
            }
        }
    ];
    
    // 尝试每种方法
    for (let i = 0; i < methods.length; i++) {
        try {
            const value = methods[i]();
            if (value !== null && value !== undefined) {
                console.log(`✅ 方法${i + 1}成功获取单元格值:`, value);
                return value;
            }
        } catch (error) {
            console.log(`❌ 方法${i + 1}失败:`, error.message);
        }
    }
    
    console.warn('⚠️ 所有单元格访问方法都失败');
    return null;
}

// 获取列字母
function getColumnLetter(columnIndex) {
    let temp;
    let letter = '';
    while (columnIndex > 0) {
        temp = (columnIndex - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        columnIndex = (columnIndex - temp - 1) / 26;
    }
    return letter;
}

// 获取工作表范围
function getWorksheetRange(worksheet) {
    const methods = [
        // 方法1: UsedRange
        () => {
            try {
                const range = worksheet.UsedRange;
                return {
                    rows: range.Rows.Count,
                    columns: range.Columns.Count
                };
            } catch (e) {
                return null;
            }
        },
        
        // 方法2: Cells.SpecialCells
        () => {
            try {
                const range = worksheet.Cells.SpecialCells(11); // xlCellTypeConstants
                return {
                    rows: range.Rows.Count,
                    columns: range.Columns.Count
                };
            } catch (e) {
                return null;
            }
        },
        
        // 方法3: 计算最大值
        () => {
            try {
                let maxRow = 1, maxCol = 1;
                // 尝试查找使用的行和列
                const testRange = worksheet.Range("A1:Z100");
                for (let row = 1; row <= 100; row++) {
                    for (let col = 1; col <= 26; col++) {
                        try {
                            const cell = worksheet.Cells.Item(row, col);
                            if (cell.Value !== null && cell.Value !== undefined && cell.Value !== '') {
                                maxRow = Math.max(maxRow, row);
                                maxCol = Math.max(maxCol, col);
                            }
                        } catch (e) {
                            break;
                        }
                    }
                }
                return {
                    rows: maxRow,
                    columns: maxCol
                };
            } catch (e) {
                return null;
            }
        }
    ];
    
    for (let i = 0; i < methods.length; i++) {
        try {
            const result = methods[i]();
            if (result) {
                console.log(`✅ 范围检测方法${i + 1}成功:`, result);
                return result;
            }
        } catch (error) {
            console.log(`❌ 范围检测方法${i + 1}失败:`, error.message);
        }
    }
    
    console.warn('⚠️ 所有范围检测方法都失败，使用默认值');
    return { rows: 1, columns: 1 };
}

// 导出供外部使用
if (typeof window !== 'undefined') {
    window.ExcelCompatTest = {
        testCellAccessMethods,
        getCellValueUniversal,
        getWorksheetRange,
        getColumnLetter
    };
}