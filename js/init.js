
// 获取当前选中单元格范围的函数
function getSelectedRange() {
    try {
        let activeSheet = Application.ActiveSheet;
        if (!activeSheet) {
            return null;
        }
        
        let selection = $RangeSelection();
        if (!selection) {
            return null;
        }
        
        return {
            address: selection.Address(),
            worksheet: activeSheet.Name,
            workbook: activeSheet.Parent.Name,
            rowCount: selection.Rows.Count,
            columnCount: selection.Columns.Count,
            firstCell: selection.Item(1).Address,
            lastCell: selection.Item(selection.Rows.Count, selection.Columns.Count).Address,
            value: selection.Value2
        };
    } catch (ex) {
        $print("错误", "获取选中范围失败: " + ex.message);
        return null;
    }
}

// 获取当前活动单元格的函数
function getActiveCell() {
    try {
        let activeCell = Application.ActiveCell;
        if (!activeCell) {
            return null;
        }
        
        return {
            address: activeCell.Address,
            row: activeCell.Row,
            column: activeCell.Column,
            value: activeCell.Value,
            worksheet: activeCell.Worksheet.Name,
            workbook: activeCell.Worksheet.Parent.Name
        };
    } catch (ex) {
        $print("错误", "获取活动单元格失败: " + ex.message);
        return null;
    }
}

// 格式化范围信息的函数
function formatRangeInfo(range) {
    if (!range) return "无选中范围";
    
    return `${range.address} (${range.rowCount}行×${range.columnCount}列)`;
}

// 格式化单元格信息的函数
function formatCellInfo(cell) {
    if (!cell) return "无活动单元格";
    
    return `${cell.address} [${cell.row},${cell.column}] 值: ${cell.value || "空"}`;
}

// 获取工作簿保存信息的函数
function getWorkbookSaveInfo(workbook) {
    try {
        return {
            name: workbook.Name,
            fullPath: workbook.FullName || "",
            path: workbook.Path || "",
            saved: workbook.Saved,
            fileFormat: workbook.FileFormat || 0,
            isReadOnly: workbook.ReadOnly || false,
            isSaved: workbook.Saved || false
        };
    } catch (ex) {
        $print("错误", "获取工作簿信息失败: " + ex.message);
        return null;
    }
}

// 格式化保存信息的函数
function formatSaveInfo(info) {
    if (!info) return "无法获取保存信息";
    
    return `文件名: ${info.name}\n` +
           `完整路径: ${info.fullPath || "未保存"}\n` +
           `目录: ${info.path || "临时目录"}\n` +
           `格式: ${info.fileFormat}\n` +
           `状态: ${info.isSaved ? "已保存" : "未保存"}`;
}

function initApiEventListeners() {
    try {
        $print("事件", "监听建立中");
        // 确保Application对象存在
        if (typeof window.Application === 'undefined') {
            $print("错误", "Application对象未定义，无法注册事件监听器");
            return;
        }

        // Application级别事件
        Application.ApiEvent.AddApiEventListener("NewWorkbook", (wb)=>{
            $print("工作簿事件", "新建工作簿: " + wb.Name);
        });
        
        // Application.ApiEvent.AddApiEventListener("WorkbookOpen", (wb)=>{
        //     $print("工作簿事件", "打开工作簿: " + wb.Name);
        // });
        
        // Application.ApiEvent.AddApiEventListener("WorkbookBeforeClose", (wb)=>{
        //     $print("工作簿事件", "工作簿即将关闭: " + wb.Name);
        // });
        
        // Application.ApiEvent.AddApiEventListener("WorkbookBeforeSave", (wb)=>{
        //     $print("工作簿事件", "工作簿即将保存: " + wb.Name);
        // });
        
        Application.ApiEvent.AddApiEventListener("WorkbookAfterSave", (wb, success)=>{
            try {
                let saveInfo = {
                    workbookName: wb.Name,
                    success: success,
                    fullPath: wb.FullName || "未保存路径",
                    fileName: wb.Name,
                    path: wb.Path || "临时文件",
                    saved: wb.Saved,
                    fileFormat: wb.FileFormat || "未知格式"
                };
                
                $print("工作簿事件", "工作簿保存完成: " + saveInfo.workbookName + 
                      " 成功: " + saveInfo.success + 
                      " 路径: " + saveInfo.path + 
                      " 文件名: " + saveInfo.fileName);
                
                if (success) {
                    $print("保存详情", "完整路径: " + saveInfo.fullPath + 
                          " 文件格式: " + saveInfo.fileFormat + 
                          " 保存状态: " + (saveInfo.saved ? "已保存" : "未保存"));
                }
            } catch (ex) {
                $print("工作簿事件", "工作簿保存完成: " + wb.Name + " 成功: " + success + " 获取路径失败: " + ex.message);
            }
        });
        
        // Application.ApiEvent.AddApiEventListener("WorkbookBeforePrint", (wb)=>{
        //     $print("工作簿事件", "工作簿即将打印: " + wb.Name);
        // });
        
        // Application.ApiEvent.AddApiEventListener("WorkbookActivate", (wb)=>{
        //     $print("工作簿事件", "激活工作簿: " + wb.Name);
        // });
        
        // Application.ApiEvent.AddApiEventListener("WorkbookDeactivate", (wb)=>{
        //     $print("工作簿事件", "工作簿失活: " + wb.Name);
        // });
        
        Application.ApiEvent.AddApiEventListener("WorkbookNewSheet", (wb, sh)=>{
            $print("工作簿事件", "工作簿新增工作表: " + wb.Name + " -> " + sh.Name);
        });

        // 工作表级别事件
        // Application.ApiEvent.AddApiEventListener("SheetActivate", (sh)=>{
        //     $print("工作表事件", "激活工作表: " + sh.Name);
        // });
        
        // Application.ApiEvent.AddApiEventListener("SheetDeactivate", (sh)=>{
        //     $print("工作表事件", "工作表失活: " + sh.Name);
        // });
        
        Application.ApiEvent.AddApiEventListener("SheetBeforeDelete", (sh)=>{
            $print("工作表事件", "工作表即将删除: " + sh.Name + " 索引: " + sh.Index);
        });
        
        Application.ApiEvent.AddApiEventListener("SheetChange", (sh, target)=>{
            let rangeInfo = formatRangeInfo({
                address: target.Address(),
                rowCount: target.Rows.Count,
                columnCount: target.Columns.Count
            });
            $print("工作表事件", "工作表内容变更: " + sh.Name + " 范围: " + rangeInfo);
        });
        
        // Application.ApiEvent.AddApiEventListener("SheetSelectionChange", (sh, target)=>{
        //     let rangeInfo = formatRangeInfo({
        //         address: target.Address(),
        //         rowCount: target.Rows.Count,
        //         columnCount: target.Columns.Count
        //     });
        //     $print("工作表事件", "工作表选择变更: " + sh.Name + " 范围: " + rangeInfo);
            
        //     // 实时显示当前选中范围
        //     let selectedRange = getSelectedRange();
        //     if (selectedRange) {
        //         $print("选中详情", "当前选中: " + selectedRange.address + 
        //               " 工作表: " + selectedRange.worksheet + 
        //               " 工作簿: " + selectedRange.workbook);
        //     }
        // });
        
        // Application.ApiEvent.AddApiEventListener("SheetFollowHyperlink", (sh, target)=>{
        //     let cellInfo = formatCellInfo({
        //         address: target.Address,
        //         row: target.Row,
        //         column: target.Column,
        //         value: target.Value
        //     });
        //     $print("工作表事件", "工作表超链接点击: " + sh.Name + " 单元格: " + cellInfo);
        // });
        
        // Application.ApiEvent.AddApiEventListener("SheetBeforeDoubleClick", (sh, target, cancel)=>{
        //     let cellInfo = formatCellInfo({
        //         address: target.Address,
        //         row: target.Row,
        //         column: target.Column,
        //         value: target.Value
        //     });
        //     $print("工作表事件", "工作表双击前: " + sh.Name + " 单元格: " + cellInfo + " 取消: " + cancel);
        // });
        
        // Application.ApiEvent.AddApiEventListener("SheetBeforeRightClick", (sh, target, cancel)=>{
        //     let cellInfo = formatCellInfo({
        //         address: target.Address,
        //         row: target.Row,
        //         column: target.Column,
        //         value: target.Value
        //     });
        //     $print("工作表事件", "工作表右键前: " + sh.Name + " 单元格: " + cellInfo + " 取消: " + cancel);
        // });
        
        // Application.ApiEvent.AddApiEventListener("SheetCalculate", (sh)=>{
        //     $print("工作表事件", "工作表计算: " + sh.Name);
        // });

        // // 窗口级别事件
        // Application.ApiEvent.AddApiEventListener("WindowActivate", (wb, wn)=>{
        //     $print("窗口事件", "窗口激活: " + wb.Name + " 窗口索引: " + wn.Index);
        // });
        
        // Application.ApiEvent.AddApiEventListener("WindowDeactivate", (wb, wn)=>{
        //     $print("窗口事件", "窗口失活: " + wb.Name + " 窗口索引: " + wn.Index);
        // });
        
        // Application.ApiEvent.AddApiEventListener("WindowResize", (wb, wn)=>{
        //     $print("窗口事件", "窗口调整大小: " + wb.Name + " 窗口索引: " + wn.Index);
        // });

        // 文档级别事件
        Application.ApiEvent.AddApiEventListener("FileAfterSave", (doc)=>{
            $print("文档事件", "文件保存后: " + doc.Name);
        });
        
        // Application.ApiEvent.AddApiEventListener("DocumentAfterPrint", (doc)=>{
        //     $print("文档事件", "文档打印后: " + doc.Name);
        // });

        // // 系统事件
        // Application.ApiEvent.AddApiEventListener("AfterCalculate", ()=>{
        //     $print("系统事件", "计算完成");
        // });
        
        Application.ApiEvent.AddApiEventListener("AfterTaskPaneShow", ()=>{
            $print("系统事件", "任务窗格显示");
        });
        
        Application.ApiEvent.AddApiEventListener("AfterTaskPaneHidden", ()=>{
            $print("系统事件", "任务窗格隐藏");
        });

        $print("事件监听器", "已成功注册所有WPS表格事件监听器");
    } catch (ex) {
        $print("事件监听器错误", "注册事件监听器时发生异常：" + ex.message);
    }
}

// 页面加载完成后初始化事件监听器
window.onload = () => {
    $print("事件", "监听初始化");
    initApiEventListeners();
};

// 导出函数供其他模块使用
window.WPSEventUtils = {
    getSelectedRange: getSelectedRange,
    getActiveCell: getActiveCell,
    formatRangeInfo: formatRangeInfo,
    formatCellInfo: formatCellInfo
};