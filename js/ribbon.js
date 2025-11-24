
//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI){
    if (typeof (window.Application.ribbonUI) != "object"){
		window.Application.ribbonUI = ribbonUI
    }
    
    if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
        window.Application.Enum = WPS_Enum
    }

    window.Application.PluginStorage.setItem("EnableFlag", false) //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    window.Application.PluginStorage.setItem("ApiEventFlag", false) //往PluginStorage中设置一个标记，用于控制ApiEvent的按钮label
    return true
}

var WebNotifycount = 0;
function OnAction(control) {
    try {
        const eleId = control.Id
        switch (eleId) {
            case "btnShowMsg":
                {
                    const doc = window.Application.ActiveWorkbook
                    if (!doc) {
                        alert("当前没有打开任何文档")
                        return
                    }
                    alert(doc.Name)
                }
                break;
            case "btnIsEnbable":
                {
                    let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                    window.Application.PluginStorage.setItem("EnableFlag", !bFlag)
                    
                    //通知wps刷新以下几个按饰的状态
                    window.Application.ribbonUI.InvalidateControl("btnIsEnbable")
                    window.Application.ribbonUI.InvalidateControl("btnShowDialog") 
                    window.Application.ribbonUI.InvalidateControl("btnShowTaskPane") 
                    window.Application.ribbonUI.InvalidateControl("btnTest") 
                    //window.Application.ribbonUI.Invalidate(); 这行代码打开则是刷新所有的按钮状态
                    break
                }
            case "btnShowDialog":
                window.Application.ShowDialog(GetUrlPath() + "/ui/dialog.html", "这是一个对话框网页", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false)
                break
            case "btnShowTaskPane":
                {
                    let tsId = window.Application.PluginStorage.getItem("taskpane_id")
                    if (!tsId) {
                        let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                        let id = tskpane.ID
                        window.Application.PluginStorage.setItem("taskpane_id", id)
                        tskpane.Visible = true
                    } else {
                        let tskpane = window.Application.GetTaskPane(tsId)
                        tskpane.Visible = !tskpane.Visible
                    }
                }
                break
            case "btnApiEvent":
                {
                    let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag")
                    let bRegister = bFlag ? false : true
                    window.Application.PluginStorage.setItem("ApiEventFlag", bRegister)
                    if (bRegister){
                        window.Application.ApiEvent.AddApiEventListener('NewWorkbook', OnNewDocumentApiEvent)
                    }
                    else{
                        window.Application.ApiEvent.RemoveApiEventListener('NewWorkbook', OnNewDocumentApiEvent)
                    }

                    window.Application.ribbonUI.InvalidateControl("btnApiEvent") 
                }
                break
            case "btnWebNotify":
                {
                    let currentTime = new Date()
                    let timeStr = currentTime.getHours() + ':' + currentTime.getMinutes() + ":" + currentTime.getSeconds()
                    window.Application.OAAssist.WebNotify("这行内容由wps加载项主动送达给业务系统，可以任意自定义, 比如时间值:" + timeStr + "，次数：" + (++WebNotifycount), true)
                }
                break
            case "btnTest":
                {
                    const doc = window.Application.ActiveWorkbook;
                    alert('李文玉测试用');
                    alert($ArrayToString($RangeSelection().Value2));
                }
                break
            case "btnSelectFilter":
                {
                    const doc = window.Application.ActiveWorkbook;
                    const activeSheet = doc.ActiveSheet;
                    try{
                        // 获取选中的值
                        let selectedValue = $RangeSelection().Value2;
                        
                        // 检查剪贴板是否有值
                        if (!selectedValue) {
                            alert('请先选择一个单元格');
                            return;
                        }
                
                        $RangeSelection().AutoFilter($RangeSelection().Column, selectedValue, null, null, true);
                        
                    } catch (error) {
                        alert('筛选过程中发生错误：' + error.message);
                        $print(error.message);
                    }
                }
                break
            case "btnSelectFilterInit":
                {
                    const doc = window.Application.ActiveWorkbook;
                    const activeSheet = doc.ActiveSheet;
                    try{
                        
                        $RangeSelection().AutoFilter(null, null, null, null, true);
                        
                    } catch (error) {
                        alert('筛选过程中发生错误：' + error.message);
                        $print(error.message);
                    }
                }
                break
            case "btnOpenHKExcel": 
                {
                    const filePath = "\\\\mesdb\\mes0\\航科EXCEL1.0.xlsx"; // 转义反斜杠以适配JavaScript字符串
                    try {
                        // 使用WPS API打开指定路径的工作簿
                        const workbook = window.Application.Workbooks.Open(filePath);
                        if (workbook) {
                            alert("文档已成功打开：" + workbook.Name);
                        } else {
                            alert("打开文档失败，请检查路径是否正确。");
                        }
                    }catch (error) {
                        alert("打开文档时发生错误：" + error.message);
                    }
                    
                }
                break       
            case "btnOpenLog": 
                {
                    let tempPath = Application.Env.GetHomePath();
                    if(Application.FileSystem.existsSync(tempPath+'/wpslog')){
                        //输入年月日查询日志 - 使用WPS原生InputBox替代window.prompt
                        try {
                            // 优先尝试使用封装的$InputBox函数
                            if (typeof $InputBox === 'function') {
                                var dateStr = $InputBox("请输入日期（格式：yyyyMMdd）：", "日期输入", "");
                            } else {
                                // 直接使用Application.InputBox
                                var dateStr = Application.InputBox("请输入日期（格式：yyyyMMdd）：", "日期输入", "");
                            }
                            
                            if (!dateStr || dateStr === false) {
                                // 在WPS中，取消操作可能返回false
                                return;
                            }
                        } catch (e) {
                            // 降级处理：如果InputBox也失败，使用简单的确认对话框引导用户
                            alert("无法显示输入框。请手动在日志文件名中输入日期。");
                            return;
                        }
                        const filePath = tempPath+'/wpslog/'+dateStr+'.txt';
                        if(Application.FileSystem.Exists(filePath)){
                            try {
                                // 使用WPS API打开指定路径的工作簿
                                const workbook = window.Application.Workbooks.Open(filePath);
                                if (workbook) {
                                    alert("文档已成功打开：" + workbook.Name);
                                } else {
                                    alert("打开文档失败，请检查路径是否正确。");
                                }
                            }catch (error) {
                                alert("打开文档时发生错误：" + error.message);
                            }
                        }else{
                            alert("文件不存在，请检查日期是否输入正确。");
                        }
                    }
                }
                break       
            // case "btnOpenERPceshi": 
            //     {
            //         window.Application.ShowDialog("http://20.19.18.18:8088/user/login", "ERP测试系统", 1280 * window.devicePixelRatio, 720 * window.devicePixelRatio, false)
            //     }
            //     break  
            case "btnDeleteSheet": {
                    // 获取当前工作簿和工作表
                const activeWorkbook = window.Application.ActiveWorkbook;
                const activeSheet = activeWorkbook.ActiveSheet;
                if (!activeSheet) {
                    alert("当前无活动工作表");
                    return;
                }
                $SheetDelOrSign(activeSheet.Name);
                
            }
                break
            case "btnCreateSOMSheet":{
                let tsId = window.Application.PluginStorage.getItem("btnCreateSOMSheet_id")
                if (!tsId) {
                    let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/pp/btnCreateSOMSheet.html")
                    let id = tskpane.ID
                    window.Application.PluginStorage.setItem("btnCreateSOMSheet_id", id)
                    tskpane.Visible = true
                } else {
                    let tskpane = window.Application.GetTaskPane(tsId)
                    tskpane.Visible = !tskpane.Visible

                }
            }
                break
            case "btnCreateInBatch":{btnCreateInBatchMain();}
                break
            case "btnCreateInBatchHaveStock":{btnCreateInBatchHaveStockMain();}
                break
            case "btnStockParts":{btnStockPartsMain();}
                break
            case "btnPTPlan":
                {
                    let tsId = window.Application.PluginStorage.getItem("ptPlan_id")
                    if (!tsId) {
                        let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/pp/ptPlan.html")
                        let id = tskpane.ID
                        window.Application.PluginStorage.setItem("ptPlan_id", id)
                        tskpane.Visible = true
                    } else {
                        let tskpane = window.Application.GetTaskPane(tsId)
                        tskpane.Visible = !tskpane.Visible
                        // window.Application.PluginStorage.removeItem("ptPlan_id");
                        // tskpane.Delete();
                    }
                }
                break
                case "btnCPInStock":
                    {
                        let tsId = window.Application.PluginStorage.getItem("btnCPInStock_id")
                        if (!tsId) {
                            let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/pp/btnCPInStock.html")
                            let id = tskpane.ID
                            window.Application.PluginStorage.setItem("btnCPInStock_id", id)
                            tskpane.Visible = true
                        } else {
                            let tskpane = window.Application.GetTaskPane(tsId)
                            tskpane.Visible = !tskpane.Visible
                        }
                    }
                    break
                // case "btnOpenPlanOld":
                //     {
                //         window.Application.ShowDialog("http://192.168.70.26/", "生产计划系统", 1280 * window.devicePixelRatio, 720 * window.devicePixelRatio, false)
                //     }
                //     break
                case "btnOpenPlan":
                    {
                        window.Application.ShowDialog("http://192.168.70.26:8080/V6R343", "计划辅助系统", 1280 * window.devicePixelRatio, 720 * window.devicePixelRatio, false)
                    }
                    break
                case "btnOpenWPSJSA":
                    {
                        // 打开我们创建的WPS JSA文档浏览器
                        const jsaDocPath = GetUrlPath() + "/standard_WPSJSA_EXCEL/standalone-viewer.html";
                        // 使用合适的页面大小：1200x800
                        window.Application.ShowDialog(jsaDocPath, "WPS JSA API 文档浏览器", 1600 * window.devicePixelRatio, 900 * window.devicePixelRatio, false);
                    }
                    break;
                    // case "btnOpenWPSJSA2":
                    //     {
                    //         // 打开我们创建的WPS JSA文档浏览器
                    //         const jsaDocPath = GetUrlPath() + "/standard_WPSJSA_EXCEL/viewer/index.html";
                    //         // 使用合适的页面大小：1200x800
                    //         window.Application.ShowDialog(jsaDocPath, "WPS JSA API 文档浏览器", 1600 * window.devicePixelRatio, 900 * window.devicePixelRatio, false);
                    //     }
                    //     break
                    // case "btnOpenERP":
                    //     {
                    //         window.Application.ShowDialog("http://20.19.4.13:8080/user/login", "ERP系统", 1280 * window.devicePixelRatio, 720 * window.devicePixelRatio, false)
                    //     }
                    //     break
                    case "btnServer":
                    {
                        // 显示当前服务器信息
                        let currentConfig = loadServerConfig();
                        let serverUrl = currentConfig.PRODUCTION;
                        let serverName = '主服务器';
                        let serverInfo = serverName + '：' + serverUrl;
                        
                        alert('当前服务器配置信息：\n\n' + 
                              '服务器类型：' + serverName + '\n' +
                              '服务器地址：' + serverUrl + '\n\n' +
                              '使用正式环境配置');
                    }
                    break
                case "btnGPTmain1":
                    {
                        let tsId = window.Application.PluginStorage.getItem("btnGPTmain0_id")
                        if (!tsId) {
                            let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/pp/btnGPTmain0.html")
                            let id = tskpane.ID
                            window.Application.PluginStorage.setItem("btnGPTmain0_id", id)
                            tskpane.Visible = true
                        } else {
                            let tskpane = window.Application.GetTaskPane(tsId)
                            tskpane.Visible = !tskpane.Visible
                        }
                    }
                break
                case "btnGPTmain2":
                    {
                        window.Application.ShowDialog(GetUrlPath() + "/ui/pp/btnGPTmain1.html", "AI机器人", 960 * window.devicePixelRatio, 720 * window.devicePixelRatio, false)
                    }
                    break;
                case "btnAIFormulaGen":
                    {
                        // 直接打开AI公式生成器
                        let tsId = window.Application.PluginStorage.getItem("btnAIFormulaGen_id")
                        if (!tsId) {
                            let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/aiHelper/aiHelper.html")
                            let id = tskpane.ID
                            window.Application.PluginStorage.setItem("btnAIFormulaGen_id", id)
                            tskpane.Visible = true
                        } else {
                            let tskpane = window.Application.GetTaskPane(tsId)
                            tskpane.Visible = !tskpane.Visible
                        }
                    }
                    break;
                case "btnERPMBOM":
                    {
                        let tsId = window.Application.PluginStorage.getItem("btnERPMBOM_id")
                        if (!tsId) {
                            let tskpane = window.Application.CreateTaskPane(GetUrlPath() + "/ui/pp/mbom/erpMBOM.html")
                            let id = tskpane.ID
                            window.Application.PluginStorage.setItem("btnERPMBOM_id", id)
                            tskpane.Visible = true
                        } else {
                            let tskpane = window.Application.GetTaskPane(tsId)
                            tskpane.Visible = !tskpane.Visible
                        }
                    }
                    break   
                case "btnRedFill":
                    onRedFill();
                    break;
                case "btnYellowFill":
                    onYellowFill();
                    break;
                case "btnGreenFill":
                    onGreenFill();
                    break;
                case "btnSkyBlueFill":
                    onSkyBlueFill();
                    break;
                case "btnLightBlueFill":
                    onLightBlueFill();
                    break;
                case "btnNoFill":
                    onNoFill();
                    break;        
            default:
                break
        }
        return true
    } catch (ex) {
        alert('错误:', ex);
        $print("错误", ex);
        return [];
    }
}

function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        case "btnTest":
            return "images/test.svg"
        case "btnCreateSOMSheet":
            return "images/pp/shortageOfMaterials/btnCreateSOM2.svg"
        case "btnCreateInBatch":
            return "images/pp/shortageOfMaterials/btnCreateInBatch.svg"
        case "btnCreateInBatchHaveStock":
            return "images/pp/shortageOfMaterials/btnCreateInBatch.svg"
        case "btnStockParts":
            return "images/pp/shortageOfMaterials/btnStockParts.svg"
        case "btnPTPlan":
            return "images/pp/ptPlan.svg"
        case "btnCPInStock":
            return "images/pp/btnCPInStock.svg"
        case "btnGPTmain1":
            return "images/pp/btnGPTmain0.svg"
        case "btnGPTmain2":
            return "images/pp/btnGPTmain0.svg"
        case "btnAIFormulaGen":
            return "images/pp/btnGPTmain0.svg"
        case "btnRedFill":
            return "images/tool/color/red.svg"
        case "btnYellowFill":
            return "images/tool/color/yellow.svg"
        case "btnGreenFill":
            return "images/tool/color/green.svg"
        case "btnSkyBlueFill":
            return "images/tool/color/blue.svg"
        case "btnLightBlueFill":
            return "images/tool/color/blue2.svg"
        case "btnNoFill":
            return "images/tool/color/nofill.svg"
        case "btnERPMBOM":
            return "images/pp/mbom/mbomERP.svg"
        case "btnOpenWPSJSA":
            return "images/jsa.svg"
        // case "btnOpenWPSJSA2":
        //     return "images/jsa.svg"           
        default:
            return "images/newFromTemp.svg";
    }
    
}

function OnGetEnabled(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return true
            break
        case "btnShowDialog":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "btnShowTaskPane":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        case "btnTest":
            {
                let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
                return bFlag
                break
            }
        default:
            break
    }
    return true
}

function OnGetVisible(control){
    return true
}

function OnGetLabel(control){
    const eleId = control.Id
    switch (eleId) {
        case "btnIsEnbable":
        {
            let bFlag = window.Application.PluginStorage.getItem("EnableFlag")
            return bFlag ?  "按钮Disable" : "按钮Enable"
            break
        }
        case "btnApiEvent":
        {
            let bFlag = window.Application.PluginStorage.getItem("ApiEventFlag")
            return bFlag ? "清除新建文件事件" : "注册新建文件事件"
            break
        }    
    }
    return ""
}

function onRedFill() {
    try {
        // 使用与其他填充函数相同的API调用方式以保持一致性
        const selection = wps.EtApplication().ActiveSheet.Application.Selection;
        if (selection) {
            // 红色的正确值：在Excel/WPS中颜色使用BGR格式，红色应为0x0000FF（BGR格式）
            // 16711680是RGB的红色，但在WPS中需要BGR格式的红色即255
            selection.Interior.Color = 255; // 255是BGR格式的红色 (0x0000FF)
            // $print("已填充红色");
        } else {
            alert("请先选中单元格");
        }
    } catch (e) {
        $print("红色填充错误:", e.message);
        alert("填充红色失败: " + e.message);
    }
}

function onYellowFill() {
    const selection = wps.EtApplication().ActiveSheet.Application.Selection;
    if (selection) {
        selection.Interior.Color = 65535; // 黄色RGB值
        // $print("已填充黄色");
    } else {
        alert("请先选中单元格");
    }
}

function onGreenFill() {
    const selection = wps.EtApplication().ActiveSheet.Application.Selection;
    if (selection) {
        selection.Interior.Color = 5296274; // 绿色RGB值
        // $print("已填充绿色");
    } else {
        alert("请先选中单元格");
    }
}

function onNoFill() {
    const selection = wps.EtApplication().ActiveSheet.Application.Selection;
    if (selection) {
        selection.Interior.ColorIndex = -4142; // 使用ColorIndex设置无填充颜色
        // $print("已填充无色");
    } else {
        alert("请先选中单元格");
    }
}
function onSkyBlueFill() {
    const selection = wps.EtApplication().ActiveSheet.Application.Selection;
    if (selection) {
        selection.Interior.Color = 15543831; // 天蓝色#87CEEB对应的数值
        // $print("已填充天蓝色");
    } else {
        alert("请先选中单元格");
    }
}
function onLightBlueFill() {
    const selection = wps.EtApplication().ActiveSheet.Application.Selection;
    if (selection) {
        selection.Interior.Color = 15128749; // 浅蓝色#ADD8E6对应的数值
        // $print("已填充浅蓝色");
    } else {
        alert("请先选中单元格");
    }
}
function OnNewDocumentApiEvent(doc){
    alert("新建文件事件响应，取文件名: " + doc.Name)
}

//添加缓存清理和预加载相关辅助函数
// 简化的清理函数，兼容旧版WPS
function cleanupBeforeLoad() {
    // 清理可能的缓存和临时数据
    if (window._aiChatCache) {
        window._aiChatCache = null;
    }
    
    // 清理可能的会话状态
    if (window._aiSessionState) {
        window._aiSessionState = null;
    }
}

// 简化的预加载检查函数
function preloadCheck(url) {
    // 直接返回true，避免在旧版WPS中可能出现的兼容性问题
    return true;
}

// 简化版对话框打开函数，兼容旧版WPS
function openDialogWithTimeout(app, url, title, width, height) {
    try {
        app.ShowDialog(url, title, width, height, false);
    } catch (e) {
        alert("打开对话框失败，请手动访问: " + url);
    }
}
var ribbonUI = null;

