function onbuttonclick_erpMBOM(idStr) {
    try {
        if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
            window.Application.Enum = WPS_Enum;
        }
        switch (idStr) {
            case "dockLeft":
                $print('dockLeft start');
                let tsIdLeft = window.Application.PluginStorage.getItem("btnERPMBOM_id");
                if (tsIdLeft) {
                    let tskpane = window.Application.GetTaskPane(tsIdLeft);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionLeft;
                }
                break;
            case "dockRight":
                $print('dockRight start');
                let tsIdRight = window.Application.PluginStorage.getItem("btnERPMBOM_id");
                if (tsIdRight) {
                    let tskpane = window.Application.GetTaskPane(tsIdRight);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight;
                }
                break;
            case "hideTaskPane":
                $print('hideTaskPane start');
                let tsIdHide = window.Application.PluginStorage.getItem("btnERPMBOM_id");
                if (tsIdHide) {
                    let tskpane = window.Application.GetTaskPane(tsIdHide);
                    tskpane.Visible = false;
                }
                break;
            case "refresh":{
                let doc = window.Application.ActiveWorkbook
                let textValue = "";
                if (!doc){
                    textValue = textValue + "当前没有打开任何文档"
                    return
                }
                textValue = textValue + doc.Name
                document.getElementById("isMBOM").innerHTML = textValue
                break
            }

            case "createAll":{
                createWorkBook("createAll");
                break
            }
            case "createMBOM":{
                createWorkBook("createMBOM");
                break
            }
            case "createZC":{
                createWorkBook("createZC");
                break
            }
            case "createGY":{
                createWorkBook("createGY");
                break
            }
            case "createGX":{
                createWorkBook("createGX");
                break
            }
            case "checkError":{
                checkError();
                break
            }
            case "syncData":{
                syncData();
                break
            }
        }
    } catch (ex) {
        $print("错误", ex.message);
        alert(ex.message);
    }
}

window.onload = () => {
    $print('start_ptPlan_js_onload');
};

function createWorkBook(type){
    $print("creatWorkBook_start:"+type);
    try{
        let doc = window.Application.ActiveWorkbook;
        let textValue = "";
        if (!doc) {
            textValue = "当前没有打开任何文档";
        } else {
            textValue = doc.Name;
        }
        if ('MBOM数据.xlsm' != textValue) {
            alert("当前表不可用，请打开“MBOM数据.xlsm”重试");
            return;
        }
        let haveModelType1 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP物料主数据') {
                haveModelType1 = true;
                break;
            }
        }
        let haveModelType2 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP平层BOM') {
                haveModelType2 = true;
                break;
            }
        }
        let haveModelType3 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP工艺路线') {
                haveModelType3 = true;
                break;
            }
        }
        let haveModelType4 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP领料清单汇总') {
                haveModelType4 = true;
                break;
            }
        }



        
        if (!haveModelType1) {
            alert('没有表“ERP物料主数据”，请检查');
            return;
        }
        if (!haveModelType2) {
            alert('没有表“ERP平层BOM”，请检查');
            return;
        }
        if (!haveModelType3) {
            alert('没有表“ERP工艺路线”，请检查');
            return;
        }
        if (!haveModelType4) {
            alert('没有表“ERP领料清单汇总”，请检查');
            return;
        }
        if($SheetTheActive().Name != 'ERP物料主数据'){
            alert('请选择“ERP物料主数据”表中B列为-BOM的有效物料代号');
            return;
        }

    //检查选中的单元格是否为B列，且不为空值,且行号不可为1
        let selectedRange = $RangeSelection();
        
        if (selectedRange.Column != 2) {
            alert('请选中B列为-BOM的有效物料代号');
            return;
        }
        if (selectedRange.Row == 1) {
            alert('请选中B列为-BOM的有效物料代号');
            return;
        }
        if (selectedRange.Value() == null||selectedRange.Value()  == '') {
            alert('请选中B列为-BOM的有效物料代号');
            return;
        }
        // 操作类型与描述的映射表
        const actionMap = {
            'createAll': 'MBOM主材工艺工序领料',
            'createMBOM': 'MBOM',
            'createZC': '主材定额',
            'createGY': '工艺路线',
            'createGX': '工序领料'
        };

        //选择的图号需为-BOM结尾，符合格式：XX(X)-BOM
        const value = selectedRange.Value();
        const isBOMFormat = /^[^()]+\([^()]+\)-BOM$/.test(value); // 正则匹配：非括号字符+(非括号字符)-BOM
        // if (!isBOMFormat) {
        //     alert('请选中B列为-BOM的有效物料代号');
        //     return;
        // }


        var part = $RangeValue(selectedRange);
        // 获取ERP平层BOM表
        let erpBOMSheet = Application.Worksheets.Item('ERP平层BOM');
        let bomUsedRange = erpBOMSheet.Range("A2:AB"+$SheetsLastRowNum('ERP平层BOM',4));
        let bomData = bomUsedRange.Value2;
        // $print($ArrayToString(bomData));

        let bomRows = $Array2DRowColCount(bomData).rows;
        let erpMaterialData = Application.Worksheets.Item('ERP物料主数据').Range("A2:T"+$SheetsLastRowNum('ERP物料主数据',2)).Value2;
        let erpMaterialRows = $Array2DRowColCount(erpMaterialData).rows;
        // 获取ERP工艺路线表数据
        let erpProcessRouteSheet = Application.Worksheets.Item('ERP工艺路线');
        let processRouteLastRow = $SheetsLastRowNum('ERP工艺路线', 1);

        // 获取ERP领料清单汇总表数据
        let erpMaterialListSheet = Application.Worksheets.Item('ERP领料清单汇总');
        let materialListLastRow = $SheetsLastRowNum('ERP领料清单汇总', 1);
        // 创建新workbook

        var workbookName = $StringReplaceAll(part,'/','_')+"-"+actionMap[type]+"-"+new Date().valueOf();
        let newWorkbook = Application.Workbooks.Add();
        window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                
        let newSheet = newWorkbook.Worksheets.Item(1);
        newSheet.Name = 'BOM导入';
        // 写入新表头（A列到AC列）
        let newHeader = ['说明','替代物料编码','父级物料编码','父级物料代号','父级版本','子件物料编码','子件物料代号','子件物料名称','设计版本',
            '物料类别','用量','计划用量','型号','是否辅料','是否组别件','组别号','是否预装件','是否选装件','主制部门','是否设计虚拟','是否工艺虚拟',
            '是否多工艺状态','损耗率','顶层BOM','BOM类型（PBOM/WBOM/TBOM）','是否必换件（大修）','是否易损件（大修）','易损率（大修）','修理车间',
            '大修物料编码','大修物料代号'];
        newSheet.Range('A1', 'AE1').Value2 = newHeader;
        for (let i = 2; i <= bomRows-1; i++) {
            if (bomData[i-2][3] === part) {
                // 写入匹配行到新表
                newSheet.Range(`A2`, `AE2`).Value2 = [
                    "", bomData[i-2][1], bomData[i-2][17], bomData[i-2][2], bomData[i-2][18], bomData[i-2][19], 
                    bomData[i-2][3], bomData[i-2][4], bomData[i-2][20], bomData[i-2][5], bomData[i-2][6], bomData[i-2][22], 
                    bomData[i-2][7], bomData[i-2][8], bomData[i-2][9], bomData[i-2][10], bomData[i-2][11], bomData[i-2][12], 
                    bomData[i-2][13], bomData[i-2][14], bomData[i-2][15], bomData[i-2][16], bomData[i-2][21], 
                    part, 'PBOM', '', '', '', '', '', ''
                ];
                break;
            }
        }
        //定义递归函数处理BOM展开
        function processBOM(parentPart, currentRow) {
            // 查找D列（物料代号）匹配的行
            for (let i = 2; i <= bomRows; i++) {
                if (bomData[i-1][2] === parentPart) {
                    // 写入匹配行到新表
                    let targetRow = currentRow;
                    newSheet.Range(`A${targetRow}`, `AC${targetRow}`).Value2 = [
                        '', bomData[i-1][1], bomData[i-1][17], bomData[i-1][2], bomData[i-1][18], bomData[i-1][19],
                         bomData[i-1][3], bomData[i-1][4], bomData[i-1][20], bomData[i-1][5], bomData[i-1][6], bomData[i-1][22], 
                         bomData[i-1][7], bomData[i-1][8], bomData[i-1][9], bomData[i-1][10], bomData[i-1][11], bomData[i-1][12], 
                         bomData[i-1][13], bomData[i-1][14], bomData[i-1][15], bomData[i-1][16], bomData[i-1][21], 
                         part, 'PBOM', '', '', '', ''
                    ];
                    currentRow++;
                    // 递归查找C列（父级物料代号）匹配的行
                    currentRow = processBOM(bomData[i-1][3], currentRow);
                }
            }
            return currentRow;
        }

        // 从初始part开始处理
        processBOM(part, 3);
        let mbomData = newSheet.Range(`A2`, `AC`+$SheetsLastRowNum('BOM导入',4)).Value2;
        let mbomRows = $Array2DRowColCount(mbomData).rows;


        $SheetsLastAdd('主材导入');
        let newSheet2 = newWorkbook.Worksheets.Item('主材导入');
        // 写入新表头（用户指定格式）
        let newHeader2 = ['主制部门', '零件代号', '零件名称', '材料代号', '材料名称', '单件定额', '毛料尺寸', '单件毛料尺寸', '一根毛料可做零件数', '零件编码', '材料编码'];
        newSheet2.Range('A1', 'K1').Value2 = newHeader2; // 调整列范围为K列（11列）
        currentRow2 = 2
        for (let i = 0; i < mbomRows; i++) {
            //检查如果已经有相应的零件代号的记录则忽略
            let isExist = false;
            for (let j = 2; j <= currentRow2; j++) {
                if (newSheet2.Range(`B${j}`).Value() == mbomData[i][6]) {
                    isExist = true;
                    break;
                }
            }
            if (isExist) {
                continue;
            }
            for (let j = 0; j < erpMaterialRows; j++) {
                if (erpMaterialData[j][0] == mbomData[i][5]&&erpMaterialData[j][12]!=null&&erpMaterialData[j][12]!="") {
                    // 查找ERP物料主数据中匹配的物料信息（复用之前定义的erpMaterialData）
                    newSheet2.Range(`A${currentRow2}`, `K${currentRow2}`).Value2 = [
                        erpMaterialData[j][6],  // 主制部门
                        erpMaterialData[j][1],  // 零件代号
                        erpMaterialData[j][2],  // 零件名称
                        erpMaterialData[j][12], // 材料代号
                        erpMaterialData[j][13], // 材料名称
                        erpMaterialData[j][14], // 单件定额
                        erpMaterialData[j][15], // 毛料尺寸
                        erpMaterialData[j][16], // 单件毛料尺寸
                        erpMaterialData[j][17], // 一根毛料可做零件数
                        erpMaterialData[j][18], // 零件编码
                        erpMaterialData[j][19]  // 材料编码
                    ];
                    currentRow2++;
                    break;
                }
            }
        }
           
        $SheetsLastAdd('工艺路线导入');   
        let newSheet3 = newWorkbook.Worksheets.Item('工艺路线导入');
        // 写入用户指定表头
        let newHeader3 = ['物料代号', '物料编码', '设计版本', '产品型号', '工艺类型', '工艺版本', '主制车间', '工序号', '工序名称', '工序版本', '工作中心', '是否领料工序', '是否可临时外协', '是否军检', '是否关键工序', '是否重要工序', '是否检验工序', '是否特殊过程记录', '准备工时_分钟', '单件加工工时_分钟', '工序车间', '工序合格率', '等待工时(分钟)'];
        newSheet3.Range('A1', 'W1').Value2 = newHeader3; // W列对应23个字段

        
        let processRange = erpProcessRouteSheet.Range(`A2:AA${processRouteLastRow}`);
        let erpProcessRouteData = processRange.Value2;
        let erpProcessRows = $Array2DRowColCount(erpProcessRouteData).rows;
        let currentRow3 = 2;
        for (let i = 0; i < mbomRows; i++) {
            //检查如果已经有相应的零件代号的记录则忽略
            let isExist = false;
            for (let k = 2; k < currentRow3; k++) {
                if (newSheet3.Range(`A${k}`).Value() == mbomData[i][6]) {
                    isExist = true;
                    break;
                }
            }
            if (isExist) {
                continue;
            }
            for (let j = 0; j < erpProcessRows; j++) {
                if (erpProcessRouteData[j][1] == mbomData[i][6]) {
                    // 查找ERP工艺路线表中匹配的物料代号
                    
                    newSheet3.Range(`A${currentRow3}`, `W${currentRow3}`).Value2 = [
                        erpProcessRouteData[j][1],   // 物料代号（ERP表第2列）
                        erpProcessRouteData[j][2],   // 物料编码（ERP表第3列）
                        erpProcessRouteData[j][3],   // 设计版本（ERP表第4列）
                        erpProcessRouteData[j][4],   // 产品型号（ERP表第5列）
                        '新机工艺',                        // 工艺类型（批产为P）
                        erpProcessRouteData[j][5],   // 工艺版本（ERP表第6列）
                        erpProcessRouteData[j][6],   // 主制车间（ERP表第7列）
                        erpProcessRouteData[j][7],   // 工序号（ERP表第8列）
                        erpProcessRouteData[j][8],   // 工序名称（ERP表第9列）
                        erpProcessRouteData[j][9],   // 工序版本（ERP表第10列）
                        erpProcessRouteData[j][10],  // 工作中心（ERP表第11列）
                        erpProcessRouteData[j][11],  // 是否领料工序（ERP表第12列）
                        erpProcessRouteData[j][12],  // 是否可临时外协（ERP表第13列）
                        erpProcessRouteData[j][13],  // 是否军检（ERP表第14列）
                        erpProcessRouteData[j][14],  // 是否关键工序（ERP表第15列）
                        erpProcessRouteData[j][15],  // 是否重要工序（ERP表第16列）
                        erpProcessRouteData[j][16],  // 是否检验工序（ERP表第17列）
                        erpProcessRouteData[j][17],  // 是否特殊过程记录（ERP表第18列）
                        erpProcessRouteData[j][18],  // 准备工时_分钟（ERP表第19列）
                        erpProcessRouteData[j][19],  // 单件加工工时_分钟（ERP表第20列）
                        erpProcessRouteData[j][20],  // 工序车间（ERP表第21列）
                        erpProcessRouteData[j][21],  // 工序合格率（ERP表第22列）
                        erpProcessRouteData[j][22]   // 等待工时(分钟)（ERP表第23列）
                    ];
                    currentRow3++;
                }
            }
        }
         
        $SheetsLastAdd('工序领料导入');
        let newSheet4 = newWorkbook.Worksheets.Item('工序领料导入');
        // 写入用户指定表头
        let newHeader4 = ['工艺物料编号', '工艺物料代号', '工艺物料名称', '工艺类型(W(大修)/p(批产)', '工艺路线版本', '工序编号', '工序名称', '物料编号', '物料代号', '物料名称', '用量', '物料设计版本'];
        newSheet4.Range('A1', 'L1').Value2 = newHeader4; // L列对应12个字段

        
        let materialListUsedRange = erpMaterialListSheet.Range(`A2:N${materialListLastRow}`);
        let erpMaterialListData = materialListUsedRange.Value2;
        let erpMaterialListRows = $Array2DRowColCount(erpMaterialListData).rows;
        let currentRow4 = 2;
        for (let i = 0; i < mbomRows; i++) {
            //检查如果已经有相应的零件代号的记录则忽略
            let isExist = false;
            for (let k = 2; k < currentRow4; k++) {
                if (newSheet4.Range(`B${k}`).Value() == mbomData[i][6]) {
                    isExist = true;
                    break;
                }
            }
            if (isExist) {
                continue;
            }
            for (let j = 0; j < erpMaterialListRows; j++) {

                if (erpMaterialListData[j][0] == mbomData[i][6]) {
                    // 写入用户指定表头数据（工艺类型固定为P）
                    newSheet4.Range(`A${currentRow4}`, `L${currentRow4}`).Value2 = [
                        erpMaterialListData[j][1],  // 工艺物料编号（父级编码）
                        erpMaterialListData[j][0],  // 工艺物料代号（父级物料代号）
                        erpMaterialListData[j][10], // 工艺物料名称（父级名称）
                        '新机工艺',                     // 工艺类型（固定为P）
                        erpMaterialListData[j][3],  // 工艺路线版本（父级版本）
                        erpMaterialListData[j][4],  // 工序编号（工序号）
                        erpMaterialListData[j][5],  // 工序名称
                        erpMaterialListData[j][8],  // 物料编号（子件物料编码）
                        erpMaterialListData[j][7],  // 物料代号（子件物料代号）
                        erpMaterialListData[j][11], // 物料名称（子件名称）
                        erpMaterialListData[j][9],  // 用量
                        erpMaterialListData[j][2]   // 物料设计版本（设计版本）
                    ];
                    currentRow4++;
                }
            }
        }
                     
        
        // 配置type与保留工作表的映射关系（键为type值，值为需要保留的工作表数组）
        const typeSheetMap = {
            createAll: [], // 保留4个表
            createMBOM: ['主材导入', '工艺路线导入', '工序领料导入'], // 只保留BOM导入表
            createZC: ['BOM导入', '工艺路线导入', '工序领料导入'], // 只保留主材导入表
            createGY: ['BOM导入', '主材导入', '工序领料导入'], // 只保留工艺路线导入表
            createGX: ['BOM导入', '主材导入', '工艺路线导入',]// 只保留工序导入表
        };
        // 根据当前type获取需要保留的工作表列表
        const keepSheets = typeSheetMap[type] || [];
        // 隐藏所有工作表
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            Application.Worksheets.Item(i).Visible = true;
        }

        // 显示需要保留的工作表
        for (let sheetName of keepSheets) {
            let sheet = newWorkbook.Worksheets.Item(sheetName);
            if (sheet) {
                sheet.Visible = false;
            }
        }

        return;
    }catch(ex){
        alert(ex.message);
        $print(ex.message);
    }
}

function checkError(){
    try{
        let doc = window.Application.ActiveWorkbook;
        let textValue = "";
        if (!doc) {
            textValue = "当前没有打开任何文档";
        } else {
            textValue = doc.Name;
        }
        if ('MBOM数据.xlsm' != textValue) {
            alert("当前表不可用，请打开“MBOM数据.xlsm”重试");
            return;
        }
        let haveModelType1 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP物料主数据') {
                haveModelType1 = true;
                break;
            }
        }
        let haveModelType2 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP平层BOM') {
                haveModelType2 = true;
                break;
            }
        }
        let haveModelType3 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP工艺路线') {
                haveModelType3 = true;
                break;
            }
        }
        let haveModelType4 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP领料清单汇总') {
                haveModelType4 = true;
                break;
            }
        }
        if (!haveModelType1) {
            alert('没有表“ERP物料主数据”，请检查');
            return;
        }
        if (!haveModelType2) {
            alert('没有表“ERP平层BOM”，请检查');
            return;
        }
        if (!haveModelType3) {
            alert('没有表“ERP工艺路线”，请检查');
            return;
        }
        if (!haveModelType4) {
            alert('没有表“ERP领料清单汇总”，请检查');
            return;
        }

        //检查以上四个表中，所有名字结尾为“校验”的列，查找内容是否为“/“，如果不是，则选中该表的该单元格
        const targetSheets = ['ERP物料主数据', 'ERP平层BOM', 'ERP工艺路线', 'ERP领料清单汇总']; 
        var haveError = false;
        for (const sheetName of targetSheets) { 
            const sheet = Application.Worksheets.Item(sheetName); 
            const usedRange = sheet.UsedRange; 
            const rowCount = usedRange.Rows.Count; 
            const colCount = usedRange.Columns.Count; 
            // 遍历第一行（表头）找结尾为“校验”的列 
            for (let col = 1; col <= colCount; col++) { 
                const header = sheet.Cells.Item(1, col).Value(); 
                if (header && header.endsWith('校验')) { 
                    // 遍历该列数据行（从第2行开始） 
                    for (let row = 2; row <= rowCount; row++) { 
                        const cellValue = sheet.Cells.Item(row, col).Value(); 
                        if (cellValue !== '/'&&cellValue !== ''&&cellValue != null) { 
                            $SheetsActivate(sheetName);
                            sheet.Cells.Item(row, col).Select(); 
                            haveError = true;
                        } 
                    } 
                } 
            } 
        } 
        if (!haveError) {
            alert('未发现错误');
            return;
        }




    }catch(ex){
        alert(ex.message);
        $print(ex.message);
    }
}

function syncData(){
    try{
        let doc = window.Application.ActiveWorkbook;
        let textValue = "";
        if (!doc) {
            textValue = "当前没有打开任何文档";
        } else {
            textValue = doc.Name;
        }
        if ('MBOM数据.xlsm' != textValue) {
            alert("当前表不可用，请打开“MBOM数据.xlsm”重试");
            return;
        }
        //询问是否确认同步数据，确认继续，取消则退出
        let confirmSync = confirm('确认同步数据吗？');
        if (!confirmSync) {
            alert('取消同步数据');
            return;
        }
        let haveModelType1 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP物料主数据') {
                haveModelType1 = true;
                break;
            }
        }
        let haveModelType2 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP平层BOM') {
                haveModelType2 = true;
                break;
            }
        }
        let haveModelType3 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP工艺路线') {
                haveModelType3 = true;
                break;
            }
        }
        let haveModelType4 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == 'ERP领料清单汇总') {
                haveModelType4 = true;
                break;
            }
        }
        let haveModelType5 = false;
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == '物料参数') {
                haveModelType5 = true;
                break;
            }
        }
        if (!haveModelType1) {
            alert('没有表“ERP物料主数据”，请检查');
            return;
        }
        if (!haveModelType2) {
            alert('没有表“ERP平层BOM”，请检查');
            return;
        }
        if (!haveModelType3) {
            alert('没有表“ERP工艺路线”，请检查');
            return;
        }
        if (!haveModelType4) {
            alert('没有表“ERP领料清单汇总”，请检查');
            return;
        }
        if (!haveModelType5) { 
            alert('没有表“物料参数”，请检查'); 
            return; 
        } 
        




    }catch(ex){
        alert(ex.message);
        $print(ex.message);
    }
}