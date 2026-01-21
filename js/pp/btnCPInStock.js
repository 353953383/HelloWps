function onbuttonclick_btnCPInStock(idStr) {
    try {
        if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
            window.Application.Enum = WPS_Enum;
        }
        switch (idStr) {
            case "dockLeft":
                $print('dockLeft start');
                let tsIdLeft = window.Application.PluginStorage.getItem("btnCPInStock_id");
                if (tsIdLeft) {
                    let tskpane = window.Application.GetTaskPane(tsIdLeft);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionLeft;
                }
                break;
            case "dockRight":
                $print('dockRight start');
                let tsIdRight = window.Application.PluginStorage.getItem("btnCPInStock_id");
                if (tsIdRight) {
                    let tskpane = window.Application.GetTaskPane(tsIdRight);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight;
                }
                break;
            case "hideTaskPane":
                $print('hideTaskPane start');
                let tsIdHide = window.Application.PluginStorage.getItem("btnCPInStock_id");
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
                document.getElementById("docName").innerHTML = textValue
                break
            }
            case "modelIndex":{
                $print('modelIndex start');
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        Application.Worksheets.Item(i).Activate();
                        haveModelType = true;
                        break;
                    }
                }
                $print("是否有型号汇总表",haveModelType);
                if (!haveModelType) {
                    if (1 == $MsgBox("没有型号汇总表，是否新建？")) {
                        $SheetsLastAdd('型号汇总');
                        Application.Worksheets.Item('型号汇总').Activate();
                        let titles = [['筛选','计划员','类型', '型号', '链接','年度总计划','军种','军种分计划','完成入库','未入库','计划完成','计划未完','说明','完成率'], 
                        ['填“是”参与筛选，所有型号均不填则全部参与筛选', '主管计划员','批产/科研','计划系统型号名', '点击更新链接', '', '', '','','', '', '', '', '']];
                        $RangeInValue2("A1", 2, 14, titles,"@",'型号汇总');
                    } else {
                        alert('已取消建表');
                    }
                }
                break;
            }
            case "plannerIndex":{
                $print('plannerIndex start');
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                 // 获取选择的计划员
                var planner = "";
                const radios = document.querySelectorAll('input[name="planner"]');
                for (let radio of radios) {
                    if (radio.checked) {
                        planner = radio.value;
                        break;
                    }
                }
                if(""==planner){
                    alert("请选择计划员姓名");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == planner) {
                        Application.Worksheets.Item(i).Activate();
                        haveModelType = true;
                        break;
                    }
                }
                $print("是否有planner表",haveModelType);
                if (!haveModelType) {
                    alert('没有计划员：'+planner+'的表，请新建');
                }
                break;
            }
            case "creatOwnModel":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                // 检查型号汇总表是否存在
                const activeWorkbook = window.Application.ActiveWorkbook;
                let modelSheet = null;
                for (let i = 1; i <= activeWorkbook.Worksheets.Count; i++) {
                    const sheet = activeWorkbook.Worksheets.Item(i);
                    if (sheet.Name === "型号汇总") {
                        modelSheet = sheet;
                        break;
                    }
                }
                if (!modelSheet) {
                    alert("型号汇总表不存在");
                    return;
                }
                // 获取选择的计划员
                var planner = "";
                const radios = document.querySelectorAll('input[name="planner"]');
                for (let radio of radios) {
                    if (radio.checked) {
                        planner = radio.value;
                        break;
                    }
                }
                if(""==planner){
                    alert("请选择计划员姓名");
                    return;
                }

                

                // 筛选B列对应计划员的数据（假设B列是第2列，从第2行开始）
                const usedRange = modelSheet.Range("A1:N"+$SheetsLastRowNum(modelSheet.Name,4));
                const data = usedRange.Value2;
                const filteredData = [];
                filteredData.push(data[0]);
                filteredData.push(data[1]);
                for (let row = 0; row < data.length; row++) {
                    if (data[row][1] == planner) {
                        filteredData.push(data[row]);
                    }
                }
                $print("筛选后的数据",$ArrayToString(filteredData));

                // 处理目标表（计划员表）
                let targetSheet = activeWorkbook.Worksheets.Item(planner);
                if (targetSheet) {
                    // 存在则重命名旧表为待处理+日期时间+计划员
                    const formattedDate = $DateFormat(new Date(),"yyyyMMddhhmmss");
                    targetSheet.Name = `待处理${formattedDate}${planner}`;
                }

                // 创建新表
                const newSheet = activeWorkbook.Worksheets.Add();
                newSheet.Name = planner;
                $RangeA1InValue2(filteredData,planner);

                // 写入筛选后的数据（包括E列超链接）
                for(let i=3;i<=$SheetsLastRowNum(planner,3);i++){
                    //超链接
                    var modelName = $SheetsByName(planner).Range("D"+i).Value2;
                    $print("modelName",$ArrayToString(modelName));
                    var modelNameFormat = modelName.replace(/\//g, '');
                    $SheetsByName(planner).Hyperlinks.Add($SheetsByName(planner).Range("E"+i), "", "'"+modelNameFormat+"'!A1", "", "链接");
                }

                alert("计划员表更新完成");
                break;
            
            }
            case "updateModelLink":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("没有型号汇总表");
                    return;
                }
                // var planner = "";
                // const radios = document.querySelectorAll('input[name="planner"]');
                // for (let radio of radios) {
                //     if (radio.checked) {
                //         planner = radio.value;
                //         break;
                //     }
                // }
                // if(""==planner){
                //     alert("请选择计划员姓名");
                //     return;
                // }
                // var modelName = $InputBox("输入要新建的型号名，例如RXB-10:",'输入型号',"");
                // $print("输入型号",modelName);
                // if(null==modelName||""==modelName){
                //     alert("未录入任何信息");
                //     return;
                // }
                
                // for (let i = 1; i <= Application.Worksheets.Count; i++) {
                //     if (Application.Worksheets.Item(i).Name == modelNameFormat) {
                //         Application.Worksheets.Item(i).Activate();
                //         alert("该型号已存在");
                //         return;
                //     }
                // }
                $SheetsActivate('型号汇总');
                // $print($SheetsLastRowNum('型号目录',1));
                var lastRow = $SheetsLastRowNum('型号汇总',3);
                // $print("lastRow",lastRow);
                if(lastRow==2){
                    alert("无型号汇总数据");
                    return;
                }
                var newModelPlanner = $SheetTheActive().Range("D3:E"+lastRow).Value2;
                newModelPlanner.forEach((item)=>{
                    item[1] = "链接"; 
                });
                $RangeInValue2("D3",lastRow-2,2,newModelPlanner,"@",'型号汇总');
                for(let i=3;i<=lastRow;i++){
                    //超链接
                    var modelName = $SheetTheActive().Range("D"+i).Value2;
                    $print("modelName",$ArrayToString(modelName));
                    var modelNameFormat = modelName.replace(/\//g, '');
                    $SheetTheActive().Hyperlinks.Add($SheetTheActive().Range("E"+i), "", "'"+modelNameFormat+"'!A1", "", "链接");
                }
                const sheets = window.Application.Worksheets;
                const excludeSheets = ["同步记录勿删", "型号汇总", "陆会罗", "白旭", "朱志超", "王海岩", "崔玉玺"];
                for (let i = 1; i <= sheets.Count; i++) {
                    const sheet = sheets.Item(i);
                    if (!excludeSheets.includes(sheet.Name)) {
                        const a1Cell = sheet.Range("A1");
                        const originalText = a1Cell.Value;
                        if (originalText !== undefined) {
                            sheet.Hyperlinks.Add(
                                a1Cell,
                                "",
                                "'型号汇总'!A1",
                                "",
                                originalText
                            );
                        }
                    }
                }
                // $RangeInValue2("A"+$SheetsLastRowNum('型号目录',1),1,3,newModelPlanner,"@",'型号目录');
                
                //创建表
                // $SheetsLastAdd(modelNameFormat);
            
                // let titles = [['型号', '计划来源', '军种','主制部门','任务年度','泵号','出库日期','待军日期','入库日期','备注','记录人'], 
                // ['型号', '规划或其他等', '军种','十车间/十一车间，不要写简称','2025等','产品序列号（补全）','短日期格式：20XX/XX/XX，勿修改格式','短日期格式：20XX/XX/XX，勿修改格式',
                //     '短日期格式：20XX/XX/XX，勿修改格式','技术管理状态等信息','记录人姓名']];
                // $RangeInValue2("A1", 2, 11, titles,"@",modelNameFormat);
                // $SheetTheActive().Columns.Item(7).NumberFormatLocal = "yyyy/mm/dd";
                // $SheetTheActive().Columns.Item(8).NumberFormatLocal = "yyyy/mm/dd";
                // $SheetTheActive().Columns.Item(9).NumberFormatLocal = "yyyy/mm/dd";
                // Application.Worksheets.Item(modelNameFormat).Activate();
                alert("更新完成");
                break;
            }
            case "updateModel":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("没有型号汇总表");
                    return;
                }
                // var planner = "";
                // const radios = document.querySelectorAll('input[name="planner"]');
                // for (let radio of radios) {
                //     if (radio.checked) {
                //         planner = radio.value;
                //         break;
                //     }
                // }
                // if(""==planner){
                //     alert("请选择计划员姓名");
                //     return;
                // }
                if($SheetTheActive().Name!='型号汇总'||$RangeSelection().Row<3){
                    alert("请选择“型号汇总”表中，要修改的型号或同行单元格");
                    return;
                }
                var oldModelName = $SheetTheActive().Cells.Item($RangeSelection().Row,4).Value2;
                if(oldModelName==''||oldModelName==null){
                    alert("请选择“型号汇总”表中，要修改的型号或同行单元格");
                    return;
                }

                var modelName = $InputBox("输入要新的型号名，例如RXB-10:",'输入型号',"");
                $print("输入型号",modelName);
                if(null==modelName||""==modelName){
                    alert("未录入任何信息");
                    return;
                }
                var oldModelNameFormat = oldModelName.replace(/\//g, '');
                var modelNameFormat = modelName.replace(/\//g, '');
                var haveSheet = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == modelNameFormat) {
                        Application.Worksheets.Item(i).Activate();
                        alert("新型号名已存在，请检查");
                        return;
                    }
                    if (Application.Worksheets.Item(i).Name == oldModelNameFormat) {
                        Application.Worksheets.Item(i).Name = modelNameFormat;
                        
                        //超链接
                        $SheetTheActive().Hyperlinks.Add($SheetTheActive().Range("E"+$RangeSelection().Row), "", "'"+modelNameFormat+"'!A1", "", "链接");

                        $SheetTheActive().Cells.Item($RangeSelection().Row,4).Value2=[modelName];
                        // $SheetTheActive().Cells.Item($RangeSelection().Row,3).Value2=[planner];
                        haveSheet = true;
                        return;
                    }
                }
                if(!haveSheet){
                    alert("型号详细记录表丢失，请检查型号名和型号详细记录表Sheet名是否一致");
                    return;
                }
                break;
            }
            case "deleteModel":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("没有型号汇总表");
                    return;
                }
                if($SheetTheActive().Name!='型号汇总'||$RangeSelection().Row<3){
                    alert("请选择“型号汇总”表中，要删除的型号或同行单元格");
                    return;
                }
                var oldModelName = $SheetTheActive().Cells.Item($RangeSelection().Row,4).Value2;
                if(oldModelName==''||oldModelName==null){
                    alert("请选择“型号汇总”表中，要删除的型号或同行单元格");
                    return;
                }
                var oldModelNameFormat = oldModelName.replace(/\//g, '');
                // if($MsgBox("删除型号将永久丢失记录数据，请再次确认要删除型号："+oldModelName)){
                    // $print(oldModelNameFormat);
                    // $SheetTheActive().Rows.Item($RangeSelection().Row).Delete();
                    if($SheetsIsHave(oldModelNameFormat)){
                        $SheetDelOrSign(oldModelNameFormat);
                        alert("删除成功,请手动处理型号汇总和各人汇总表中的记录");
                    }else{
                        alert("没有该型号的入库详情表,请检查");
                    return;
                    }
                // }
                
                break;
            }
            case "creatModel":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("没有型号汇总表");
                    return;
                }
                var planner = "";
                const radios = document.querySelectorAll('input[name="planner"]');
                for (let radio of radios) {
                    if (radio.checked) {
                        planner = radio.value;
                        break;
                    }
                }
                if(""==planner){
                    alert("请选择计划员姓名");
                    return;
                }
                var modelName = $InputBox("输入要新建的型号名，例如RXB-10:",'输入型号',"");
                // $print("输入型号",modelName);
                if(null==modelName||""==modelName){
                    alert("未录入任何信息");
                    return;
                }
                var modelType = $MsgBox("是否为批产型号？是为批产，否为科研")
                if(modelType){
                    modelType = "批产";
                }else{
                    modelType = "科研";
                }

                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == modelNameFormat) {
                        Application.Worksheets.Item(i).Activate();
                        alert("该型号已存在");
                        return;
                    }
                }
                $SheetsActivate('型号汇总');

                // $print($SheetsLastRowNum('型号目录',1));
                var lastRow = $SheetsLastRowNum('型号汇总',3);
                // $print("lastRow",lastRow);
                // if(lastRow==2){
                //     alert("无型号汇总数据");
                //     return;
                // }
                // var newModelPlanner = $SheetTheActive().Range("D3:E"+lastRow).Value2;
                var newModelPlanner = ["",planner,modelType,modelName,"链接"];
                // newModelPlanner.forEach((item)=>{
                //     item[1] = "链接"; 
                // });
                $RangeInValue2("A"+(lastRow+1),1,5,newModelPlanner,"@",'型号汇总');
                var modelNameFormat = modelName.replace(/\//g, '');
                $SheetTheActive().Hyperlinks.Add($SheetTheActive().Range("E"+(lastRow+1)), "", "'"+modelNameFormat+"'!A1", "", "链接");
                
                
                //创建表
                $SheetsLastAdd(modelNameFormat);
                Application.Worksheets.Item(modelNameFormat).Activate();
                let titles = [['序号', '型号', '泵号','军种','主制车间','出库日期','待军日期','入库日期','备注','需求日期','其他'], 
                ['点击序号回到型号汇总', '', '','','十车间/十一车间，不要写简称','短日期格式：20XX/XX/XX，勿修改格式','短日期格式：20XX/XX/XX，勿修改格式','短日期格式：20XX/XX/XX，勿修改格式','短日期格式：20XX/XX/XX，勿修改格式',
                    '短日期格式：20XX/XX/XX，勿修改格式','','','包括此列开始可自由补充列']];
                $RangeInValue2("A1", 2, 11, titles,"@",modelNameFormat);
                $SheetTheActive().Columns.Item(5).NumberFormatLocal = "yyyy/mm/dd";
                $SheetTheActive().Columns.Item(6).NumberFormatLocal = "yyyy/mm/dd";
                $SheetTheActive().Columns.Item(7).NumberFormatLocal = "yyyy/mm/dd";
                const a1Cell = $SheetTheActive().Range("A1");
                const b1Cell = $SheetTheActive().Range("B1");
                const originalText = a1Cell.Value;
                if (originalText !== undefined){
                    $SheetTheActive().Hyperlinks.Add(
                        a1Cell,
                        "",
                        "'型号汇总'!A1",
                        "",
                        originalText
                    );
                }
                if (originalText !== undefined){
                    $SheetTheActive().Hyperlinks.Add(
                        b1Cell,
                        "",
                        "'"+planner+"'!A1",
                        "",
                        originalText
                    );
                }
                //冻结前2行
                $RangeSelect("A3");
                Application.Windows.Item(1).FreezePanes = true;
                $SheetsActivate('型号汇总');
                alert("创建完成");
                break;
            }
            case "reSet":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("没有型号汇总表");
                    return;
                }
                var modelSheet = Application.Worksheets.Item('型号汇总');
                if($SheetsLastRowNum('型号汇总',4)<3){
                    return;
                }
                
                var resultModel = [];
                modelSheet.Range("A3:A"+$SheetsLastRowNum('型号汇总',4)).Value2.forEach(eachRow=>{
                    eachRow[0]=null;
                    resultModel.push(eachRow);
                });
                modelSheet.Range("A3:A"+$SheetsLastRowNum('型号汇总',4)).Value2 = resultModel;

                // document.getElementById('rwnd1').value = "";
                // document.getElementById('rwnd2').value = "";
                document.getElementById('ckrq1').value = null;
                document.getElementById('ckrq2').value = null;
                document.getElementById('rkrq1').value = null;
                document.getElementById('rkrq2').value = null;
                document.getElementById('jz').value = null;
                document.getElementById('zzbm').value = null;
                break;
            }
            case "startCreate":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号汇总') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("没有型号汇总表");
                    return;
                }
                var modelSheet = Application.Worksheets.Item('型号汇总');
                if($SheetsLastRowNum('型号汇总',4)<3){
                    alert("型号汇总表没有型号数据");
                    return;
                }
                
                var resultModel = [];
                var haveSelectModel = false;
                modelSheet.Range("D3:D"+$SheetsLastRowNum('型号汇总',4)).Value2.forEach(eachRow=>{
                    if("是"==eachRow[0]){
                        resultModel.push(eachRow);
                        haveSelectModel = true;
                    }
                });
                if(!haveSelectModel){
                    resultModel = modelSheet.Range("D3:D"+$SheetsLastRowNum('型号汇总',4)).Value2;
                }
                // var rwnd1=document.getElementById('rwnd1').value;
                // if(rwnd1!=null&&rwnd1!=""){
                //     rwnd1=Number(rwnd1);
                // }else{
                //     rwnd1=0;
                // }
                // var rwnd2=document.getElementById('rwnd2').value;
                // if(rwnd2!=null&&rwnd2!=""){
                //     rwnd2=Number(rwnd2);
                // }else{
                //     rwnd2=9999;
                // }
                var ckrq1=document.getElementById('ckrq1').value;
                if(ckrq1!=null&&ckrq1!=""){
                    ckrq1=new Date(ckrq1+"T00:00:00Z");
                }else{
                    ckrq1=new Date("1990-01-01T00:00:00Z");
                }
                var ckrq2=document.getElementById('ckrq2').value;
                if(ckrq2!=null&&ckrq2!=""){
                    ckrq2=new Date(ckrq2+"T00:00:00Z");
                }else{
                    ckrq2=new Date("2099-01-01T00:00:00Z");
                }
                var rkrq1=document.getElementById('rkrq1').value;
                if(rkrq1!=null&&rkrq1!=""){
                    rkrq1=new Date(rkrq1+"T00:00:00Z");
                }else{
                    rkrq1=new Date("1990-01-01T00:00:00Z");
                }
                var rkrq2=document.getElementById('rkrq2').value;
                if(rkrq2!=null&&rkrq2!=""){
                    rkrq2=new Date(rkrq2+"T00:00:00Z");
                }else{
                    rkrq2=new Date("2099-01-01T00:00:00Z");
                }
                var jz=document.getElementById('jz').value;
                var zzbm=document.getElementById('zzbm').value;

                // $print("rwnd1",document.getElementById('rwnd1').value);
                // $print("rwnd2",document.getElementById('rwnd2').value);
                $print("ckrq1",document.getElementById('ckrq1').value);
                $print("ckrq2",document.getElementById('ckrq2').value);
                $print("rkrq1",document.getElementById('rkrq1').value);
                $print("rkrq2",document.getElementById('rkrq2').value);
                $print("jz",document.getElementById('jz').value);
                $print("zzbm",document.getElementById('zzbm').value);
                $print("resultModel",$ArrayToString(resultModel));
                var resultArray = [];
                resultModel.forEach(eachModel=>{
                    var modelNameFormat = eachModel[0].replace(/\//g, '');
                    if($SheetsIsHave(modelNameFormat)){
                        if($SheetsLastRowNum(modelNameFormat,3)>2){
                            var eachModelArray = Application.Worksheets.Item(modelNameFormat).Range("A3:K"+$SheetsLastRowNum(modelNameFormat,3)).Value2;
                            eachModelArray.forEach((eachRow,i,arraySelf)=>{
                                eachRow.forEach((eachval,j,arr)=>{
                                    if(eachval==null){arr[j]=""};
                                });
                                var ckrq = eachRow[5];
                                if($DateExcelValueIsDate(ckrq)){
                                    ckrq=$DateExcelToJS(ckrq);
                                }else if(ckrq!=null&&ckrq!=""){
                                    alert("型号："+eachModel+"的第"+(i+3)+"行的出库日期数据："+ckrq+" 格式错误");
                                    return;
                                }
                                var rkrq = eachRow[7];
                                if($DateExcelValueIsDate(rkrq)){
                                    rkrq=$DateExcelToJS(rkrq);
                                }else if(rkrq!=null&&rkrq!=""){
                                    alert("型号："+eachModel+"的第"+(i+3)+"行的入库日期数据："+rkrq+" 格式错误");
                                    return;
                                }
                                // var rwnd = eachRow[4];
                                // if($StringIsNumber(rwnd)){
                                //     rwnd=Number(rwnd);
                                // }else if(rkrq!=null&&rkrq!=""){
                                //     alert("型号："+eachModel+"的第"+(i+3)+"行的年度任务数据："+rkrq+" 格式错误");
                                //     return;
                                // }
                                var zzbm0 = eachRow[3].toString();
                                var jz0 = eachRow[2].toString();
                                if((ckrq>=ckrq1&&ckrq<=ckrq2)||(ckrq==null||ckrq=="")){
                                    if((rkrq>=rkrq1&&rkrq<=rkrq2)||(rkrq==null||rkrq=="")){
                                        if((jz!=null&&jz!=""&&jz0.includes(jz))||(jz==null||jz=="")){
                                            if((zzbm!=null&&zzbm!=""&&zzbm0.includes(zzbm))||(zzbm==null||zzbm=="")){
                                                // if(rwnd>=rwnd1&&rwnd<=rwnd2){
                                                    resultArray.push(eachRow);
                                                // }
                                            }
                                        }
                                    }
                                }
                            });
                        }
                    }
                });
                $print("resultArray",$ArrayToString(resultArray));
                // $print("rwnd1",rwnd1);
                // $print("rwnd2",rwnd2);
                $print("ckrq1",ckrq1);
                $print("ckrq2",ckrq2);
                $print("rkrq1",rkrq1);
                $print("rkrq2",rkrq2);
                $print("jz",jz);
                $print("zzbm",zzbm);


                var workbookName = "入库记录筛选结果"+new Date().valueOf(); 
                Application.Workbooks.Add();
                window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                $SheetsLastAdd("入库记录筛选结果");
                Application.Worksheets.Item("入库记录筛选结果").Activate();
                let titles = ['序号', '型号', '泵号','军种','主制车间','出库日期','待军日期','入库日期','备注','需求日期','其他'];
                $RangeInValue2("A1", 1, 11, titles,"@","入库记录筛选结果");
                if(resultArray.length==0){
                    alert("筛选结果为空");
                    return;
                }
                $RangeInValue2("A2", resultArray.length, 11, resultArray,"@","入库记录筛选结果");
                $SheetTheActive().Columns.Item(7).NumberFormatLocal = "yyyy/mm/dd";
                $SheetTheActive().Columns.Item(8).NumberFormatLocal = "yyyy/mm/dd";
                $SheetTheActive().Columns.Item(9).NumberFormatLocal = "yyyy/mm/dd";

                break;
            }
            case "toPlanOld":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if (!/^(?:\d{4}年)?产品入库记录\.xlsx$/.test(textValue)) {
                    alert("当前表不可用，请打开 “YYYY产品入库记录.xlsx”，其中YYYY为年份");
                    return;
                }
                var planner = "";
                const radios = document.querySelectorAll('input[name="planner"]');
                for (let radio of radios) {
                    if (radio.checked) {
                        planner = radio.value;
                        break;
                    }
                }
                if(""==planner){
                    alert("请选择计划员姓名");
                    return;
                }
                let haveGiveLog = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '同步记录勿删') {
                        haveGiveLog = true;
                        break;
                    }
                }
                if (!haveGiveLog) {
                    alert("同步记录勿删表丢失，若无法找回请先创建空表");
                    return;
                }

                var headSdr_found = $SheetTheActive().Range("A1:K1").Value2;
                //序号	型号	泵号	军种	主制车间	出库日期	待军日期	入库日期	备注	需求日期	其他
                var headSdr = [['序号', '型号', '泵号', '军种', '主制车间', '出库日期', '待军日期', '入库日期', '备注', '需求日期', '其他']];

                if(JSON.stringify(headSdr_found) !== JSON.stringify(headSdr)){
                    alert("型号表表头格式检查错误，请对比：序号\t型号\t泵号\t军种\t主制车间\t出库日期\t待军日期\t入库日期\t备注\t需求日期\t其他");
                    return;
                }
                var rowCount = $RangeSelection().Rows.Count;
                var rowStart = $RangeSelection().Rows.Item(0).Row+1;
                var rowEnd = $RangeSelection().Rows.Item(rowCount-1).Row+1;
                if(rowStart<3){
                    alert("不可选择前两行");
                    return;
                }
                var checkYes = true;
                var toPlanoldArray = $SheetTheActive().Range("A"+rowStart+":K"+rowEnd).Value2;
                toPlanoldArray.forEach((eachRow,i,arraySelf)=>{
                    if(eachRow[1]==null||eachRow[1]==""){
                        alert("第"+(i+rowStart)+"行型号不可为空");
                        checkYes = false;
                    }
                    // if(eachRow[4]==null||eachRow[4]==""){
                    //     alert("第"+(i+rowStart)+"行任务年度不可为空");
                    //     return;
                    // }
                    if(eachRow[2]==null||eachRow[2]==""){
                        alert("第"+(i+rowStart)+"行泵号不可为空");
                        checkYes = false;
                    }
                    if(eachRow[7]==null||eachRow[7]==""){
                        alert("第"+(i+rowStart)+"行入库日期不可为空");
                        checkYes = false;
                    }else if(!$DateExcelValueIsDate(eachRow[7])){
                        alert("第"+(i+rowStart)+"行入库日期格式错误");
                        checkYes = false;
                    }
                });
                if(!checkYes){
                    return;
                }

                var dataGridColModel =  [
                    { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                    { label : '型号', name : 'sn', width: 150,align: 'left'},
                    { label : '型号名', name : 'name', width: 150, align: 'left'},
                    { label : '分类', name : 'category', width: 150, align: 'center'},
                    { label : '系列', name : 'series', width: 150, align: 'left'},
                    { label : '特殊型号', name : 'specialName', width: 150, align: 'center'},
                    { label : '状态', name : 'state', width: 150, align: 'center'},
                    { label : '说明', name : 'description', width: 150, align: 'left'},
                    { label : '创建时间', name : 'createTime',  width: 150, align: 'center'},
                    { label : '搜索编号', name : 'searchSn', width: 150, align: 'left'}
                ];
                var params = {};
                var method = "getModel";
                // const url = "http://192.168.70.26:8080/V6R343/ws/dynPMResModelWS0";
                selectServer().then(baseServerUrl => {
                    var url = baseServerUrl + "ws/dynPMResModelWS0";
                    var targetNamespace0 = "http://ws.dynpmresmodel.planold.avicit/";
                    // 原有的$WebService调用逻辑const url = baseServerUrl + "ws/dynPMRepInBatchWS0";
                    return $WebService(url, targetNamespace0, method, params,dataGridColModel);
                }).then(tableData => {
                    // console.log('返回的 JSON 字符串:', tableData);
                    // var arr = $Array2DFromJsonStr(jsonStr);
                    // $print($ArrayToString(tableData));
                    if (!tableData || !Array.isArray(tableData)) {
                        alert("获取型号数据失败，返回数据格式不正确");
                        return Promise.reject("获取型号数据失败");
                    }
                    var haveModel=[];
                    var models = [];
                    for(var i=0;i<toPlanoldArray.length;i++){
                        haveModel.push(false);
                        models.push(toPlanoldArray[i][1]);
                    }
                    for(var j=0;j<tableData.length;j++){
                        for(var i=0;i<models.length;i++){
                            if(models[i]==tableData[j][1]){
                                toPlanoldArray[i].push(tableData[j][3]);
                                toPlanoldArray[i].push(planner);
                                toPlanoldArray[i].push($DateTimeJsToExcel(new Date()));
                                toPlanoldArray[i].push("发起导入计划系统");
                                haveModel[i]=true;
                                continue;
                            }
                        }
                    }
                    var modelNotFound = "";
                    haveModel.forEach((eachBoolean,i,arraySelf)=>{
                        if(!eachBoolean){
                            modelNotFound+="第"+(i+rowStart)+"行型号:"+models[i]+" 计划系统不存在;";
                        }
                    });
                    if(modelNotFound!=""){
                        alert(modelNotFound);
                        return;
                    }
                    $print($ArrayToString(toPlanoldArray));
                    // $print($ArrayToString(tableData));
                    var toPlanoldLogArray = $SheetsByName("同步记录勿删").Range("A1:N"+$SheetsLastRowNum("同步记录勿删",2)).Value2;
                    var hadIn = "";
                    for(var j=0;j<toPlanoldLogArray.length;j++){
                        for(var i=0;i<toPlanoldArray.length;i++){
                            if(toPlanoldLogArray[j][1]==toPlanoldArray[i][1]&&toPlanoldLogArray[j][2]==toPlanoldArray[i][2]){
                                hadIn+="型号："+toPlanoldArray[i][1]+" 的泵号："+toPlanoldArray[i][2] + "存在同步记录；";
                            }
                        }
                    }
                    if($MsgBox(hadIn+"确定导入？")){
                        
                        var toPlanoldMapList = [];
                        toPlanoldArray.forEach(eachToPlanold=>{
                            var eachMap = {
                                "型号":String(eachToPlanold[1]),
                                "军种":String(eachToPlanold[3]),
                                "泵号":String(eachToPlanold[2]),
                                "任务年度":String(new Date().getFullYear()),
                                "产品类型":String(eachToPlanold[11]),
                                "入库时间":String(eachToPlanold[7])
                            };
                            toPlanoldMapList.push(eachMap);
                        });
                        var jsonArrayString = JSON.stringify(toPlanoldMapList);
                        $SheetsByName("同步记录勿删").Range("A"+($SheetsLastRowNum("同步记录勿删",2)+1)
                             +":N"+($SheetsLastRowNum("同步记录勿删",1)+toPlanoldArray.length)).Value2=toPlanoldArray;
                        $SheetsActivate("同步记录勿删");

                        var dataGridColModel2 =  [
                            { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                            { label : '型号', name : 'modelSn',width: 150, align: 'left'},
                            { label : '产品编号', name : 'productNos', width: 150, align: 'left'},
                            { label : '军种', name : 'armyServices', width: 150, align: 'left'},
                            { label : '任务年度', name : 'year', width: 150, align: 'left'},
                            { label : '分类', name : 'category', width: 150, align: 'center'},
                            { label : '入库数量', name : 'count', width: 150, align: 'left'},
                            { label : '入库时间', name : 'createTime', width: 150, align: 'center'}  
                        ];    
                        var today = new Date();
                        var params2 = {jsonData:jsonArrayString};
                        var method2 = "addProductInStocks";
                        selectServer().then(baseServerUrl => {
                            var url2 = baseServerUrl + "ws/dynMRepProductRecordWS0";
                            var targetNamespace2 = "http://ws.dynmrepproductrecord.planold.avicit/";
                            return $WebService(url2, targetNamespace2, method2, params2, dataGridColModel2);
                        }).then(tableData => {
                            console.log('返回的 JSON 字符串:', tableData);
                            // var arr = $Array2DFromJsonStr(jsonStr);
                            $RangeInValue2("A"+($SheetsLastRowNum("同步记录勿删",2)+1), $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@","同步记录勿删");
                            alert("导入成功");

                        }).catch(error => {
                            console.error('调用 Web 服务失败:', error);
                            alert('调用 Web 服务失败:'+error);
                        });

                    }
                }).catch(error => {
                    console.error('调用 Web 服务失败:', error);
                    alert('调用 Web 服务失败:'+error);
                });
                break;
            }
            case "planOldProductInStock":{


                var dataGridColModel =  [
                    { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                    { label : '型号', name : 'modelSn',width: 150, align: 'left'},
                    { label : '军种', name : 'armyServices', width: 150, align: 'left'},
                    { label : '入库数量', name : 'count', width: 150, align: 'left'},
                    { label : '产品编号', name : 'productNos', width: 150, align: 'left'},
                    { label : '入库时间', name : 'createTime', width: 150, align: 'center'},
                    { label : '任务年度', name : 'year', width: 150, align: 'left'},
                    { label : '分类', name : 'category', width: 150, align: 'center'}
                ];
                try{
                    

                    const today = new Date();
                    const params = {};
                    const method = "getProductInStocks";
                    const targetNamespace="http://ws.dynmrepproductrecord.planold.avicit/"
                    var sheetName = "计划系统产品入库记录"+$StringReplaceAll(today.toLocaleDateString(),"/",".");
                    selectServer().then(baseServerUrl => {
                        const url = baseServerUrl + "ws/dynMRepProductRecordWS0";
                        
                        var isDo = $MsgBox("确认读取计划系统产品入库信息？");
                        if(1==isDo){
                            var workbookName = "计划系统产品入库记录"+new Date().valueOf(); 
                            Application.Workbooks.Add();
                            window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                            
                            $SheetsLastAdd(sheetName);
                            Application.Worksheets.Item(sheetName).Activate();
                            
                            // 调用函数
                            return $WebService(url, targetNamespace, method, params, dataGridColModel);
                        }else{
                            alert("已取消");
                            return Promise.reject("用户取消操作");
                        }
                    }).then(tableData => {
                        // console.log('返回的 JSON 字符串:', tableData);
                        // var arr = $Array2DFromJsonStr(jsonStr);
                        $RangeInValue2("A1", $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@",sheetName);
                    }).catch(error => {
                        if(error !== "用户取消操作") {
                            console.error('调用 Web 服务失败:', error);
                            alert('调用 Web 服务失败:'+error);
                        }
                    });
                            
                        
                }catch(ex){
                    $print("错误：",);
                    alert(ex);
                }
                break;
            }
        }
    } catch (ex) {
        $print("错误", ex.message);
        alert(ex.message);
    }
}


