// function btnCreateSOMSheetMain(){
//     try{
//         var sheetsCount = $SheetsCount();
//         // $print("返回类型",$ObjectGetType(sheetsCount));
//         var sheetName = $InputBox("输入表名，不要与现有表重名，建议为姓名+日期，例如：李文玉241112",'输入新建表名',"Sheet"+(sheetsCount+1));
//         // $print("后台输出输入值类型",$ObjectGetType(sheetName));
//         if(null!=sheetName&&""!=sheetName){
//             var sheetNames = $SheetsNamesArray();
//             if(sheetNames.includes(sheetName)){
//                 alert("创建失败！工作表：" + sheetName +"已存在！");
//             }else{
//                 // $print(Application.ActiveWorkbook.Worksheets.Count);
//                 $print("是否创建了1个新工作表："+sheetName,$SheetsLastAdd(sheetName));
//                 var thisSheet = $SheetTheActive();
//                 // $print("新建工作表的名字",thisSheet.Name);
//                 var titles = [['序号','主制车间','图号','名称','缺件数','备注','','','','',],
//                             ['数值','文本','文本','文本','数值','文本','','','','',]]
//                 // thisSheet.Range("A1:J2").Value2=titles;
//                 $RangeInValue2("A1",2,6,titles);
//                 alert("工作表：" + sheetName +"创建成功！");
//             }
//         }
//     }catch(ex){
//         $print("错误",ex);
//         alert("错误:"+ex);
//     }
    
// }

// 确保util.js已经被加载
function onbuttonclick_btnCreateSOMSheet(idStr) {
    try {
        if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
            window.Application.Enum = WPS_Enum;
        }
        switch (idStr) {
            case "dockLeft":
                $print('dockLeft start');
                let tsIdLeft = window.Application.PluginStorage.getItem("btnCreateSOMSheet_id");
                if (tsIdLeft) {
                    let tskpane = window.Application.GetTaskPane(tsIdLeft);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionLeft;
                }
                break;
            case "dockRight":
                $print('dockRight start');
                let tsIdRight = window.Application.PluginStorage.getItem("btnCreateSOMSheet_id");
                if (tsIdRight) {
                    let tskpane = window.Application.GetTaskPane(tsIdRight);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight;
                }
                break;
            case "hideTaskPane":
                $print('hideTaskPane start');
                let tsIdHide = window.Application.PluginStorage.getItem("btnCreateSOMSheet_id");
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
                document.getElementById("isbtnCreateSOMSheet").innerHTML = textValue
                break
            }
            case "modelType":{
                $print('modelType start');
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if (doc.Worksheets.Item(i).Name == '型号属性') {
                        doc.Worksheets.Item(i).Activate();
                        haveModelType = true;
                        break;
                    }
                }
                $print("是否有型号属性表",haveModelType);
                if (!haveModelType) {
                    if (1 == $MsgBox("没有型号属性表，是否新建？")) {
                        $SheetsLastAdd('型号属性');
                        doc.Worksheets.Item('型号属性').Activate();
                        let titles = [['型号', '批产/科研属性',''], ['填写型号名', '填写 批产/科研','缺件在型号需求表维护']];
                        $RangeInValue2("A1", 2, 3, titles,"@",'型号属性');
                    } else {
                        alert('已取消建表，需维护型号属性后才能正确区分科研/批产配套计划');
                    }
                }
                break;
            }
            case "modelSum":{
                $print('modelSum start');
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if (doc.Worksheets.Item(i).Name == '型号需求') {
                        doc.Worksheets.Item(i).Activate();
                        haveModelType = true;
                        break;
                    }
                }
                $print("是否有型号需求表",haveModelType);
                if (!haveModelType) {
                    if (1 == $MsgBox("没有型号需求表，是否新建？")) {
                        $SheetsLastAdd('型号需求');
                        doc.Worksheets.Item('型号需求').Activate();
                        let titles = [['型号', '备注', '月份1', '月份2', '月份3', ''], ['填写型号名', '备注','月份1需求', '月份2需求', '月份3需求', '目前暂时支持计算三个月缺件']];
                        $RangeInValue2("A1", 2, 6, titles,"@",'型号需求');
                    } else {
                        alert('已取消建表');
                    }
                }
                break;
            }
            case "creatPlan":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                // var selectYear = document.getElementById('yearSelect').value;
                // var selectMonth = document.getElementById('monthSelect').value;
                // $print(selectYear,typeof selectYear);
                // $print(selectMonth,typeof selectMonth);

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
                // var sheetName = planner + "@" + selectYear + "@" + selectMonth;
                var sheetName = planner;
                let haveSheet = false;
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if (doc.Worksheets.Item(i).Name == sheetName) {
                        doc.Worksheets.Item(i).Activate();
                        haveSheet = true;
                        break;
                    }
                }
                if (!haveSheet) {
                    if (1 == $MsgBox("没有该月缺件表："+sheetName+",是否新建？")) {
                        $SheetsLastAdd(sheetName);
                        doc.Worksheets.Item(sheetName).Activate();

                        // var months = [];
                        // var tempDate = Number(selectMonth);
                        // months.push(tempDate+"月");

                        // for(var i=0;i<11;i++){
                        //     if(tempDate==12){
                        //         tempDate = 1;
                        //     }else{
                        //         tempDate++;
                        //     }
                        //     months.push(tempDate+"月");
                        // }
                        // $print($ArrayToString(months));

                        let titles = [["缺件-"+planner,'','','','','','','','','','',''],
                        ['类型','型号', '图号','名称','主制','总计',
                            '1月','2月','3月','节点','库房','备注'], 
                        ['批产/科研','所有数据从第四列开始维护，不要删除表头和插入空行','','','一车间、二车间、三车间、六车间、八车间、十车间、十一车间、采购','','','','','','']];
                        $RangeInValue2("A1", 3, 12, titles,"@",sheetName);
                        doc.Worksheets.Item(sheetName).Range("A1:L1").Merge();
                    } else {
                        alert('已取消');
                    }
                }
                break;
            }

            case "addPTType":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                let thisSheet = window.Application.ActiveWorkbook.ActiveSheet
                // const regex = /^[\u4e00-\u9fa5]+@\d{4}@[1-9]|1[0-2]$/;
                var planners = ['陆会罗','白旭','王海岩','朱志超','崔玉玺'];
                // if (!regex.test(thisSheet.Name)) {
                if (planners.indexOf(thisSheet.Name)==-1) {
                    alert("工作表名称格式不符");
                    return;
                }else{
                    $print(thisSheet.Name,"表名验证完成");
                }
                if(thisSheet.Range("A2").Value2!="类型"||thisSheet.Range("B2").Value2!="型号"){
                    $print("A2\B2为类型\型号");
                    alert("工作表格式不符");
                    return;
                }else{
                    $print(thisSheet.Name,"表列验证完成");
                }

                let haveModelType = false;
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if (doc.Worksheets.Item(i).Name == '型号属性') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("型号属性表不存在，需维护");
                    return;
                }else{
                    var modelTypeSheet = doc.Worksheets.Item("型号属性");
                    const lastRow = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 2).End(-4162).Row;
                    if(lastRow<3){
                        alert("型号属性表无内容，需维护");
                        return; 
                    }
                }
                var typeIsOK = true;//型号是否全部正确匹配到了类型
                // 获取 B 列的最后一行
                const lastRow = thisSheet.Cells.Item(thisSheet.Rows.Count, 2).End(-4162).Row; // -4162 表示向上查找最后一个非空单元格
                if(lastRow>3){
                    var modelArray = [];
                    if(lastRow==4){
                        modelArray.push([thisSheet.Range("B4").Value2]);
                    }else{
                        modelArray = thisSheet.Range("B4:B"+lastRow).Value2;
                    }
                    // $print($ArrayToString(modelArray));
                    var modelTypeSheet = doc.Worksheets.Item("型号属性");
                    const lastRow0 = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 2).End(-4162).Row;
                    var modelTypeArray = modelTypeSheet.Range("A3:B"+lastRow0).Value2;
                    const modelTypeMap = new Map();
                    for (const [key, value] of modelTypeArray) {
                        modelTypeMap.set(key, value);
                    }
                    $print("modelArray",lastRow+":"+$ArrayToString(modelArray));
                    $print("modelTypeMap",JSON.stringify(Array.from(modelTypeMap)));
                    var typeArray = [];
                    
                    for(var i=0;i<modelArray.length;i++){
                        var typeVal = modelTypeMap.get(modelArray[i][0]);
                            if(undefined===typeVal){
                                typeArray.push("无数据");
                                typeIsOK = false;
                            }else{
                                typeArray.push(typeVal);
                            }
                    }
                    $print("typeArray",$ArrayToString(typeArray));
                    $RangeInValue2("A4", modelArray.length, 1, $ArrayTranspose(typeArray),"@",thisSheet.Name);
                }
                alert("刷新完成,"+(typeIsOK?"不":"注意！")+"存在型号类型不对应情况");
                break;
            }
            
            case "checkPlan":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                let thisSheet = window.Application.ActiveWorkbook.ActiveSheet
                // const regex = /^[\u4e00-\u9fa5]+@\d{4}@[1-9]|1[0-2]$/;
                var planners = ['陆会罗','白旭','王海岩','朱志超','崔玉玺'];
                // if (!regex.test(thisSheet.Name)) 
                if (planners.indexOf(thisSheet.Name)==-1) {
                    alert("工作表名称格式不符");
                    return;
                }else{
                    $print(thisSheet.Name,"表名验证完成");
                }
                if(thisSheet.Range("A2").Value2!="类型"||thisSheet.Range("B2").Value2!="型号"){
                    $print("A2\B2为类型\型号");
                    alert("工作表格式不符");
                    return;
                }else{
                    $print(thisSheet.Name,"表列验证完成");
                }
                const lastRow = thisSheet.Cells.Item(thisSheet.Rows.Count, 2).End(-4162).Row;
                var valArray = thisSheet.Range("A4:T"+lastRow).Value2;
                var errorMsg = "";
                var zz = ['一车间','二车间','三车间','六车间','八车间','十车间','十一车间','采购']; //主制
                // var kf = ['中央库','采购库','成品库','一车间','二车间','三车间','六车间','八车间']; //库房
                valArray.forEach((item,row,arraySelf1) => {
                        if(item[0]==null||item[0]=="无数据"){
                            errorMsg+="第"+(4+row)+"行第A列类型错误;";
                        }
                        if(item[1]==null||item[1]==""){
                            errorMsg+="第"+(4+row)+"行第B列型号错误;";
                        }
                        if(item[2]==null||item[2]==""){
                            errorMsg+="第"+(4+row)+"行第C列图号错误;";
                        }
                        if(item[3]==null||item[3]==""){
                            errorMsg+="第"+(4+row)+"行第D列名称错误;";
                        }
                        if(item[4]==null||item[4]==""||!zz.includes(item[4])){
                            errorMsg+="第"+(4+row)+"行第E列主制错误;";
                        }
                });
                alert("检查完成："+errorMsg);
                break;
            }
            case "checkPlanIsAll":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                // var selectYear = document.getElementById('yearSelect').value;
                // var selectMonth = document.getElementById('monthSelect').value;
                var planners = ['陆会罗','白旭','王海岩','朱志超','崔玉玺'];
                var planNames = [];
                planners.forEach( planner=> {
                    // planNames.push(planner + "@" + selectYear + "@" + selectMonth);
                    planNames.push(planner);
                });
                const plans = new Map();
                planNames.forEach( planName=> {
                    plans.set(planName,false);
                });
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if(Array.from(plans.keys()).includes(doc.Worksheets.Item(i).Name)){
                        plans.set(doc.Worksheets.Item(i).Name,true);
                    }
                }
                var errorMsg = '';
                Array.from(plans.keys()).forEach((planName,i,arrayself)=> {
                    if(!plans.get(planName)){
                        errorMsg += planName + "不存在；";
                    }
                });
                alert("检查完成："+errorMsg);
                break;
            }
            case "createPC":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                // var selectYear = document.getElementById('yearSelect').value;
                // var selectMonth = document.getElementById('monthSelect').value;
                // var workbookName = selectYear + "年" + selectMonth +"批产缺件汇总"+new Date().valueOf(); 
                var planners = ['陆会罗','白旭','王海岩','朱志超','崔玉玺'];
                var planNames = [];
                planners.forEach( planner=> {
                    // planNames.push(planner + "@" + selectYear + "@" + selectMonth);
                    planNames.push(planner);
                });
                var pcValue = [];
                // var sumValue = [];
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if(planNames.includes(doc.Worksheets.Item(i).Name)){
                       //提取合并相关表格数据
                        var theSheet = doc.Worksheets.Item(i);
                        var lastRow = theSheet.Cells.Item(theSheet.Rows.Count, 2).End(-4162).Row;
                        if(lastRow<4){
                            continue;
                        }
                        var planArray = theSheet.Range("A4:L"+lastRow).Value2;
                        planArray.forEach( item=> {
                            if(item[0]=='批产'){
                                item.push(doc.Worksheets.Item(i).Name);
                                pcValue.push(item);
                            }
                            // sumValue.push(item);
                        });
                    }
                }
                if(pcValue.length==0){
                    alert('无数据');
                    return;
                }
                // $print($ArrayToString(sumValue));
                // $print($ArrayToString(pcValue));

                // Application.Workbooks.Add();
                // window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                var sheetName = "批产缺件汇总"+$StringReplaceAll(new Date().toLocaleDateString(),"/",".")+"."+new Date().valueOf();
                $SheetsLastAdd(sheetName);
                doc.Worksheets.Item(sheetName).Activate();

                // var months = [];
                // var tempDate = Number(selectMonth);
                // months.push(tempDate+"月");

                // for(var i=0;i<11;i++){
                //     if(tempDate==12){
                //         tempDate = 1;
                //     }else{
                //         tempDate++;
                //     }
                //     months.push(tempDate+"月");
                // }

                // $RangeInValue2("A1", 1, 12, ['类型','型号', '图号','名称','主制','总计',months[0],months[1],months[2],'节点','库房','备注']);
                $RangeInValue2("A1", 1, 13, ['类型','型号', '图号','名称','主制','总计',"月份1","月份2","月份3",'节点','库房','备注','分工'],"@",sheetName);
                $RangeInValue2("A2",pcValue.length,13,pcValue,"@",sheetName);
                // var zz = ['一车间','二车间','三车间','六车间','八车间','十车间','十一车间','采购']; //主制
                // zz.forEach(eachZZ=>{
                    
                //     $SheetsLastAdd(eachZZ);
                //     Application.Worksheets.Item(eachZZ).Activate();
                //     $RangeInValue2("A1", 1, 1, [eachZZ+selectYear + "年" +selectMonth+"月份缺件"]);
                //     Application.Worksheets.Item(eachZZ).Range("A1:L1").Merge();
                //     Application.Worksheets.Item(eachZZ).Range("A1:L1").HorizontalAlignment = -4108;
                //     Application.Worksheets.Item(eachZZ).Range("A1:L1").Font.Name = "宋体";      // 设置字体名称为宋体
                //     Application.Worksheets.Item(eachZZ).Range("A1:L1").Font.Bold = true;        // 设置字体加粗
                //     Application.Worksheets.Item(eachZZ).Range("A1:L1").Font.Size = 18;          // 设置字体大小为18号
                //     // $RangeInValue2("A2", 1, 12, ['序号','型号', '图号','名称','主制','总计',months[0],months[1],months[2],'节点','库房','备注']);
                //     $RangeInValue2("A2", 1, 12, ['序号','型号', '图号','名称','主制','总计',"1月","2月","3月",'节点','库房','备注']);
                //     Application.Worksheets.Item(eachZZ).Range("A2:L2").HorizontalAlignment = -4108;
                //     Application.Worksheets.Item(eachZZ).Range("A2:L2").Font.Name = "宋体";      // 设置字体名称为宋体
                //     Application.Worksheets.Item(eachZZ).Range("A2:L2").Font.Bold = true;        // 设置字体加粗
                //     Application.Worksheets.Item(eachZZ).Range("A2:L2").Font.Size = 12;          // 设置字体大小为12号

                //     eachZZArray = [];
                //     pcValue.forEach(eachValue=>{
                //         if(eachValue[4]==eachZZ){
                //             eachZZArray.push(eachValue);
                //         }
                //     });
                //     var indexNum = 1;
                //     eachZZArray.forEach(eachValue=>{
                //         if(eachValue[6]!=""&&eachValue[6]!=null){
                //             eachValue[0]=indexNum;
                //             indexNum++;
                //         }else{
                //             eachValue[0]='';
                //         }
                //     });
                //     if(eachZZArray.length!=0){
                //         $RangeInValue2("A3", eachZZArray.length, 12,eachZZArray,"@");
                //     }
                //     Application.Worksheets.Item(eachZZ).Range("A1").Resize(eachZZArray.length+2,12).HorizontalAlignment = -4108;
                //     Application.Worksheets.Item(eachZZ).Range("A1").Resize(eachZZArray.length+2,12).Borders.LineStyle = 1;
                //     Application.Worksheets.Item(eachZZ).Range("A1").Resize(eachZZArray.length+2,12).Borders.Weight = 2 ;
                // });

                break;
            }
            case "createXP":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('缺件管理.xlsx' != textValue) {
                    alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                    return;
                }
                // var selectYear = document.getElementById('yearSelect').value;
                // var selectMonth = document.getElementById('monthSelect').value;
                // var workbookName = selectYear + "年" + selectMonth +"月科研缺件汇总"+new Date().valueOf(); 
                var planners = ['陆会罗','白旭','王海岩','朱志超','崔玉玺'];
                var planNames = [];
                planners.forEach( planner=> {
                    // planNames.push(planner + "@" + selectYear + "@" + selectMonth);
                    planNames.push(planner);
                });
                var pcValue = [];
                // var sumValue = [];
                for (let i = 1; i <= doc.Worksheets.Count; i++) {
                    if(planNames.includes(doc.Worksheets.Item(i).Name)){
                       //提取合并相关表格数据
                        var theSheet = doc.Worksheets.Item(i);
                        var lastRow = theSheet.Cells.Item(theSheet.Rows.Count, 2).End(-4162).Row;
                        if(lastRow<4){
                            continue;
                        }
                        var planArray = theSheet.Range("A4:L"+lastRow).Value2;
                        planArray.forEach( item=> {
                            if(item[0]=='科研'){
                                item.push(doc.Worksheets.Item(i).Name);
                                pcValue.push(item);
                            }
                            // sumValue.push(item);
                        });
                        
                    }
                }
                if(pcValue.length==0){
                    alert('无数据');
                    return;
                }
                // $print($ArrayToString(sumValue));
                // $print($ArrayToString(pcValue));

                // Application.Workbooks.Add();
                // window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                var sheetName = "科研缺件汇总"+$StringReplaceAll(new Date().toLocaleDateString(),"/",".")+"."+new Date().valueOf();
                $SheetsLastAdd(sheetName);
                doc.Worksheets.Item(sheetName).Activate();
                

                // var months = [];
                // var tempDate = Number(selectMonth);
                // months.push(tempDate+"月");

                // for(var i=0;i<11;i++){
                //     if(tempDate==12){
                //         tempDate = 1;
                //     }else{
                //         tempDate++;
                //     }
                //     months.push(tempDate+"月");
                // }

                // $RangeInValue2("A1", 1, 12, ['类型','型号', '图号','名称','主制','总计',months[0],months[1],months[2],'节点','库房','备注']);
                $RangeInValue2("A1", 1, 13, ['类型','型号', '图号','名称','主制','总计','月份1','月份2','月份3','节点','库房','备注','分工'],"@",sheetName);
                $RangeInValue2("A2",pcValue.length,13,pcValue,"@",sheetName);

                break;
            }
            
            case "compute1": {
                var isDo = $MsgBox("是否开始计算缺件1？可能需要一些时间");

                if(isDo){
                    let doc = window.Application.ActiveWorkbook;
                    let textValue = "";
                    if (!doc) {
                        textValue = "当前没有打开任何文档";
                    } else {
                        textValue = doc.Name;
                    }
                    if ('缺件管理.xlsx' != textValue) {
                        alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                        return;
                    }
                    
                    let haveModelType = false;
                    for (let i = 1; i <= doc.Worksheets.Count; i++) {
                        if (doc.Worksheets.Item(i).Name == '型号需求') {
                            haveModelType = true;
                            break;
                        }
                    }
                    let thisSheet = window.Application.ActiveWorkbook.ActiveSheet;
                    if (!haveModelType) {
                        alert("型号需求表不存在，需维护");
                        return;
                    }else{
                        var modelTypeSheet = window.Application.ActiveWorkbook.Worksheets.Item("型号需求");
                        const lastRow = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 1).End(-4162).Row;
                        if(lastRow<3){
                            alert("型号需求表无内容，需维护");
                            return; 
                        }
                    }

                    var workbookName = "缺件结果"+new Date().valueOf(); 
                    Application.Workbooks.Add();
                    window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                    $SheetsLastAdd("计算进度");
                    Application.Worksheets.Item("计算进度").Activate();
                    var msgNum = 0;
                    var msgPercent = 1;
                    var msgItem = [["序号","进度","进展","时间"],
                        [msgNum++,msgPercent.toFixed(2)+"%","校验文件完成，开始读取零组件BOM",$DateFormatDateTime(new Date())]];
                    $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);


                    var dataGridColModel =  [
                        { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                        { label : '上级号', name : 'partsSn',width: 150, align: 'left'},
                        { label : '上级名', name : 'partsSnName', width: 120, align: 'left'},
                        { label : '上级主制', name : 'partsSnDept', width: 100, align: 'left'},
                        { label : '下级号', name : 'bomPartsSn', width: 150, align: 'left'},
                        { label : '下级名', name : 'bomPartsSnName', width: 120, align: 'left'},
                        { label : '下级货位', name : 'positionNo', width: 120, align: 'left'},
                        { label : '下级主制', name : 'bomPartsDept', width: 100, align: 'left'},
                        { label : '需求类型', name : 'typeName', width: 80, align: 'center'},
                        { label : '需求数量', name : 'requireCount', width: 80, align: 'right'},
                        { label : '顺序号', name : 'sequence', width: 80, align: 'right'},
                        { label : '创建时间', name : 'createTime',  width: 120, align: 'center'}
                    ];
                    // var sheetName = "零件BOM" + +new Date().valueOf(); 
                    // $SheetsLastAdd(sheetName);

                    // 调用函数
                    selectServer().then(baseServerUrl => {
                        const url = baseServerUrl + "ws/dynPMResPartsBomWS0";
                        const method = "getPartsBom";
                        const targetNamespace="http://ws.dynpmrespartsbom.planold.avicit/"
                        // const params = {days:dayIndex};
                        const params = {};
                        
                        return $WebService(url, targetNamespace, method, params, dataGridColModel);
                    }).then(tableDataPartsBom => {
                        // console.log('返回的 JSON 字符串:', tableData);
                        // var arr = $Array2DFromJsonStr(jsonStr);
                        // $RangeInValue2("A1", $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@");
                        msgPercent = 2;
                        msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","开始读取型号BOM",$DateFormatDateTime(new Date())]);
                        $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                        var dataGridColModel0 =  [
                            { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                            { label : '型号', name : 'modelSn', width: 150, align: 'left'},
                            { label : '零组件号', name : 'partsSn', width: 150, align: 'left'},
                            { label : '零组件名', name : 'partsName', width: 150, align: 'left'},
                            { label : '货位', name : 'positionNo', width: 150, align: 'left'},
                            { label : '主制部门', name : 'partsDept', width: 150, align: 'left'},
                            { label : '需求类型', name : 'typeName', width: 150, align: 'center'},
                            { label : '单组数量', name : 'requireCount', width: 150, align: 'right'},
                            { label : '创建时间', name : 'createTime', width: 150, align: 'center'},
                            { label : 'BOM类型', name : 'categoryName', width: 150, align: 'center'}
                        ];
                        // var sheetName0 = "型号BOM" + +new Date().valueOf(); 
                        // $SheetsLastAdd(sheetName0);
                    
                        // 调用函数
                        // const url0 = "http://192.168.70.26:8080/V6R343/ws/dynPMResModelBomWS0";
                        // const url0 = "http://192.168.70.17:8888/V6/ws/dynPMResModelBomWS0";
                        // const method0 = "getModelBom";
                        // const targetNamespace0="http://ws.dynpmresmodelbom.planold.avicit/";
                        // const params0 = {};
                        selectServer().then(baseServerUrl => {
                            const url0 = baseServerUrl + "ws/dynPMResModelBomWS0";
                            const method0 = "getModelBom";
                            const targetNamespace0 = "http://ws.dynpmresmodelbom.planold.avicit/";
                            const params0 = {};
                            return $WebService(url0, targetNamespace0, method0, params0, dataGridColModel0);
                        }).then(tableDataModelBom => {
                            // console.log('返回的 JSON 字符串:', tableData);
                            // var arr = $Array2DFromJsonStr(jsonStr);
                            // $RangeInValue2("A1", $Array2DRowColCount(tableData0).rows,$Array2DRowColCount(tableData0).cols,tableData0,"@");
                            msgPercent = 3;
                            msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","开始读取物料与库存信息",$DateFormatDateTime(new Date())]);
                            $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                            var dataGridColModel1 =  [
                                { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                                { label : '所属型号', name : 'belongModelSn',width: 150, align: 'left'},
                                { label : '图号', name : 'sn', width: 150, align: 'left'},
                                { label : '名称', name : 'name', width: 150, align: 'left'},
                                { label : '主制车间', name : 'deptname', width: 150, align: 'left'},
                                { label : '库房', name : 'storageCategory', width: 150, align: 'left'},
                                { label : '中央库库位库存', name : 'normalCount', width: 150, align: 'right'},
                                { label : '配套库存', name : 'ptsum', width: 150, align: 'right'},
                                { label : '组件在制', name : 'zzsum', width: 150, align: 'right'},
                                { label : '别名', name : 'alias', width: 150, align: 'left'},
                                { label : '是否组件', name : 'component', width: 150, align: 'left'},
                                { label : '当前时间', name : 'nowtime', width: 150, align: 'left'}
                            ];

                            const params1 = {};
                            const method1 = "getStocks";
                            // const url1 = "http://192.168.70.26:8080/V6R343/ws/dynStockPartsWS0";
                            // const url1 = "http://192.168.70.17:8888/V6/ws/dynStockPartsWS0";
                            // const targetNamespace1="http://ws.dynstockparts.planand.avicit/"
                            selectServer().then(baseServerUrl => {
                                const url1 = baseServerUrl + "ws/dynStockPartsWS0";
                                const targetNamespace1 = "http://ws.dynstockparts.planand.avicit/";
                                return $WebService(url1, targetNamespace1, method1, params1, dataGridColModel1);
                            }).then(tableDataStock => {
                                msgPercent = 5;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","读取物料与库存信息完成，开始整理型号需求",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                const errorFloor = 15; //最深计算层数
                                // 修复null错误：添加检查确保工作簿和工作表存在
                                var workbooks = window.Application.Workbooks;
                                if (!workbooks) {
                                    alert("无法访问工作簿集合");
                                    return;
                                }
                                
                                var缺件管理Workbook = workbooks.Item("缺件管理.xlsx");
                                if (!缺件管理Workbook) {
                                    alert("无法找到'缺件管理.xlsx'工作簿");
                                    return;
                                }
                                
                                var modelTypeSheet = 缺件管理Workbook.Worksheets.Item("型号需求");
                                if (!modelTypeSheet) {
                                    alert("无法找到'型号需求'工作表");
                                    return;
                                }
                                
                                const lastRow0 = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 1).End(-4162).Row;
                                var modelTypeArray = modelTypeSheet.Range("A3:E"+lastRow0).Value2;
                                var modelTypeArray1 = [];
                                for (var [model, modelType,month1,month2,month3] of modelTypeArray) {
                                    if(month1==null||month1==""){
                                        month1=0;
                                    }
                                    if(month2==null||month2==""){
                                        month2=0;
                                    }
                                    if(month3==null||month3==""){
                                        month3=0;
                                    }
                                    modelTypeArray1.push([model, modelType,month1,month2,month3]);
                                }
                                var testResult = [];
                                msgPercent = 6;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","完成型号需求汇总",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);


                                testResult.push(["0型号需求"]);
                                modelTypeArray1.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                var computResult = [['型号', '图号','名称','主制','总计','1月','2月','3月','组件信息','是否拆解','别名','库存','在制','剩余数量','库房']];
                                // var partIndex = [];
                                modelTypeArray1.forEach((eachModelType,num,modelTypeArr)=>{//逐型号计算
                                    tableDataModelBom.forEach(eachModelBom=>{
                                        // var itPartBom = [];    
                                        if(eachModelBom[1]==eachModelType[0]){
                                            var perNum = 1;
                                            if(eachModelBom[7]!=null&&eachModelBom[7]!=0&&eachModelBom[7]!=""){
                                                perNum=eachModelBom[7];
                                            }
                                            var havePart = false;
                                            for(i=0;i<computResult.length;i++){
                                                if(eachModelBom[2]==computResult[i][1]){
                                                    computResult[i][5]+=eachModelType[2]*perNum;
                                                    computResult[i][6]+=eachModelType[3]*perNum;
                                                    computResult[i][7]+=eachModelType[4]*perNum;
                                                    // arrResult[numResult][4]+=(eachModelType[2]*perNum+eachModelType[3]*perNum+eachModelType[4]*perNum);

                                                    // 使用 splice 删除指定索引的元素，并将其存储在变量中
                                                    const [element] = computResult.splice(i, 1);
                                                    // 使用 push 将元素添加到数组末尾
                                                    computResult.push(element);

                                                    havePart = true;
                                                    break;
                                                }  
                                            }
                                            if(!havePart){
                                                for(i=0;i<tableDataStock.length;i++){
                                                    if(tableDataStock[i][2]==eachModelBom[2]){
                                                        var modelName = tableDataStock[i][1];
                                                        var partName = tableDataStock[i][3];
                                                        var deptName = tableDataStock[i][4];
                                                        // var partSum = eachModelType[2]*perNum+eachModelType[3]*perNum+eachModelType[4]*perNum;
                                                        var partSum = 0;
                                                        var alias = tableDataStock[i][9];
                                                        var component = 0;
                                                        if(tableDataStock[i][10]==1.0||tableDataStock[i][10]=="1.0"||tableDataStock[i][10]=="1"){
                                                            component=1;
                                                        }
                                                        var stockIn = 0;
                                                        if(tableDataStock[i][6]!=null&&tableDataStock[i][6]!=""&&tableDataStock[i][6]!=0){
                                                            stockIn=Number(tableDataStock[i][6]);
                                                        }
                                                        var stockPT = 0;
                                                        if(tableDataStock[i][7]!=null&&tableDataStock[i][7]!=""&&tableDataStock[i][7]!=0){
                                                            stockPT=Number(tableDataStock[i][7]);
                                                        }
                                                        var stockZZ = 0;
                                                        if(tableDataStock[i][8]!=null&&tableDataStock[i][8]!=""&&tableDataStock[i][8]!=0){
                                                            stockZZ=Number(tableDataStock[i][8]);
                                                        }
                                                        var componentItem = [];
                                                        if(component==1){//添加组件信息【下级号，需求数量】
                                                            tableDataPartsBom.forEach(eachPartBom=>{
                                                                if(eachPartBom[1]==eachModelBom[2]){
                                                                    componentItem.push([eachPartBom[4],eachPartBom[9]]);
                                                                }
                                                            });
                                                        }
                                                        var perPartRequest = [modelName,eachModelBom[2],partName,deptName,partSum,
                                                            eachModelType[2]*perNum,eachModelType[3]*perNum,eachModelType[4]*perNum,
                                                            componentItem,0,alias,stockIn+stockPT,stockZZ,stockIn+stockPT,tableDataStock[i][5]];
                                                        computResult.push(perPartRequest);
                                                        break;
                                                    }
                                                }
                                            }
                                        }   
                                    });
                                });
                                testResult.push(["1.型号分解"]);
                                computResult.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });

                                //分解组件
                                //var computResult = [['0型号', '1图号','2名称',3'主制','4总计',5'1月',6'2月',7'3月',8'组件信息',9'是否拆解',10'别名',11'库存',12'在制']];
                                var compListAdd = [];
                                msgPercent = 11;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","正在分解组件",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                computResult.forEach((eachResult,numResult,arrResult)=>{
                                    

                                    if(eachResult[8].length!=0&&eachResult[9]==0){
                                        eachResult[8].forEach(eachPartBom=>{
                                            var perNum = 1;
                                            if(eachPartBom[1]!=null&&eachPartBom[1]!=0&&eachPartBom[1]!=""){
                                                perNum=eachPartBom[1];
                                            }
                                            for(i=0;i<tableDataStock.length;i++){
                                                if(tableDataStock[i][2]==eachPartBom[0]){
                                                    var modelName = tableDataStock[i][1];
                                                    var partName = tableDataStock[i][3];
                                                    var deptName = tableDataStock[i][4];
                                                    var partSum = 0;
                                                    // var partSum = eachResult[5]*perNum+eachResult[6]*perNum+eachResult[7]*perNum;
                                                    var alias = tableDataStock[i][9];
                                                    var component = 0;
                                                    if(tableDataStock[i][10]==1.0||tableDataStock[i][10]=="1.0"||tableDataStock[i][10]=="1"){
                                                        component=1;
                                                    }
                                                    var stockIn = 0;
                                                    if(tableDataStock[i][6]!=null&&tableDataStock[i][6]!=""&&tableDataStock[i][6]!=0){
                                                        stockIn=Number(tableDataStock[i][6]);
                                                    }
                                                    var stockPT = 0;
                                                    if(tableDataStock[i][7]!=null&&tableDataStock[i][7]!=""&&tableDataStock[i][7]!=0){
                                                        stockPT=Number(tableDataStock[i][7]);
                                                    }
                                                    var stockZZ = 0;
                                                    if(tableDataStock[i][8]!=null&&tableDataStock[i][8]!=""&&tableDataStock[i][8]!=0){
                                                        stockZZ=Number(tableDataStock[i][8]);
                                                    }
                                                    var componentItem = [];
                                                    if(component==1){//添加组件信息【下级号，需求数量】
                                                        tableDataPartsBom.forEach(eachPartBom0=>{
                                                            if(eachPartBom[0]==eachPartBom0[1]){
                                                                componentItem.push([eachPartBom0[4],eachPartBom0[9]]);
                                                            }
                                                        });
                                                    }
                                                    var perPartRequest = [modelName,eachPartBom[0],partName,deptName,partSum,
                                                    eachResult[5]*perNum,eachResult[6]*perNum,eachResult[7]*perNum,
                                                        componentItem,0,alias,stockIn+stockPT,stockZZ,stockIn+stockPT,tableDataStock[i][6]];
                                                        compListAdd.push(perPartRequest);
                                                    break;
                                                }
                                            }
                                        });    
                                    }
                                });
                                testResult.push(["2.组件初步分解","","","","","","","","","","","","",""]);
                                compListAdd.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });

                                msgPercent = 20;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","分解组件完成",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                //循环处理组件，不超过一定层数
                                msgPercent = 21;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","开始循环处理组件",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                for(i=0;i<errorFloor;i++){
                                    compListAdd.forEach((eachResult,numResult,arrResult)=>{
                                        if(eachResult[8].length!=0&&eachResult[9]==0){
                                            eachResult[8].forEach(eachPartBom=>{
                                                var perNum = 1;
                                                if(eachPartBom[1]!=null&&eachPartBom[1]!=0&&eachPartBom[1]!=""){
                                                    perNum=eachPartBom[1];
                                                }
                                                for(t=0;t<tableDataStock.length;t++){
                                                    if(tableDataStock[t][2]==eachPartBom[0]){
                                                        var modelName = tableDataStock[t][1];
                                                        var partName = tableDataStock[t][3];
                                                        var deptName = tableDataStock[t][4];
                                                        var partSum = 0;
                                                        // var partSum = eachResult[5]*perNum+eachResult[6]*perNum+eachResult[7]*perNum;
                                                        var alias = tableDataStock[t][9];
                                                        var component = 0;
                                                        if(tableDataStock[t][10]==1.0||tableDataStock[t][10]=="1.0"||tableDataStock[t][10]=="1"){
                                                            component=1;
                                                        }
                                                        var stockIn = 0;
                                                        if(tableDataStock[t][6]!=null&&tableDataStock[t][6]!=""&&tableDataStock[t][6]!=0){
                                                            stockIn=Number(tableDataStock[t][6]);
                                                        }
                                                        var stockPT = 0;
                                                        if(tableDataStock[t][7]!=null&&tableDataStock[t][7]!=""&&tableDataStock[t][7]!=0){
                                                            stockPT=Number(tableDataStock[t][7]);
                                                        }
                                                        var stockZZ = 0;
                                                        if(tableDataStock[t][8]!=null&&tableDataStock[t][8]!=""&&tableDataStock[t][8]!=0){
                                                            stockZZ=Number(tableDataStock[t][8]);
                                                        }
                                                        var componentItem = [];
                                                        if(component==1){//添加组件信息【下级号，需求数量】
                                                            if(tableDataPartsBom.length!=0){
                                                                tableDataPartsBom.forEach(eachPartBom1=>{
                                                                    if(eachPartBom1[1]==eachPartBom[0]){
                                                                        componentItem.push([eachPartBom1[4],eachPartBom1[9]]);
                                                                    }
                                                                });
                                                            }

                                                        }
                                                        var perPartRequest = [modelName,eachPartBom[0],partName,deptName,partSum,
                                                            eachResult[5]*perNum,eachResult[6]*perNum,eachResult[7]*perNum,
                                                            componentItem,0,alias,stockIn+stockPT,stockZZ,stockIn+stockPT,tableDataStock[t][6]];
                                                        compListAdd.push(perPartRequest);
                                                        break;
                                                    }
                                                }
                                            });
                                            arrResult[numResult][9]=1;
                                        }
                                    });
                                }
                                msgPercent = 25;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","循环处理组件完成，开始合并需求",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                testResult.push(["3.组件循环分解","","","","","","","","","","","","",""]);
                                compListAdd.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                // $print($ArrayToString(computResult));

                                //合并需求
                                compListAdd.forEach((eachAdd,numAdd,arrAdd)=>{
                                    var isHave = 0;
                                    for(i=0;i<computResult.length;i++){
                                        if(eachAdd[1]==computResult[i][1]){
                                            computResult[i][5]+=eachAdd[5];
                                            computResult[i][6]+=eachAdd[6];
                                            computResult[i][7]+=eachAdd[7];
                                            
                                            // 使用 splice 删除指定索引的元素，并将其存储在变量中
                                            const [element] = computResult.splice(i, 1);
                                            // 使用 push 将元素添加到数组末尾
                                            computResult.push(element);

                                            isHave = 1;
                                            break;
                                        }
                                    }
                                    if(isHave==0){
                                        computResult.push(eachAdd);
                                    }
                                });
                                msgPercent = 30;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","合并需求完成，开始反冲组件资源",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

	                            testResult.push(["4.组件需求加和","","","","","","","","","","","","",""]);
                                computResult.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });

                                //处理库存反冲
                                //var computResult = [['0型号', '1图号','2名称',3'主制','4总计',5'1月',6'2月',7'3月',8'组件信息',9'是否拆解',10'别名',11'库存',12'在制']];
                                var compListSub = [];
                                computResult.forEach((eachResult,numResult,arrResult)=>{
                                    if(!Array.isArray(eachResult[8])){
                                        return;
                                    }
                                    if(eachResult[8].length>0){
										compListSub = [];
										eachResult[8].forEach(eachPartBom=>{
											var perNum = 1;
											if(eachPartBom[1]!=null&&eachPartBom[1]!=0&&eachPartBom[1]!=""){
												perNum=eachPartBom[1];
											}
											for(i=0;i<tableDataStock.length;i++){
												if(tableDataStock[i][2]==eachPartBom[0]){
													var modelName = tableDataStock[i][1];
													var partName = tableDataStock[i][3];
													var deptName = tableDataStock[i][4];
													var partSum = 0;
													// var partSum = eachResult[5]*perNum+eachResult[6]*perNum+eachResult[7]*perNum;
													var alias = tableDataStock[i][9];
													var component = 0;
													if(tableDataStock[i][10]==1.0||tableDataStock[i][10]=="1.0"||tableDataStock[i][10]=="1"){
														component=1;
													}
													var stockIn = 0;
													var stockPT = 0;
													var stockZZ = 0;
													var zzStock = Number(eachResult[11])+Number(eachResult[12]);

													var componentItem = [];
													if(component==1){//添加组件信息【下级号，需求数量】
														tableDataPartsBom.forEach(eachPartBom0=>{
															if(eachPartBom[0]==eachPartBom0[1]){
																componentItem.push([eachPartBom0[4],eachPartBom0[9]]);
															}
														});
													}
													var sub5 = 0;
													var sub6 = 0;
													var sub7 = 0;
													
													if(eachResult[5]>zzStock){
														sub5=zzStock*perNum;
														zzStock=0;
													}else{
														sub5=eachResult[5]*perNum;
														zzStock-=eachResult[5];
													}
													if(eachResult[6]>zzStock){
														sub6=zzStock*perNum;
														zzStock=0;
													}else{
														sub6=eachResult[6]*perNum;
														zzStock-=eachResult[6];
													}
													if(eachResult[7]>zzStock){
														sub7=zzStock*perNum;
														zzStock=0;
													}else{
														sub7=eachResult[7]*perNum;
														zzStock-=eachResult[7];
													}
													var perPartRequest = [modelName,eachPartBom[0],partName,deptName,partSum,
														sub5,sub6,sub7,componentItem,0,alias,stockIn+stockPT,stockZZ,stockIn+stockPT];
														compListSub.push(perPartRequest);
													break;
												}
												
											}
										
										});	
										var errorSub = 0; //待计算的总数减去当前计算的索引差
										const errorMax = 99;//待计算的总数减去当前计算的索引没有减少的总次数不超过99次
										var errorNum = 0;//待计算的总数减去当前计算的索引没有减少的总次数    
										var errStr = "";//零件信息记录
										//var computResult = [['0型号', '1图号','2名称',3'主制','4总计',5'1月',6'2月',7'3月',8'组件信息',9'是否拆解',10'别名',11'库存',12'在制']];

										for(j=0;j<compListSub.length;j++){
											if(compListSub.length-j>=errorSub){
												errorSub=compListSub.length-j;
												errorNum++;
												errStr += "->";
												errStr += compListSub[j][1];
											}else{
												errStr = "";
												errorNum = 0;
												errorSub = 0;
											}
											if(errorNum>errorMax){
												throw new Error('组件可能存在无限循环：' + errStr);
											}

											if(!Array.isArray(compListSub[j][8])){
												return;
											}
											if(compListSub[j][8].length>0){
												compListSub[j][8].forEach(eachPartBom=>{
													var perNum = 1;
													if(eachPartBom[1]!=null&&eachPartBom[1]!=0&&eachPartBom[1]!=""){
														perNum=eachPartBom[1];
													}
													for(t=0;t<tableDataStock.length;t++){
														if(tableDataStock[t][2]==eachPartBom[0]){
															var modelName = tableDataStock[t][1];
															var partName = tableDataStock[t][3];
															var deptName = tableDataStock[t][4];
															var partSum = 0;
															// var partSum = eachResult[5]*perNum+eachResult[6]*perNum+eachResult[7]*perNum;
															var alias = tableDataStock[t][9];
															var component = 0;
															if(tableDataStock[t][10]==1.0||tableDataStock[t][10]=="1.0"||tableDataStock[t][10]=="1"){
																component=1;
															}
															var stockIn = 0;
															var stockPT = 0;
															var stockZZ = 0;
															var componentItem = [];
															if(component==1){//添加组件信息【下级号，需求数量】
																if(tableDataPartsBom.length!=0){
																	tableDataPartsBom.forEach(eachPartBom1=>{
																		if(eachPartBom1[1]==eachPartBom[0]){
																			componentItem.push([eachPartBom1[4],eachPartBom1[9]]);
																		}
																	});
																}

															}
															var perPartRequest = [modelName,eachPartBom[0],partName,deptName,partSum,
																compListSub[j][5]*perNum,compListSub[j][6]*perNum,compListSub[j][7]*perNum,
																componentItem,0,alias,stockIn+stockPT,stockZZ,stockIn+stockPT];
															compListSub.splice(j+1, 0, perPartRequest);
															return;
														}
													}
												
												});	
											}											
                                        } 
										//库存反冲
										compListSub.forEach((eachSub,numSub,arrSub)=>{
											for(i=0;i<computResult.length;i++){
												if(eachSub[1]==computResult[i][1]){
													computResult[i][5]-=eachSub[5];
													computResult[i][6]-=eachSub[6];
													computResult[i][7]-=eachSub[7];
													break;
												}
											}
										});
                                    } 
                                });

                                msgPercent = 50;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","反冲组件完成，开始核减库存",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                testResult.push(["5.组件库存反冲","","","","","","","","","","","","",""]);
                                computResult.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                // $print($ArrayToString(computResult));
                                //处理本身库存
                                //考虑别名
                                //var computResult = [['0型号', '1图号','2名称',3'主制','4总计',5'1月',6'2月',7'3月',8'组件信息',9'是否拆解',10'别名',11'库存',12'在制',13'剩余']];
                                var aliases = [];
                                for(i=0;i<computResult.length;i++){
                                    
                                    if(i==0){
                                        continue;
                                    }
                                    if(aliases.indexOf(computResult[i][1])!=-1){
                                        computResult[i][4]=Number(computResult[i][5])+Number(computResult[i][6])+Number(computResult[i][7]);
                                        continue;
                                    }
                                    //查询别名库存
                                    var zzStock = Number(computResult[i][13]);
                                    //处理1月需求
                                    if(computResult[i][5]>zzStock){
                                        computResult[i][5]-=zzStock;
                                        zzStock=0;
                                    }else{
                                        zzStock-=computResult[i][5];
                                        computResult[i][5]=0;
                                    }
                                    // arrResult[numResult][13]=zzStock;//剩余库存
                                    if(computResult[i][10]!=null&&computResult[i][10]!=""){//有别名
                                        var alias = computResult[i][10];
                                        var aliasStock = 0;
                                        var partIndex = i;
                                        var partsn = computResult[i][1];
                                        for(j=0;j<errorFloor;j++){
                                            var aliasIndex = -1;
                                            for(t=0;t<computResult.length;t++){
                                                if(computResult[t][1]==alias){
                                                    aliasIndex=t;
                                                    aliasStock=Number(computResult[t][13]);
                                                    break;
                                                }
                                            }
                                            if(aliasIndex==-1){
                                                var modelname = '';
                                                var partname = '';
                                                var deptname = '';
                                                for(k=0;k<tableDataStock.length;k++){
                                                    // var dataGridColModel1 =  [
                                                    //     { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                                                    //     { label : '所属型号', name : 'belongModelSn',width: 150, align: 'left'},
                                                    //     { label : '图号', name : 'sn', width: 150, align: 'left'},
                                                    //     { label : '名称', name : 'name', width: 150, align: 'left'},
                                                    //     { label : '主制车间', name : 'deptname', width: 150, align: 'left'},
                                                    //     { label : '库房', name : 'storageCategory', width: 150, align: 'left'},
                                                    //     { label : '中央库库位库存', name : 'normalCount', width: 150, align: 'right'},
                                                    //     { label : '配套库存', name : 'ptsum', width: 150, align: 'right'},
                                                    //     { label : '组件在制', name : 'zzsum', width: 150, align: 'right'},
                                                    //     { label : '别名', name : 'alias', width: 150, align: 'left'},
                                                    //     { label : '是否组件', name : 'component', width: 150, align: 'left'},
                                                    //     { label : '当前时间', name : 'nowtime', width: 150, align: 'left'}
                                                    // ];
                                                    if(tableDataStock[k][2]==alias){
                                                        var stockIn = 0;
                                                        if(tableDataStock[k][6]!=null&&tableDataStock[k][6]!=""&&tableDataStock[k][6]!=0){
                                                            stockIn=tableDataStock[k][6]
                                                        }
                                                        var stockPT = 0;
                                                        if(tableDataStock[k][7]!=null&&tableDataStock[k][7]!=""&&tableDataStock[k][7]!=0){
                                                            stockPT=tableDataStock[k][7]
                                                        }
                                                        aliasStock=stockIn+stockPT;
                                                        modelname=tableDataStock[k][1];
                                                        partname=tableDataStock[k][3];
                                                        deptname=tableDataStock[k][4];
                                                        break;
                                                    }     
                                                }
                                                computResult.push([modelname, alias,partname,deptname,0,0,0,0,[],0,'',aliasStock,0,aliasStock]);
                                                aliasIndex = computResult.length-1;
                                            }
                                            //处理原件需求
                                            if(computResult[partIndex][5]>aliasStock){
                                                computResult[partIndex][5]-=aliasStock;
                                                aliasStock=0;
                                            }else{
                                                aliasStock-=computResult[partIndex][5];
                                                computResult[partIndex][5]=0;
                                            }
                                            //处理别名需求
                                            if(computResult[aliasIndex][5]>aliasStock){
                                                computResult[aliasIndex][5]-=aliasStock;
                                                aliasStock=0;
                                            }else{
                                                aliasStock-=computResult[aliasIndex][5];
                                                computResult[aliasIndex][5]=0;
                                            }
                                            computResult[aliasIndex][13] = aliasStock;//剩余可用库存更新

                                            aliases.push(alias);
                                            partIndex=aliasIndex;
                                            partsn=alias;
                                            alias=computResult[aliasIndex][10];
                                            if(alias==null||alias==""){
                                                break;
                                            }
                                        }
                                    }
                                    //处理2月需求
                                    if(computResult[i][6]>zzStock){
                                        computResult[i][6]-=zzStock;
                                        zzStock=0;
                                    }else{
                                        zzStock-=computResult[i][6];
                                        computResult[i][6]=0;
                                    }

                                    if(computResult[i][10]!=null&&computResult[i][10]!=""){//有别名
                                        var alias = computResult[i][10];
                                        var aliasStock = 0;
                                        var partIndex = i;
                                        var partsn = computResult[i][1];
                                        for(j=0;j<errorFloor;j++){
                                            var aliasIndex = -1;
                                            for(t=0;t<computResult.length;t++){
                                                if(computResult[t][1]==alias){
                                                    aliasIndex=t;
                                                    aliasStock=Number(computResult[t][13]);
                                                    break;
                                                }
                                            }
                                            //处理原件需求
                                            if(computResult[partIndex][6]>aliasStock){
                                                computResult[partIndex][6]-=aliasStock;
                                                aliasStock=0;
                                            }else{
                                                aliasStock-=computResult[partIndex][6];
                                                computResult[partIndex][6]=0;
                                            }
                                            //处理别名需求
                                            if(computResult[aliasIndex][6]>aliasStock){
                                                computResult[aliasIndex][6]-=aliasStock;
                                                aliasStock=0;
                                            }else{
                                                aliasStock-=computResult[aliasIndex][6];
                                                computResult[aliasIndex][6]=0;
                                            }
                                            computResult[aliasIndex][13] = aliasStock;//剩余可用库存更新

                                            aliases.push(alias);
                                            partIndex=aliasIndex;
                                            partsn=alias;
                                            alias=computResult[aliasIndex][10];
                                            if(alias==null||alias==""){
                                                break;
                                            }
                                        }
                                    }
                                    //处理3月需求    
                                    if(computResult[i][7]>zzStock){
                                        computResult[i][7]-=zzStock;
                                        zzStock=0;
                                    }else{
                                        zzStock-=computResult[i][7];
                                        computResult[i][7]=0;
                                    }    
                                    if(computResult[i][10]!=null&&computResult[i][10]!=""){//有别名
                                        var alias = computResult[i][10];
                                        var aliasStock = 0;
                                        var partIndex = i;
                                        var partsn = computResult[i][1];
                                        for(j=0;j<errorFloor;j++){
                                            var aliasIndex = -1;
                                            for(t=0;t<computResult.length;t++){
                                                if(computResult[t][1]==alias){
                                                    aliasIndex=t;
                                                    aliasStock=Number(computResult[t][13]);
                                                    break;
                                                }
                                            }
                                            //处理原件需求
                                            if(computResult[partIndex][7]>aliasStock){
                                                computResult[partIndex][7]-=aliasStock;
                                                aliasStock=0;
                                            }else{
                                                aliasStock-=computResult[partIndex][7];
                                                computResult[partIndex][7]=0;
                                            }
                                            //处理别名需求
                                            if(computResult[aliasIndex][7]>aliasStock){
                                                computResult[aliasIndex][7]-=aliasStock;
                                                aliasStock=0;
                                            }else{
                                                aliasStock-=computResult[aliasIndex][7];
                                                computResult[aliasIndex][7]=0;
                                            }
                                            computResult[aliasIndex][13] = aliasStock;//剩余可用库存更新

                                            aliases.push(alias);
                                            partIndex=aliasIndex;
                                            partsn=alias;
                                            alias=computResult[aliasIndex][10];
                                            if(alias==null||alias==""){
                                                break;
                                            }
                                        }
                                    }   
                                    computResult[i][13]=zzStock;
                                    computResult[i][14]=tableDataStock[t][5],
                                    computResult[i][4]=Number(computResult[i][5])+Number(computResult[i][6])+Number(computResult[i][7]);

                                }
                                msgPercent = 90;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","核减库存完成，汇总缺件结果",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                testResult.push(["6.库存别名核减","","","","","","","","","","","","",""]);
                                computResult.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
								
                                computResult.forEach((eachResult,numResult,arrResult)=>{
                                    if(numResult==0){
                                        return;
                                    }
                                    arrResult[numResult][8]=$ArrayToString(eachResult[8]);
                                });
                                // $print($ArrayToString(computResult));
                                msgPercent = 95;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","汇总缺件结果完成，过滤缺件",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                testResult.push(["7.初步结果","","","","","","","","","","","","",""]);
                                computResult.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                testResult.push(["8.过滤结果","","","","","","","","","","","","",""]);
                                computResult.filter(eachResult=>Number(eachResult[4])!=0).forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                msgPercent = 100;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","缺件结果计算完成，正在写入缺件结果和计算过程",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);
                                $SheetsLastAddWorkBook(workbookName,"计算过程");
                                Application.Worksheets.Item("计算过程").Activate();
                                $RangeInValue2ToSheetA1(workbookName,"计算过程",testResult);
                                $SheetsLastAddWorkBook(workbookName,"缺件结果");
                                Application.Worksheets.Item("缺件结果").Activate();
                                $RangeInValue2ToSheetA1(workbookName,"缺件结果",computResult.filter(eachResult=>Number(eachResult[4])!=0));

                                
                                alert("计算完成");
                            }).catch(error => {
                                console.error('调用 Web Stock服务失败:', error);
                                alert('调用 Web Stock服务失败,请关闭并重新计算:'+error);
                            });
                        }).catch(error => {
                            console.error('调用 Web ModelBom服务失败:', error);
                            alert('调用 Web ModelBom服务失败,请关闭并重新计算:'+error);
                        });
                    }).catch(error => {
                        console.error('调用 Web PartsBom服务失败:', error);
                        alert('调用 Web PartsBom服务失败,请关闭并重新计算:'+error);
                    });    
                }else{
                    alert("已取消");
                    return;
                } 
                break;
            }
            case "compute2": {
                // var isDo = $MsgBox("是否开始计算缺件2？可能需要一些时间");
                var isDo = $MsgBox("是否开始计算缺件？可能需要一些时间");


                if(isDo){
                    let doc = window.Application.ActiveWorkbook;
                    let textValue = "";
                    if (!doc) {
                        textValue = "当前没有打开任何文档";
                    } else {
                        textValue = doc.Name;
                    }
                    if ('缺件管理.xlsx' != textValue) {
                        alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                        return;
                    }
                    
                    let haveModelType = false;
                    for (let i = 1; i <= doc.Worksheets.Count; i++) {
                        if (doc.Worksheets.Item(i).Name == '型号需求') {
                            haveModelType = true;
                            break;
                        }
                    }
                    if (!haveModelType) {
                        alert("当前表不可用，请打开“缺件管理.xlsx”重试");
                        return;
                    }
                    let thisSheet = window.Application.ActiveWorkbook.ActiveSheet;
                    if (!haveModelType) {
                        alert("型号需求表不存在，需维护");
                        return;
                    }else{
                        var modelTypeSheet = window.Application.ActiveWorkbook.Worksheets.Item("型号需求");
                        const lastRow = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 1).End(-4162).Row;
                        if(lastRow<3){
                            alert("型号需求表无内容，需维护");
                            return; 
                        }
                    }

                    var workbookName = "缺件结果"+new Date().valueOf(); 
                    Application.Workbooks.Add();
                    window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
                    $SheetsLastAdd("计算进度");
                    Application.Worksheets.Item("计算进度").Activate();
                    var msgNum = 0;
                    var msgPercent = 1;
                    // var msgItem = [["序号2","进度2","进展2","时间2"],
                    var msgItem = [["序号","进度","进展","时间"],
                        [msgNum++,msgPercent.toFixed(2)+"%","校验文件完成，开始读取零组件BOM",$DateFormatDateTime(new Date())]];
                    $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);


                    var dataGridColModel =  [
                        { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                        { label : '上级号', name : 'partsSn',width: 150, align: 'left'},
                        { label : '上级名', name : 'partsSnName', width: 120, align: 'left'},
                        { label : '上级主制', name : 'partsSnDept', width: 100, align: 'left'},
                        { label : '下级号', name : 'bomPartsSn', width: 150, align: 'left'},
                        { label : '下级名', name : 'bomPartsSnName', width: 120, align: 'left'},
                        { label : '下级货位', name : 'positionNo', width: 120, align: 'left'},
                        { label : '下级主制', name : 'bomPartsDept', width: 100, align: 'left'},
                        { label : '需求类型', name : 'typeName', width: 80, align: 'center'},
                        { label : '需求数量', name : 'requireCount', width: 80, align: 'right'},
                        { label : '顺序号', name : 'sequence', width: 80, align: 'right'},
                        { label : '创建时间', name : 'createTime',  width: 120, align: 'center'}
                    ];
                    // var sheetName = "零件BOM" + +new Date().valueOf(); 
                    // $SheetsLastAdd(sheetName);

                    // 调用函数
                    // const url = "http://192.168.70.26:8080/V6R343/ws/dynPMResPartsBomWS0";
                    // const url = "http://192.168.70.17:8888/V6/ws/dynPMResPartsBomWS0";
                    // const method = "getPartsBom";
                    // const targetNamespace="http://ws.dynpmrespartsbom.planold.avicit/"
                    // const params = {};
                    selectServer().then(baseServerUrl => {
                        const url = baseServerUrl + "ws/dynPMResPartsBomWS0";
                        const method = "getPartsBom";
                        const targetNamespace = "http://ws.dynpmrespartsbom.planold.avicit/";
                        const params = {};
                        return $WebService(url, targetNamespace, method, params, dataGridColModel);
                    }).then(tableDataPartsBom => {
                        // console.log('返回的 JSON 字符串:', tableData);
                        // var arr = $Array2DFromJsonStr(jsonStr);
                        // $RangeInValue2("A1", $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@");
                        msgPercent = 2;
                        msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","开始读取型号BOM",$DateFormatDateTime(new Date())]);
                        $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                        var dataGridColModel0 =  [
                            { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                            { label : '型号', name : 'modelSn', width: 150, align: 'left'},
                            { label : '零组件号', name : 'partsSn', width: 150, align: 'left'},
                            { label : '零组件名', name : 'partsName', width: 150, align: 'left'},
                            { label : '货位', name : 'positionNo', width: 150, align: 'left'},
                            { label : '主制部门', name : 'partsDept', width: 150, align: 'left'},
                            { label : '需求类型', name : 'typeName', width: 150, align: 'center'},
                            { label : '单组数量', name : 'requireCount', width: 150, align: 'right'},
                            { label : '创建时间', name : 'createTime', width: 150, align: 'center'},
                            { label : 'BOM类型', name : 'categoryName', width: 150, align: 'center'}
                        ];
                        // var sheetName0 = "型号BOM" + +new Date().valueOf(); 
                        // $SheetsLastAdd(sheetName0);
                    
                        // 调用函数
                        // const url0 = "http://192.168.70.26:8080/V6R343/ws/dynPMResModelBomWS0";
                        // const url0 = "http://192.168.70.17:8888/V6/ws/dynPMResModelBomWS0";
                        // const method0 = "getModelBom";
                        // const targetNamespace0="http://ws.dynpmresmodelbom.planold.avicit/";
                        // const params0 = {};
                        selectServer().then(baseServerUrl => {
                            const url0 = baseServerUrl + "ws/dynPMResModelBomWS0";
                            const method0 = "getModelBom";
                            const targetNamespace0 = "http://ws.dynpmresmodelbom.planold.avicit/";
                            const params0 = {};
                            return $WebService(url0, targetNamespace0, method0, params0, dataGridColModel0);
                        }).then(tableDataModelBom => {
                            // console.log('返回的 JSON 字符串:', tableData);
                            // var arr = $Array2DFromJsonStr(jsonStr);
                            // $RangeInValue2("A1", $Array2DRowColCount(tableData0).rows,$Array2DRowColCount(tableData0).cols,tableData0,"@");
                            msgPercent = 3;
                            msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","开始读取物料与库存信息",$DateFormatDateTime(new Date())]);
                            $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                            var dataGridColModel1 =  [
                                { label : 'id', name : 'id', key : true, width : 75, hidden : true},
                                { label : '所属型号', name : 'belongModelSn',width: 150, align: 'left'},
                                { label : '图号', name : 'sn', width: 150, align: 'left'},
                                { label : '名称', name : 'name', width: 150, align: 'left'},
                                { label : '主制车间', name : 'deptname', width: 150, align: 'left'},
                                { label : '库房', name : 'storageCategory', width: 150, align: 'left'},
                                { label : '中央库库位库存', name : 'normalCount', width: 150, align: 'right'},
                                { label : '配套库存', name : 'ptsum', width: 150, align: 'right'},
                                { label : '组件在制', name : 'zzsum', width: 150, align: 'right'},
                                { label : '别名', name : 'alias', width: 150, align: 'left'},
                                { label : '是否组件', name : 'component', width: 150, align: 'left'},
                                { label : '当前时间', name : 'nowtime', width: 150, align: 'left'}
                            ];

                            const params1 = {};
                            const method1 = "getStocks";
                            // const url1 = "http://192.168.70.26:8080/V6R343/ws/dynStockPartsWS0";
                            // const url1 = "http://192.168.70.17:8888/V6/ws/dynStockPartsWS0";
                            // const targetNamespace1="http://ws.dynstockparts.planand.avicit/"
                            selectServer().then(baseServerUrl => {
                                const url1 = baseServerUrl + "ws/dynStockPartsWS0";
                                const targetNamespace1 = "http://ws.dynstockparts.planand.avicit/";
                                const params1 = {};
                                return $WebService(url1, targetNamespace1, method1, params1, dataGridColModel1);
                            }).then(tableDataStock => {
								// 删除重复声明的 errorFloor 常量
								
                                msgPercent = 5;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","读取物料与库存信息完成，开始整理型号需求",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                const errorFloor = 15; //最深计算层数
                                // 修复null错误：添加检查确保工作簿和工作表存在
                                var workbooks = window.Application.Workbooks;
                                if (!workbooks) {
                                    alert("无法访问工作簿集合");
                                    return;
                                }
                                
                                var 缺件管理Workbook = workbooks.Item("缺件管理.xlsx");
                                if (!缺件管理Workbook) {
                                    alert("无法找到'缺件管理.xlsx'工作簿");
                                    return;
                                }
                                
                                var modelTypeSheet = 缺件管理Workbook.Worksheets.Item("型号需求");
                                if (!modelTypeSheet) {
                                    alert("无法找到'型号需求'工作表");
                                    return;
                                }
                                
                                const lastRow0 = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 1).End(-4162).Row;
                                var modelTypeArray = modelTypeSheet.Range("A3:E"+lastRow0).Value2;
                                var modelTypeArray1 = [];
                                for (var [model, modelType,month1,month2,month3] of modelTypeArray) {
                                    if(month1==null||month1==""){
                                        month1=0;
                                    }
                                    if(month2==null||month2==""){
                                        month2=0;
                                    }
                                    if(month3==null||month3==""){
                                        month3=0;
                                    }
                                    modelTypeArray1.push([model, modelType,month1,month2,month3]);
                                }
                                var testResult = [];
                                msgPercent = 6;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","完成型号需求汇总，开始整理计算顺序",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);


                                testResult.push(["0型号需求0层"]);
                                modelTypeArray1.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                var computResult = [['型号', '图号','名称','主制','总计','1月','2月','3月','组件信息','是否拆解','别名','库存','在制','库房','别名库存']];
								var compListAdd0 = [];
								var compListAdd = [];
								
								//低阶码整理
								let partSnCodes = new Map();
								//逐型号计算
                                modelTypeArray1.forEach((eachModelType,num,modelTypeArr)=>{
                                    tableDataModelBom.forEach(eachModelBom=>{
                                        // var itPartBom = [];    
                                        if(eachModelBom[1]==eachModelType[0]){
                                            var perNum = 1;
                                            if(eachModelBom[7]!=null&&eachModelBom[7]!=0&&eachModelBom[7]!=""){
                                                perNum=eachModelBom[7];
                                            }
											for(i=0;i<tableDataStock.length;i++){
												if(tableDataStock[i][2]==eachModelBom[2]){
													var modelName = tableDataStock[i][1];
													var partName = tableDataStock[i][3];
													var deptName = tableDataStock[i][4];
													var partSum = 0;
													var alias = tableDataStock[i][9];
													var component = 0;
													if(tableDataStock[i][10]==1.0||tableDataStock[i][10]=="1.0"||tableDataStock[i][10]=="1"){
														component=1;
													}
													var stockIn = 0;
													if(tableDataStock[i][6]!=null&&tableDataStock[i][6]!=""&&tableDataStock[i][6]!=0){
														stockIn=Number(tableDataStock[i][6]);
													}
													var stockPT = 0;
													if(tableDataStock[i][7]!=null&&tableDataStock[i][7]!=""&&tableDataStock[i][7]!=0){
														stockPT=Number(tableDataStock[i][7]);
													}
													var stockZZ = 0;
													if(tableDataStock[i][8]!=null&&tableDataStock[i][8]!=""&&tableDataStock[i][8]!=0){
														stockZZ=Number(tableDataStock[i][8]);
													}
													var componentItem = [];
													if(component==1){//添加组件信息【下级号，需求数量】
														tableDataPartsBom.forEach(eachPartBom=>{
															if(eachPartBom[1]==eachModelBom[2]){
																componentItem.push([eachPartBom[4],eachPartBom[9]]);
															}
														});
													}
													var perPartRequest = [modelName,eachModelBom[2],partName,deptName,partSum,
														Number(eachModelType[2])*perNum,Number(eachModelType[3])*perNum,Number(eachModelType[4])*perNum,
														componentItem,0,alias,stockIn+stockPT,stockZZ,1];
													compListAdd0.push(perPartRequest);
													compListAdd.push(perPartRequest);
													break;
												}
											}
											partSnCodes.set(eachModelBom[2],1);
                                        }   
                                    });
                                });
                                testResult.push(["1.型号分解0-1层"]);
                                compListAdd.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });

                                //循环分解组件
                                //var computResult = [['0型号', '1图号','2名称',3'主制','4总计',5'1月',6'2月',7'3月',8'组件信息',9'是否拆解',10'别名',11'库存',12'在制']];
                                msgPercent = 11;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","正在分解组件1-n层",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                for(i=0;i<errorFloor;i++){
                                    compListAdd.forEach((eachResult,numResult,arrResult)=>{
                                        if(eachResult[8].length!=0&&eachResult[9]==0){
                                            eachResult[8].forEach(eachPartBom=>{
                                                var perNum = 1;
                                                if(eachPartBom[1]!=null&&eachPartBom[1]!=0&&eachPartBom[1]!=""){
                                                    perNum=eachPartBom[1];
                                                }
												var partSnCode = Number(eachResult[13])+1;
                                                for(t=0;t<tableDataStock.length;t++){
                                                    if(tableDataStock[t][2]==eachPartBom[0]){
                                                        var modelName = tableDataStock[t][1];
                                                        var partName = tableDataStock[t][3];
                                                        var deptName = tableDataStock[t][4];
                                                        var partSum = 0;
                                                        // var partSum = eachResult[5]*perNum+eachResult[6]*perNum+eachResult[7]*perNum;
                                                        var alias = tableDataStock[t][9];
                                                        var component = 0;
                                                        if(tableDataStock[t][10]==1.0||tableDataStock[t][10]=="1.0"||tableDataStock[t][10]=="1"){
                                                            component=1;
                                                        }
                                                        var stockIn = 0;
                                                        if(tableDataStock[t][6]!=null&&tableDataStock[t][6]!=""&&tableDataStock[t][6]!=0){
                                                            stockIn=Number(tableDataStock[t][6]);
                                                        }
                                                        var stockPT = 0;
                                                        if(tableDataStock[t][7]!=null&&tableDataStock[t][7]!=""&&tableDataStock[t][7]!=0){
                                                            stockPT=Number(tableDataStock[t][7]);
                                                        }
                                                        var stockZZ = 0;
                                                        if(tableDataStock[t][8]!=null&&tableDataStock[t][8]!=""&&tableDataStock[t][8]!=0){
                                                            stockZZ=Number(tableDataStock[t][8]);
                                                        }
                                                        var componentItem = [];
                                                        if(component==1){//添加组件信息【下级号，需求数量】
                                                            if(tableDataPartsBom.length!=0){
                                                                tableDataPartsBom.forEach(eachPartBom1=>{
                                                                    if(eachPartBom1[1]==eachPartBom[0]){
                                                                        componentItem.push([eachPartBom1[4],eachPartBom1[9]]);
                                                                    }
                                                                });
                                                            }
                                                        }
                                                        var perPartRequest = [modelName,eachPartBom[0],partName,deptName,partSum,
                                                            eachResult[5]*perNum,eachResult[6]*perNum,eachResult[7]*perNum,
                                                            componentItem,0,alias,stockIn+stockPT,stockZZ,partSnCode];
                                                        compListAdd.push(perPartRequest);
                                                        // compListAdd0.push(perPartRequest);
                                                        break;
                                                    }
                                                }
												partSnCodes.set(eachPartBom[0],partSnCode);
                                            });
                                            arrResult[numResult][9]=1;
                                        }
                                    });
                                }
                                msgPercent = 20;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","循环处理组件完成，开始处理计算顺序",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                testResult.push(["2.组件循环分解","","","","","","","","","","","","",""]);
                                compListAdd.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
								
                                // 将 Map 转换为数组，并根据 partSnCode 排序
								let sortedParts = Array.from(partSnCodes)
								  .sort((a, b) => a[1] - b[1]) // 按照 partSnCode (即每个元素的第二个位置) 排序
								  .map(item => item[0]); // 提取 partName

								msgPercent = 25;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","计算顺序处理完成，开始逐项计算",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);
                                $print($ArrayToString(sortedParts));
								//var computResult = [['0型号', '1图号','2名称',3'主制','4总计',5'1月',6'2月',7'3月',8'组件信息',9'是否拆解',10'别名',11'库存',12'在制']];
                                var partRequests = new Map();//净需求
                                var partRequests2 = new Map();//需下级件的需求

                                for(var month=5;month<=7;month++){
                                    msgPercent = 25+70*(month-4)/3;
                                    msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","正在计算"+(month-4)+"月需求缺件",$DateFormatDateTime(new Date())]);
                                    $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                    sortedParts.forEach((eachPart,partIdx,sortedPartsArr)=>{
                                        
                                        // msgPercent = 25+70*(partIdx+1)*(month-4)/sortedPartsArr.length*3;
                                        // msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","正在计算"+month+"月需求零件："+eachPart,$DateFormatDateTime(new Date())]);
                                        // if(msgNum%500==0){
                                        //     $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);
                                        // }
                                        
                                        for(t=0;t<tableDataStock.length;t++){
                                            if(tableDataStock[t][2]==eachPart){
                                                var modelName = tableDataStock[t][1];
                                                var partName = tableDataStock[t][3];
                                                var deptName = tableDataStock[t][4];
                                                var partSum = 0;
                                                var alias = tableDataStock[t][9];
                                                var component = 0;
                                                if(tableDataStock[t][10]==1.0||tableDataStock[t][10]=="1.0"||tableDataStock[t][10]=="1"){
                                                    component=1;
                                                }
                                                // var requests = [];
                                                // var requests2 = [];
                                                
                                                
                                                //读取库存
                                                var stockIn = 0;
                                                if(tableDataStock[t][6]!=null&&tableDataStock[t][6]!=""&&tableDataStock[t][6]!=0){
                                                    stockIn=Number(tableDataStock[t][6]);
                                                }
                                                var stockPT = 0;
                                                if(tableDataStock[t][7]!=null&&tableDataStock[t][7]!=""&&tableDataStock[t][7]!=0){
                                                    stockPT=Number(tableDataStock[t][7]);
                                                }
                                                var stockZZ = 0;
                                                if(tableDataStock[t][8]!=null&&tableDataStock[t][8]!=""&&tableDataStock[t][8]!=0){
                                                    stockZZ=Number(tableDataStock[t][8]);
                                                }
                                                var stockInAlias = 0;
                                                var stockPTAlias = 0;
                                                var stockZZAlias = 0;
                                                var aliasIndex = -1;
                                                if(alias!=null&&alias!=""){//有别名
                                                    for(ta=0;ta<tableDataStock.length;ta++){
                                                        if(tableDataStock[ta][2]==alias){
                                                            if(tableDataStock[ta][6]!=null&&tableDataStock[ta][6]!=""&&tableDataStock[ta][6]!=0){
                                                                stockInAlias=Number(tableDataStock[ta][6]);
                                                            }
                                                            if(tableDataStock[ta][7]!=null&&tableDataStock[ta][7]!=""&&tableDataStock[ta][7]!=0){
                                                                stockPTAlias=Number(tableDataStock[ta][7]);
                                                            }
                                                            if(tableDataStock[ta][8]!=null&&tableDataStock[ta][8]!=""&&tableDataStock[ta][8]!=0){
                                                                stockZZAlias=Number(tableDataStock[ta][8]);
                                                            }
                                                            aliasIndex=ta;
                                                            break;
                                                        }
                                                    }
                                                }
                                                //计算需求
                                                if(eachPart=='0164242700G'){
                                                    debugger;
                                                }
                                                //毛需求汇总
                                                var request = 0;
                                                compListAdd0.forEach(eachCompListAdd0=>{
                                                    if(eachCompListAdd0[1]==eachPart){
                                                        request+=eachCompListAdd0[month];
                                                    }
                                                });
                                                //需求核减库存
                                                if(request>tableDataStock[t][6]){
                                                    request-=tableDataStock[t][6];
                                                    tableDataStock[t][6]=0;
                                                }else{
                                                    tableDataStock[t][6]-=request;
                                                    request=0;
                                                }
                                                if(request>tableDataStock[t][7]){
                                                    request-=tableDataStock[t][7];
                                                    tableDataStock[t][7]=0;
                                                }else{
                                                    tableDataStock[t][7]-=request;
                                                    request=0;
                                                }
                                                if(aliasIndex!=-1){//有别名
                                                    if(request>tableDataStock[aliasIndex][6]){
                                                        request-=tableDataStock[aliasIndex][6];
                                                        tableDataStock[aliasIndex][6]=0;
                                                    }else{
                                                        tableDataStock[aliasIndex][6]-=request;
                                                        request=0;
                                                    }
                                                    if(request>tableDataStock[aliasIndex][7]){
                                                        request-=tableDataStock[aliasIndex][7];
                                                        tableDataStock[aliasIndex][7]=0;
                                                    }else{
                                                        tableDataStock[aliasIndex][7]-=request;
                                                        request=0;
                                                    }
                                                }
                                                // if(partRequests.has(eachPart)){
                                                //     partRequests.get(eachPart).set(month,request);
                                                // }else{
                                                //     var partRequestMonth = new Map();
                                                //     partRequestMonth.set(month,request);
                                                //     partRequests.set(eachPart,partRequestMonth);
                                                // }

                                                //需求核减在制
                                                var request2 = request;
                                                if(request2>tableDataStock[t][8]){
                                                    request2-=tableDataStock[t][8];
                                                    tableDataStock[t][8]=0;
                                                }else{
                                                    tableDataStock[t][8]-=request2;
                                                    request2=0;
                                                }
                                                if(aliasIndex!=-1){//有别名
                                                    if(request2>tableDataStock[aliasIndex][8]){
                                                        request2-=tableDataStock[aliasIndex][8];
                                                        tableDataStock[aliasIndex][8]=0;
                                                    }else{
                                                        tableDataStock[aliasIndex][8]-=request2;
                                                        request2=0;
                                                    }
                                                }

                                                // if(partRequests2.has(eachPart)){
                                                //     partRequests2.get(eachPart).set(month,request2);
                                                // }else{
                                                //     var partRequestMonth = new Map();
                                                //     partRequestMonth.set(month,request2);
                                                //     partRequests2.set(eachPart,partRequestMonth);
                                                // }
                                                
                                                
                                                //下级需求展开，增加毛需求
                                                var componentItem = [];
                                                if(component==1){//添加组件信息【下级号，需求数量】
                                                    if(tableDataPartsBom.length!=0){
                                                        tableDataPartsBom.forEach(eachPartBom1=>{
                                                            if(eachPartBom1[1]==eachPart){
                                                                componentItem.push([eachPartBom1[4],eachPartBom1[9]]);
                                                                var perNum = 1;
                                                                if(eachPartBom1[9]!=null&&eachPartBom1[9]!=0&&eachPartBom1[9]!=""){
                                                                    perNum=Number(eachPartBom1[9]);
                                                                }

                                                                var perPartRequest=[];
                                                                switch(month){
                                                                    case(5):{
                                                                        perPartRequest = [eachPart,eachPartBom1[4],"","",0,
                                                                            request2*perNum,0,0,"",0,"",0,0,0];
                                                                        break;
                                                                    }
                                                                    case(6):{
                                                                        perPartRequest = [eachPart,eachPartBom1[4],"","",0,
                                                                            0,request2*perNum,0,"",0,"",0,0,0];
                                                                        break;
                                                                    }
                                                                    case(7):{
                                                                        perPartRequest = [eachPart,eachPartBom1[4],"","",0,
                                                                            0,0,request2*perNum,"",0,"",0,0,0];
                                                                        break;
                                                                    }
                                                                }
                                                                compListAdd0.push(perPartRequest);
                                                            }
                                                        });
                                                    }
                                                }
                                                switch(month){
                                                    case(5):{
                                                        var perPartRequest = [modelName,eachPart,partName,deptName,0,
                                                            request,0,0,componentItem,0,alias,stockIn+stockPT,stockZZ,
                                                            tableDataStock[t][5],alias+":"+stockInAlias+"+"+stockPTAlias+"+"+stockZZAlias];
                                                        computResult.push(perPartRequest);
                                                        break;
                                                    }
                                                    case(6):{
                                                        for(var i=0;i<computResult.length;i++){
                                                            if(computResult[i][1]==eachPart){
                                                                computResult[i][6]=request;
                                                                break;
                                                            }
                                                        }
                                                        break;
                                                    }
                                                    case(7):{
                                                        for(var i=0;i<computResult.length;i++){
                                                            if(computResult[i][1]==eachPart){
                                                                computResult[i][7]=request;
                                                                computResult[i][4]=computResult[i][5]+computResult[i][6]+computResult[i][7];
                                                                break;
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    });
                                }
                                msgPercent = 96;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","核减库存完成，汇总缺件结果",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                computResult.forEach((eachResult,numResult,arrResult)=>{
                                    if(numResult==0){
                                        return;
                                    }
                                    arrResult[numResult][8]=$ArrayToString(eachResult[8]);
                                });
								                                testResult.push(["3.初步结果","","","","","","","","","","","","",""]);
                                computResult.forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
								
								
                                msgPercent = 98;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","过滤缺件，写入缺件结果和计算过程",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);

                                testResult.push(["4.过滤结果","","","","","","","","","","","","",""]);
                                computResult.filter(eachResult=>Number(eachResult[4])!=0).forEach(eachMsg=>{
                                    testResult.push([$ArrayToString(eachMsg)]);
                                });
                                msgPercent = 100;
                                msgItem.splice(1, 0, [msgNum++,msgPercent.toFixed(2)+"%","缺件计算完成",$DateFormatDateTime(new Date())]);
                                $RangeInValue2ToSheetA1(workbookName,"计算进度",msgItem);
								
                                $SheetsLastAddWorkBook(workbookName,"计算过程");
                                Application.Worksheets.Item("计算过程").Activate();
                                $RangeInValue2ToSheetA1(workbookName,"计算过程",testResult);
								
                                $SheetsLastAddWorkBook(workbookName,"缺件结果");
                                Application.Worksheets.Item("缺件结果").Activate();
                                $RangeInValue2ToSheetA1(workbookName,"缺件结果",computResult.filter(eachResult=>Number(eachResult[4])!=0));
								
								
                                
                                alert("计算完成");
								
                            }).catch(error => {
                                console.error('调用 Web Stock服务失败:', error);
                                alert('调用 Web Stock服务失败,请关闭并重新计算:'+error);
                            });
                        }).catch(error => {
                            console.error('调用 Web ModelBom服务失败:', error);
                            alert('调用 Web ModelBom服务失败,请关闭并重新计算:'+error);
                        });
                    }).catch(error => {
                        console.error('调用 Web PartsBom服务失败:', error);
                        alert('调用 Web PartsBom服务失败,请关闭并重新计算:'+error);
                    });    
                }else{
                    alert("已取消");
                    return;
                } 
                break;
            }
        }
    } catch (ex) {
        $print("错误", ex.message);
        alert(ex.message);
    }
}

window.onload = () => {
    $print('start_btnCreateSOMSheet_js_onload');
};
