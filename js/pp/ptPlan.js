function onbuttonclick_ptPlan(idStr) {
    try {
        if (typeof (window.Application.Enum) != "object") { // 如果没有内置枚举值
            window.Application.Enum = WPS_Enum;
        }
        switch (idStr) {
            case "dockLeft":
                $print('dockLeft start');
                let tsIdLeft = window.Application.PluginStorage.getItem("ptPlan_id");
                if (tsIdLeft) {
                    let tskpane = window.Application.GetTaskPane(tsIdLeft);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionLeft;
                }
                break;
            case "dockRight":
                $print('dockRight start');
                let tsIdRight = window.Application.PluginStorage.getItem("ptPlan_id");
                if (tsIdRight) {
                    let tskpane = window.Application.GetTaskPane(tsIdRight);
                    tskpane.DockPosition = window.Application.Enum.msoCTPDockPositionRight;
                }
                break;
            case "hideTaskPane":
                $print('hideTaskPane start');
                let tsIdHide = window.Application.PluginStorage.getItem("ptPlan_id");
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
                document.getElementById("isPTPlan").innerHTML = textValue
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
                if ('配套计划协同.xlsx' != textValue) {
                    alert("当前表不可用，请打开“配套计划协同.xlsx”重试");
                    return;
                }
                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号分工') {
                        Application.Worksheets.Item(i).Activate();
                        haveModelType = true;
                        break;
                    }
                }
                $print("是否有型号分工表",haveModelType);
                if (!haveModelType) {
                    if (1 == $MsgBox("没有型号分工表，是否新建？")) {
                        $SheetsLastAdd('型号分工');
                        Application.Worksheets.Item('型号分工').Activate();
                        let titles = [['型号', '批产/科研类型', '主管计划员', '排序号'], ['填写型号名', '填写 批产/科研', '','例如：10001']];
                        $RangeInValue2("A1", 2, 4, titles,"@",'型号分工');
                    } else {
                        alert('已取消');
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
                if ('配套计划协同.xlsx' != textValue) {
                    alert("当前表不可用，请打开“配套计划协同.xlsx”重试");
                    return;
                }

                if (1 == $MsgBox("确认重置配套计划草案？")) {
                    for (let i = 1; i <= Application.Worksheets.Count; i++) {
                        if (Application.Worksheets.Item(i).Name == "配套计划草案"){
                            Application.Worksheets.Item(i).Delete()
                            break;
                        }
                    }
                    $SheetsLastAdd("配套计划草案");
                    Application.Worksheets.Item("配套计划草案").Activate();
                    
                    let titles = [["配套计划草案-汇总待检查",,'','','','', '','','','','','','','','','','','','','',
                        '','','','','','','','','','','','','','',''],
                    ['计划员','型号顺序','型号类型','车间','型号', '零件图号','零件名称','1月','2月','3月','4月','5月','6月','7月','8月','9月',
                        '10月','11月','12月','1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月','总计','库房','备注'], 
                    ['前三列不维护，所有数据从第四行D列开始维护，不要删除表头和插入空行，车间仅限于一车间,二车间,三车间,六车间,八车间,十车间,十一车间,采购与供应链管理部，优异中心',
                        '','','','', '','','','','','','','','','','','','','','','','','','','','','','','','','','','','']];
                    $RangeInValue2("A1", 3, 34, titles,"@","配套计划草案");
                    Application.Worksheets.Item("配套计划草案").Range("A1:AH1").Merge();
                } else {
                    alert('已取消');
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
                if ('配套计划协同.xlsx' != textValue) {
                    alert("当前表不可用，请打开“配套计划协同.xlsx”重试");
                    return;
                }
                let thisSheet = window.Application.ActiveWorkbook.ActiveSheet

                if (!thisSheet.Name== "配套计划草案") {
                    alert("工作表名称格式不符,请选择表“配套计划草案”");
                    return;
                }else{
                    $print(thisSheet.Name,"表名验证完成");
                }
                if(thisSheet.Range("A2").Value2!="计划员"||thisSheet.Range("B2").Value2!="型号顺序"||thisSheet.Range("E2").Value2!="型号"){
                    $print("A2\B2\E2为计划员\型号顺序\型号");
                    alert("工作表格式不符");
                    return;
                }else{
                    $print(thisSheet.Name,"表列验证完成");
                }

                let haveModelType = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == '型号分工') {
                        haveModelType = true;
                        break;
                    }
                }
                if (!haveModelType) {
                    alert("型号分工表不存在，需维护");
                    return;
                }else{
                    var modelTypeSheet = window.Application.ActiveWorkbook.Worksheets.Item("型号分工");
                    const lastRow = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 2).End(-4162).Row;
                    if(lastRow<3){
                        alert("型号属性表无内容，需维护");
                        return; 
                    }
                }
                var typeIsOK = true;//型号是否全部正确匹配到了类型
                // 获取 B 列的最后一行
                const lastRow = thisSheet.Cells.Item(thisSheet.Rows.Count, 6).End(-4162).Row; // -4162 表示向上查找最后一个非空单元格
                if(lastRow>3){
                    var modelArray = [];
                    if(lastRow==4){
                        modelArray.push([thisSheet.Range("E4").Value2]);
                    }else{
                        modelArray = thisSheet.Range("E4:E"+lastRow).Value2;
                    }
                    var modelTypeSheet = window.Application.ActiveWorkbook.Worksheets.Item("型号分工");
                    const lastRow0 = modelTypeSheet.Cells.Item(thisSheet.Rows.Count, 2).End(-4162).Row;
                    var aColumn = modelTypeSheet.Range("A3:A" + lastRow0).Value2;
                    var bColumn = modelTypeSheet.Range("B3:B" + lastRow0).Value2;
                    var cColumn = modelTypeSheet.Range("C3:C" + lastRow0).Value2;
                    var dColumn = modelTypeSheet.Range("D3:D" + lastRow0).Value2;

                    const modelTypeMap = new Map();
                    const modelPlannerMap = new Map();
                    const modelNumMap = new Map();
                    for (let i = 0; i < aColumn.length; i++) {
                        const key = aColumn[i][0]; // A列第i行的值（注意：Value2返回的是二维数组，格式为 [行][列]）
                        const valueB = bColumn[i][0]; // B列第i行的值
                        const valueC = cColumn[i][0]; // C列第i行的值
                        const valueD = dColumn[i][0]; // D列第i行的值
                        modelTypeMap.set(key, valueB);
                        modelPlannerMap.set(key, valueC);
                        modelNumMap.set(key, valueD);
                    }
                    $print("modelArray",$ArrayToString(modelArray));
                    var typeArray = [];
                    var plannerArray = [];
                    var numArray = [];
                    var esgString = "";
                    for(var i=0;i<modelArray.length;i++){
                        var plannerVal = modelPlannerMap.get(modelArray[i][0]);
                        if(undefined===plannerVal||''==plannerVal||plannerVal==null){
                            plannerArray.push("无数据");
                            typeIsOK = false;
                            esgString += "第"+(i+4)+"行第A列型号分工属性不全\n";
                        }else{
                            plannerArray.push(plannerVal);
                        }
                        var numVal = modelNumMap.get(modelArray[i][0]);
                        if(undefined===numVal||''==numVal||numVal==null){
                            numArray.push("无数据");
                            typeIsOK = false;
                            esgString += "第"+(i+4)+"行第B列型号分工属性不全\n";
                        }else{
                            numArray.push(numVal);
                        }
                        var typeVal = modelTypeMap.get(modelArray[i][0]);
                        if(undefined===typeVal||''==typeVal||typeVal==null){
                            typeArray.push("无数据");
                            typeIsOK = false;
                            esgString += "第"+(i+4)+"行第C列型号分工属性不全\n";
                        }else{
                            typeArray.push(typeVal);
                        }
                    }
                    $print("typeArray",$ArrayToString(typeArray));
                    $print("plannerArray",$ArrayToString(plannerArray));
                    $print("numArray",$ArrayToString(numArray));
                    $RangeInValue2("C4", modelArray.length, 1, $ArrayTranspose(typeArray),"@",thisSheet.Name);
                    $RangeInValue2("A4", modelArray.length, 1, $ArrayTranspose(plannerArray),"@",thisSheet.Name);
                    $RangeInValue2("B4", modelArray.length, 1, $ArrayTranspose(numArray),"@",thisSheet.Name);
                }
                var valArray = thisSheet.Range("A4:AH"+lastRow).Value2;
                var zz = ['一车间','二车间','三车间','六车间','八车间','十车间','十一车间','采购与供应链管理部','优异中心']; //主制
                // var kf = ['中央库','采购库','成品库','一车间','二车间','三车间','六车间','八车间']; //库房
                valArray.forEach((item,row,arraySelf1) => {
                        if(item[3]==null||item[3]==""||!zz.includes(item[3])){
                            esgString+="第"+(4+row)+"行第D列主制错误\n";
                        }
                        if(item[4]==null||item[4]==""){
                            esgString+="第"+(4+row)+"行第E列型号为空\n";
                        }
                        if(item[5]==null||item[5]==""){
                            esgString+="第"+(4+row)+"行第F列零件图号为\n;";
                        }
                });

                alert("型号分工分配完成,"+(typeIsOK?"未识别分工或数据问题":"注意！！！\n"+esgString));

                
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
                if ('配套计划协同.xlsx' != textValue) {
                    alert("当前表不可用，请打开“配套计划协同.xlsx”重试");
                    return;
                }
                var selectedPlannerRadio = document.querySelector('input[name="planner"]:checked');
                var selectedPlanner = selectedPlannerRadio ? selectedPlannerRadio.value : null;
                if (!selectedPlanner) {
                    alert("请选择计划员");
                    return;
                }

                var haveSheet = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == selectedPlanner){
                        haveSheet = true;
                        break;
                    }
                }
                if (!haveSheet) {
                    alert("缺少表："+selectedPlanner);
                    return;
                }else{
                    Application.Worksheets.Item(selectedPlanner).Activate();
                }
                
               
                break;
            }
            case "creatOwnPlan":{
                let doc = window.Application.ActiveWorkbook;
                let textValue = "";
                if (!doc) {
                    textValue = "当前没有打开任何文档";
                } else {
                    textValue = doc.Name;
                }
                if ('配套计划协同.xlsx' != textValue) {
                    alert("当前表不可用，请打开“配套计划协同.xlsx”重试");
                    return;
                }
                var haveSheet = false;
                for (let i = 1; i <= Application.Worksheets.Count; i++) {
                    if (Application.Worksheets.Item(i).Name == "配套计划草案"){
                        haveSheet = true;
                        break;
                    }
                }
                if (!haveSheet) {
                    alert("缺少表“配套计划草案”");
                }
                
                var selectedPlannerRadio = document.querySelector('input[name="planner"]:checked');
                var planner = selectedPlannerRadio ? selectedPlannerRadio.value : null;
                // if (!planners.includes(planner)) {
                //     alert("计划员不存在");
                //     return;
                // }
                if(planner==null||planner==""){
                    alert("请选择计划员");
                    return;
                }
                if (1 == $MsgBox("确认重置计划员【"+planner+"】的配套计划？")) {
                    for (let i = 1; i <= Application.Worksheets.Count; i++) {
                        if (Application.Worksheets.Item(i).Name == planner){
                            Application.Worksheets.Item(i).Name = planner+$DateFormat(new Date(),"yyyyMMdd24hhmmss");
                            break;
                        }
                    }
                    $SheetsLastAdd(planner);
                    Application.Worksheets.Item(planner).Activate();
                    let thisSheet = window.Application.ActiveWorkbook.ActiveSheet
                    thisSheet.Range("A1").Value2 = "配套计划草案-"+planner;
                    var ptSheet = window.Application.ActiveWorkbook.Worksheets.Item("配套计划草案");
                    var lastRow = ptSheet.Cells.Item(thisSheet.Rows.Count, 6).End(-4162).Row; // -4162 表示向上查找最后一个非空单元格
                    var ptArray= [];
                    if(lastRow>3){
                        ptArray = ptSheet.Range("A4:AH"+lastRow).Value2;
                        ptArray = $ArrayByRegex(ptArray,0,planner);
                        ptArray = $Array2DSort(ptArray,5,"asc");
                        ptArray = $Array2DSort(ptArray,1,"asc");
                    }else{
                        alert("缺少数据，请先维护“配套计划草案”表");
                        return;
                    }
                    $RangeInValue2("A2", 2, 34, ptSheet.Range("A2:AH3").Value2,"@",planner);
                    $RangeInValue2("A4", $Array2DRowColCount(ptArray).rows, 34, ptArray,"@",planner);
                   //AI2 为 备注2
                   $RangeInValue2("AI2", 1, 1, ["备注2"],"@",planner);
                }
                alert("个人配套计划重置完成");
                break;
            }
            case "createPC":{
                var selectMonth = document.getElementById('monthSelect').value;
                createPT("批产",selectMonth);
                break;
            }
            case "createXP":{
                var selectMonth = document.getElementById('monthSelect').value;
                createPT("科研",selectMonth);
                break;
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

function createPT(planType,selectMonth){
    let doc = window.Application.ActiveWorkbook;
    let textValue = "";
    if (!doc) {
        textValue = "当前没有打开任何文档";
    } else {
        textValue = doc.Name;
    }
    if ('配套计划协同.xlsx' != textValue) {
        alert("当前表不可用，请打开“配套计划协同.xlsx”重试");
        return;
    }
    var planners = ['陆会罗','白旭','王海岩','朱志超','崔玉玺'];
    var ptValue = [];
    planners.forEach(planner=>{
        var haveSheetPlanners = "";
        for (let i = 1; i <= Application.Worksheets.Count; i++) {
            if (Application.Worksheets.Item(i).Name == planner){
                haveSheetPlanners += haveSheetPlanners==""?planner:","+planner;

                let thisSheet = Application.Worksheets.Item(i)
                var lastRow = thisSheet.Cells.Item(thisSheet.Rows.Count, 6).End(-4162).Row; // -4162 表示向上查找最后一个非空单元格
                var ptArray= [];
                if(lastRow>3){
                    ptArray = thisSheet.Range("A4:AI"+lastRow).Value2;
                }
                ptArray.forEach(ptArrayRow=>{
                    ptValue.push(ptArrayRow);
                });
                break;
            }
        }
    });
    if(ptValue.length==0){
        alert('无数据');
        return;
    }
    var selectYear = new Date().getFullYear();
    

    var delColumn = [];
    for(var i=1;i<selectMonth;i++){
        delColumn.push(6+i);
    }
    ptValue = $Array2DSort($Array2DDeleteCol($ArrayByRegex(ptValue,2,planType),delColumn),5,"asc");
    ptValue = $Array2DSort(ptValue,1,"asc");
    var row1Array = [['计划员','型号顺序','型号类型','车间','型号', '零件图号','零件名称','1月','2月','3月','4月','5月','6月','7月','8月','9月',
        '10月','11月','12月','1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月','总计','库房','备注','备注2']];

    row1Array0=$Array2DDeleteCol(row1Array,delColumn);
    row1Array=$Array2DDeleteCol($Array2DDeleteCol(row1Array,delColumn),[0,1,2]);
    var workbookName = selectYear + "年" + (Number(selectMonth)>12?Number(selectMonth)-12:selectMonth) +"月"+planType+"配套计划"+new Date().valueOf(); 


    Application.Workbooks.Add();
    window.Application.ActiveWorkbook.SaveAs(workbookName, null, null, null, null, null, null, 2, null, null, null, null);
    $SheetsLastAdd(planType+"配套计划汇总");
    Application.Worksheets.Item(planType+"配套计划汇总").Activate();


    $RangeInValue2("A1",1,$Array2DRowColCount(row1Array0).cols,row1Array0,"@",planType+"配套计划汇总");
    $RangeInValue2("A2",$Array2DRowColCount(ptValue).rows,$Array2DRowColCount(ptValue).cols,ptValue,"@",planType+"配套计划汇总");
    row1Array[0][0]='序号';
    var zz = ['一车间','二车间','三车间','六车间','八车间','十车间','十一车间','采购与供应链管理部','优异中心']; //主制
    zz.forEach(eachZZ=>{
        
        $SheetsLastAdd(eachZZ);
        Application.Worksheets.Item(eachZZ).Activate();
        $RangeInValue2("A1", 1, 1, [eachZZ+(Number(selectMonth)>12?Number(selectMonth)-12:selectMonth)+"月"+planType+"配套计划"],"@",eachZZ);
        Application.Worksheets.Item(eachZZ).Range("A1:AI1").Merge();
        Application.Worksheets.Item(eachZZ).Range("A1:AI1").HorizontalAlignment = -4108;
        Application.Worksheets.Item(eachZZ).Range("A1:AI1").Font.Name = "宋体";      // 设置字体名称为宋体
        Application.Worksheets.Item(eachZZ).Range("A1:AI1").Font.Bold = true;        // 设置字体加粗
        Application.Worksheets.Item(eachZZ).Range("A1:AI1").Font.Size = 18;          // 设置字体大小为18号
        $RangeInValue2("A2", 1, $Array2DRowColCount(row1Array).cols, row1Array,"@",eachZZ);
        Application.Worksheets.Item(eachZZ).Range("A2:AI2").HorizontalAlignment = -4108;
        Application.Worksheets.Item(eachZZ).Range("A2:AI2").Font.Name = "宋体";      // 设置字体名称为宋体
        Application.Worksheets.Item(eachZZ).Range("A2:AI2").Font.Bold = true;        // 设置字体加粗
        Application.Worksheets.Item(eachZZ).Range("A2:AI2").Font.Size = 12;          // 设置字体大小为12号
        Application.Worksheets.Item(eachZZ).Range("C:D").ColumnWidth = 20;
        eachZZArray = [];
        ptValue.forEach(eachValue=>{
            if(eachValue[3]==eachZZ){
                eachZZArray.push(eachValue.slice(3));
            }
        });
        $print($ArrayToString(eachZZArray));
        if(eachZZArray.length!=0){
            var indexNum = 1;
            eachZZArray.forEach(eachValue=>{
                if(eachValue[4]!=""&&eachValue[4]!=null){
                    eachValue[0]=indexNum;
                    indexNum++;
                }else{
                    eachValue[0]='';
                }
            });
        
            $RangeInValue2("A3", $Array2DRowColCount(eachZZArray).rows, $Array2DRowColCount(eachZZArray).cols,eachZZArray,"@",eachZZ);
            
        
            Application.Worksheets.Item(eachZZ).Range("A1").Resize(eachZZArray.length+2,$Array2DRowColCount(eachZZArray).cols).HorizontalAlignment = -4108;
            Application.Worksheets.Item(eachZZ).Range("A1").Resize(eachZZArray.length+2,$Array2DRowColCount(eachZZArray).cols).Borders.LineStyle = 1;
            Application.Worksheets.Item(eachZZ).Range("A1").Resize(eachZZArray.length+2,$Array2DRowColCount(eachZZArray).cols).Borders.Weight = 2 ;
        }
    });

}