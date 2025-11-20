function btnStockPartsMain(){
    var dataGridColModel =  [
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
    try{
        
    
        const today = new Date();
        const params = {};
        const method = "getStocks";
        // const url = "http://192.168.70.26:8080/V6R343/ws/dynStockPartsWS0";
        // 修改前（第23行附近）
        // const url = "http://192.168.70.17:8888/V6/ws/dynStockPartsWS0";
        
        // 修改后
        selectServer().then(baseServerUrl => {
            const url = baseServerUrl + "ws/dynStockPartsWS0";
            // 原有的$WebService调用逻辑
        });
        const targetNamespace="http://ws.dynstockparts.planand.avicit/"
    
        var isDo = $MsgBox("确认读取实时库存信息？");
            if(1==isDo){
                // 选择服务器
                selectServer().then(baseServerUrl => {
                    const url = baseServerUrl + "ws/dynStockPartsWS0";
                    var sheetName =$StringReplaceAll(today.toLocaleDateString(),"/",".") + "实时库存" + $StringReplaceAll(today.toLocaleTimeString(),":","");
                    $SheetsLastAdd(sheetName);
                    // 调用函数
                    $WebService(url, targetNamespace, method, params, dataGridColModel).then(tableData => {
                        // console.log('返回的 JSON 字符串:', tableData);
                        // var arr = $Array2DFromJsonStr(jsonStr);
                        $RangeInValue2("A1", $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@",sheetName);
    
                    }).catch(error => {
                        console.error('调用 Web 服务失败:', error);
                        alert('调用 Web 服务失败:'+error);
                    });
                }).catch(error => {
                    alert("操作已取消");
                    return;
                });
            }else{
                alert("已取消");
            }
    }catch(ex){
        $print("错误：",);
        alert(ex);
    }
    
}