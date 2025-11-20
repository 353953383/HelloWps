function btnCreateInBatchMain(){
    var dataGridColModel =  [
		{ label : 'id', name : 'id', key : true, width : 75, hidden : true},
		{ label : '货位', name : 'position',width: 150, align: 'center'},
		{ label : '主制部门', name : 'deptname',width: 120, align: 'center'},
		{ label : '质保单号', name : 'inNumber',width: 120, align: 'left'},
		{ label : '所属型号', name : 'model', width: 120, align: 'left'},
		{ label : '零组件编号', name : 'partSn', width: 200, align: 'left'},
		{ label : '零组件名', name : 'partName', width: 150, align: 'left'},
		{ label : '批次号', name : 'batch', width: 120, align: 'left'},
		{ label : '军种', name : 'armyServices', width: 100, align: 'left'},
		{ label : '入库数量', name : 'inCount', width: 100, align: 'right'},
		{ label : '剩余数量', name : 'allowCount', width: 100, align: 'right'},
		{ label : '油封有效期', name : 'sealingExpiryDate', width: 150, align: 'center'},
		{ label : '保证有效期', name : 'guaranteeExpiryDate', width: 150, align: 'center'},
		{ label : '创建时间', name : 'createTime',  width: 200, align: 'center'},
		{ label : '创建人id', name : 'creatorId', width: 150, align: 'left', hidden : true},
		{ label : '创建人', name : 'creator', width: 150, align: 'left'},
		{ label : '备注', name : 'description', width: 150, align: 'left'}
	];

    try{


        // 获取今天的日期
        const today = new Date();
        var startDate = new Date();
        // 设置开始日期
        const examStartDate = new Date("2024/1/1 00:00:00");
        // 计算两个日期之间的毫秒差
        const timeDifference = today- examStartDate;
        // 将毫秒差转换为天数
        const daysDifference = Math.ceil(timeDifference / (1000 * 60 * 60 * 24));

        var isDo = $MsgBox("确认读取入库信息？");
        if(1==isDo){
            var dayIndex = $InputBox("输入要查询的日期至今天的天数，例如要查询2024/1/1至今天的入库记录，则输入:"+daysDifference,'输入天数',"1");
            if($NumberStrIs(dayIndex)&dayIndex>0){
                $print(dayIndex);
                $print(startDate);
                startDate.setDate(startDate.getDate()-Number(dayIndex));
                $print(startDate);
                var sheetName = $StringReplaceAll(startDate.toLocaleDateString(),"/",".") + 
                    "至" + $StringReplaceAll(today.toLocaleDateString(),"/",".") + "入库信息" + $StringReplaceAll(today.toLocaleTimeString(),":","");
                $SheetsLastAdd(sheetName);
            
                dayIndex = Math.floor(dayIndex);
                // 调用函数
                // const url = "http://192.168.70.26:8080/V6R343/ws/dynPMRepInBatchWS0";
                // const url = "http://192.168.70.17:8888/V6/ws/dynPMRepInBatchWS0";
                const method = "getInBatch";
                const targetNamespace="http://ws.dynpmrepinbatch.planold.avicit/"
                const params = {days:dayIndex};
                
                // 选择服务器
                selectServer().then(baseServerUrl => {
                    const url = baseServerUrl + "ws/dynPMRepInBatchWS0";
                    $WebService(url, targetNamespace, method, params,dataGridColModel).then(tableData => {
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
                alert("输入非有效数字"); 
            }
            
        }else{
            alert("已取消");
        }

        
      
    }catch(ex){
        $print("错误",ex);
        alert("错误:"+ex);
    }
    
}
function btnCreateInBatchHaveStockMain(){
    var dataGridColModel =  [
		{ label : 'id', name : 'id', key : true, width : 75, hidden : true},
		{ label : '货位', name : 'position',width: 150, align: 'center'},
		{ label : '主制部门', name : 'deptname',width: 120, align: 'center'},
		{ label : '质保单号', name : 'inNumber',width: 120, align: 'left'},
		{ label : '所属型号', name : 'model', width: 120, align: 'left'},
		{ label : '零组件编号', name : 'partSn', width: 200, align: 'left'},
		{ label : '零组件名', name : 'partName', width: 150, align: 'left'},
		{ label : '批次号', name : 'batch', width: 120, align: 'left'},
		{ label : '军种', name : 'armyServices', width: 100, align: 'left'},
		{ label : '入库数量', name : 'inCount', width: 100, align: 'right'},
		{ label : '剩余数量', name : 'allowCount', width: 100, align: 'right'},
		{ label : '油封有效期', name : 'sealingExpiryDate', width: 150, align: 'center'},
		{ label : '保证有效期', name : 'guaranteeExpiryDate', width: 150, align: 'center'},
		{ label : '创建时间', name : 'createTime',  width: 200, align: 'center'},
		{ label : '创建人id', name : 'creatorId', width: 150, align: 'left', hidden : true},
		{ label : '创建人', name : 'creator', width: 150, align: 'left'},
		{ label : '备注', name : 'description', width: 150, align: 'left'}
	];

    try{
        // const url = "http://192.168.70.26:8080/V6R343/ws/dynPMRepInBatchWS0";
        // const url = "http://192.168.70.17:8888/V6/ws/dynPMRepInBatchWS0";
        const method = "getInBatchHaveStock";
        const targetNamespace="http://ws.dynpmrepinbatch.planold.avicit/"
        const params = {};
        var isDo = $MsgBox("确认读取批次库存详情信息？");
        if(1==isDo){
            // 选择服务器
            selectServer().then(baseServerUrl => {
                const url = baseServerUrl + "ws/dynPMRepInBatchWS0";
                $WebServiceToRangeA1("批次库存详情",url, targetNamespace, method, params, dataGridColModel);
            }).catch(error => {
                alert("操作已取消");
                return;
            });
        }else{
            alert("已取消");
        }
      
    }catch(ex){
        $print("错误",ex);
        alert("错误:"+ex);
    }
    
}