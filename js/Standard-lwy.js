//Standard V0.0.1 liwenyu
//Standard_Application
/*
	InputBox快捷方法
*/
function $InputBox(...args){
	return Application.InputBox(...args);
}
/*
	弹出提示框，返回1确认，返回0取消
*/
function $MsgBox(msg){
	return Application.confirm(msg);
}

/**
 * 
 * @param {url} url 
 * @param {命名空间} targetNamespace 
 * @param {调用方法} method 
 * @param {参数json} params 
 * @param {表头映射数组} dataGridColModel 
 * @returns POST访问返回，处理为二维数组结果：
 * $WebService(url, targetNamespace, method, params, dataGridColModel).then(tableData => {

                    $RangeInValue2("A1", $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@");
    
                }).catch(error => {
                    console.error('调用 Web 服务失败:', error);
                });
 */
async function $WebService(url, targetNamespace, method, params, dataGridColModel) {
    try {
        // 动态生成请求参数
        let paramsString = '';
        for (const key in params) {
            if (params.hasOwnProperty(key)) {
                paramsString += `<${key}>${params[key]}</${key}>\n`;
            }
        }

        // 设置请求参数
        const requestBody = `<?xml version="1.0" encoding="UTF-8"?>
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tns="${targetNamespace}">
           <soapenv:Header/>
           <soapenv:Body>
              <tns:${method}>
                 ${paramsString}
              </tns:${method}>
           </soapenv:Body>
        </soapenv:Envelope>`;

        // 发起请求
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/xml;charset=UTF-8'
            },
            body: requestBody
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
        }

        const textResponse = await response.text(); // 获取响应文本

        // XML 解析与处理逻辑
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(textResponse, "text/xml");

        const parserError = xmlDoc.getElementsByTagName("parsererror").length > 0;
        if (parserError) {
            throw new Error('XML 解析错误: ' + textResponse);
        }

        const resultNodes = xmlDoc.getElementsByTagName("result");

        if (resultNodes.length > 0) {
            const results = Array.from(resultNodes).map(resultNode => {
                const dynPMRepInBatchDTO = {};
                for (let i = 0; i < resultNode.childNodes.length; i++) {
                    const childNode = resultNode.childNodes[i];
                    if (childNode.nodeType === Node.ELEMENT_NODE) {
                        dynPMRepInBatchDTO[childNode.tagName] = childNode.textContent;
                    }
                }
                return dynPMRepInBatchDTO;
            });

            // 找到包含最多子元素的结果节点
            const maxResult = results.reduce((max, current) => Object.keys(current).length > Object.keys(max).length ? current : max, {});

            // 使用包含最多子元素的结果节点的键作为表头
            const headers = Object.keys(maxResult);

            // 如果提供了 dataGridColModel，则创建一个从 name 到 label 的映射对象
            let headerLabels = headers;
            let orderedHeaders = headers;

            if (dataGridColModel && Array.isArray(dataGridColModel)) {
                const nameToLabelMap = dataGridColModel.reduce((map, col) => {
                    map[col.name] = col.label;
                    return map;
                }, {});

                // 按照 dataGridColModel 的顺序排列表头
                const definedHeaders = dataGridColModel.map(col => col.name);
                const undefinedHeaders = headers.filter(header => !definedHeaders.includes(header));

                orderedHeaders = [...definedHeaders, ...undefinedHeaders];
                headerLabels = orderedHeaders.map(name => nameToLabelMap[name] || name);
            }

            // 构建二维数组
            const tableData = [headerLabels].concat(results.map(item => orderedHeaders.map(key => item[key] || '')));

            return tableData;
        } else {
            throw new Error('无法找到结果节点');
        }

    } catch (error) {
        console.error('Error calling web service:', error.message);
        if (error instanceof SyntaxError) {
            console.error('解析响应时发生错误:', error);
			alert("解析响应时发生错误:" + error);
        } else if (error instanceof TypeError) {
            // 类型错误日志已简化
            // console.error('类型错误:', error);
			alert("类型错误:" + error);
        } else {
            // 未知错误日志已简化
            // console.error('未知错误:', error);
			alert("未知错误:" + error);
        }

        // console.error('请求体:', requestBody);

        if (error.response) {
            const textResponse = await error.response.text();
            // console.error('响应体:', textResponse);
            alert('响应体:'+textResponse);
        }
        alert("错误:" + error);
        alert("请求体:" + requestBody);
        alert("响应体:" + textResponse);
		
        return [];
    }
}
/*
	新建类型sheetNameType的工作表，调用WebService后将结果填充到A1单元格
*/
function $WebServiceToRangeA1(sheetNameType,url, targetNamespace, method, params, dataGridColModel) {
	const today = new Date();
	var sheetName =$StringReplaceAll(today.toLocaleDateString(),"/",".") + sheetNameType + $StringReplaceAll(today.toLocaleTimeString(),":","");
	$SheetsLastAdd(sheetName);
	// 调用函数
	$WebService(url, targetNamespace, method, params, dataGridColModel).then(tableData => {
		// console.log('返回的 JSON 字符串:', tableData);
		// var arr = $Array2DFromJsonStr(jsonStr);
		$RangeInValue2("A1", $Array2DRowColCount(tableData).rows,$Array2DRowColCount(tableData).cols,tableData,"@",sheetName);

	}).catch(error => {
		console.error('调用 Web 服务失败:', error);
		throw new Error(ex.message);
	});



}

//Standard_print
/*
	立即窗口打印
	*/
function $print(str1,str2){
    try{
	// 获取当前时间    
		const now = new Date();    
		const timestamp = $DateFormatDateTime(now);    
		let tempPath = Application.Env.GetHomePath();
		//tempPath 为C:/Users/win 要将 win 截取出来
		const userName = tempPath.substring(tempPath.lastIndexOf("/") + 1);

		// 构建日志内容    
		let logContent = "";
		let consoleContent = "";
		
		if(str2==null){
			consoleContent = str1;
			logContent = `[${timestamp}] [${userName}] ${str1}`;
		}else if(str1==null){
			return;
		}else{
			consoleContent = str1+" : "+str2;
			logContent = `[${timestamp}] [${userName}] ${str1} : ${str2}`;
		}

		// 控制台输出日志已简化
		console.log(logContent);  
		// 写入日志

		// document.getElementById("log").innerHTML = document.getElementById("log").innerHTML + "<br>"+ logContent ;
		// // 当前工作簿创建日志sheet，将日志写入sheet，若不存在则新建，若存在则在A列向下非空单元格继续记录
		// if(!$SheetsIsHave("日志")){
		// 	$SheetsLastAdd("日志"); 
		// }
		// $RangeInValue2ToSheetNewRow("日志",[[logContent]]);
		
		// alert(tempPath);
		if(!Application.FileSystem.existsSync(tempPath+'/wpslog')){
			Application.FileSystem.Mkdir(tempPath+'/wpslog'); 
		}
		if(!Application.FileSystem.Exists(tempPath+'/wpslog/'+$DateFormat(new Date(),"yyyyMMdd")+'.txt')){
			Application.FileSystem.WriteFile(tempPath+'/wpslog/'+$DateFormat(new Date(),"yyyyMMdd")+'.txt',''); 
		}
		Application.FileSystem.AppendFile(tempPath+'/wpslog/'+$DateFormat(new Date(),"yyyyMMdd")+'.txt',logContent+"\r\n");
		// if(userName!='win'){
		// 	let serverPath = '\\\\192.168.70.17/mes0/wpslogs';
		// 	if(!Application.FileSystem.Exists(serverPath+'/wpslog'+$DateFormat(new Date(),"yyyyMMdd24HH")+'.txt')){
		// 		Application.FileSystem.WriteFile(serverPath+'/wpslog'+$DateFormat(new Date(),"yyyyMMdd24HH")+'.txt','');
		// 	}
		// 	Application.FileSystem.AppendFile(serverPath+'/wpslog'+$DateFormat(new Date(),"yyyyMMdd24HH")+'.txt',logContent+"\r\n");
		// }
		
		//判断当前文档是否有sheet名为logging
		// if(!$SheetsIsHave("logging")){
		// 	$SheetsLastAdd("logging"); 
		// 	//将logging sheet设为隐藏
		// 	$SheetsVisible("logging",false);
		// }
		// $RangeInValue2ToSheetNewRow("logging",[[logContent]]);
		// $SheetsVisible("logging",false);
	}catch(ex){
		alert(ex.message);
	}
}




	
	/*
		立即窗口打印后，输出一定长度的“--”
		*/
	function $printBreak(str1,str2){
		$print(str1,str2);
		str1=String.valueOf(str1);
		str2=String.valueOf(str2);
		var len = 0;
		if(str1!=null){
			if(str1.length>30){
				len+=30;
			}else{
				len+=str1.length;
			}
		}
		if(str2!=null){
			if(str2.length>30){
				len+=30;
			}else{
				len+=str2.length;
			}
		}
		$print("--".repeat(len))
	}
	/*
		清屏,并输出
		*/
	function $printClear(str1,str2){
		console.clear();
		$print(str1,str2);
	}
	//Standard_Object
	/*
		获取对象的类型
		关键词：类型类型
		let num = 123;
		let str = "hello";
		let bool = true;
		let und = undefined;
		let nul = null;
		let obj = {};
		let arr = [];
		let func = function() {};
		
	console.log(getType(num));    // "Number"
	console.log(getType(str));    // "String"
	console.log(getType(bool));   // "Boolean"
	console.log(getType(und));    // "Undefined"
	console.log(getType(nul));    // "Null"
	console.log(getType(obj));    // "Object"
	console.log(getType(arr));    // "Array"
	console.log(getType(func));   // "Function"
		
			*/
	function $ObjectGetType(obj) {
		return Object.prototype.toString.call(obj).slice(8, -1);
	}
	//Standard_Date
	/*
		输出日期的标准时间格式: 24HH:MM:SS
		*/
	function $DateFormatTime(date){
		var d = new Date();
		d = date;
		var dTimeStr = d.toLocaleTimeString().slice(2);
		if("下午"==d.toLocaleTimeString().substring(0,2)){
			dTimeStr = d.getHours() + dTimeStr.substring(dTimeStr.indexOf(":"));
		}else if("上午12"==d.toLocaleTimeString().substring(0,4)){
			dTimeStr = dTimeStr.replace("12","00");
		}else{
			var dTimeStr = d.toLocaleTimeString();
		}
		return dTimeStr
	}
	
	/*
		输出日期的标准日期+时间格式: YYYY/MM/DD 24HH:MM:SS
		*/
	function $DateFormatDateTime(date){
		var d = new Date();
		d = date;
		var dDateStr = d.toLocaleDateString();
		return dDateStr + " " + $DateFormatTime(date);
	}

/*
		输出日期的标准日期格式: 年yyyy YYYY yy YY  月MM  日dd DD 时24HH 24hh 12HH 12hh 分mm 秒ss 毫秒SSS
		*/
		function $DateFormat(date,format){
			// 处理年份（yyyy/YYYY四位数，yy/YY两位数）
			const year = date.getFullYear().toString();
			format = format.replace(/yyyy|YYYY/g, year)
							.replace(/yy|YY/g, year.slice(-2));
		
			// 处理月份（MM补零两位）
			const month = (date.getMonth() + 1).toString().padStart(2, '0');
			format = format.replace(/MM/g, month);
		
			// 处理日期（dd/DD补零两位）
			const day = date.getDate().toString().padStart(2, '0');
			format = format.replace(/dd|DD/g, day);
		
			// 处理小时（24HH/24hh为24小时制，12HH/12hh为12小时制）
			const hours24 = date.getHours().toString();
			const hours12 = (hours24 % 12 || 12).toString(); // 12小时制转换
			format = format.replace(/24HH/g, hours24.padStart(2, '0'))
							.replace(/24hh/g, hours24)
							.replace(/12HH/g, hours12.padStart(2, '0'))
							.replace(/12hh/g, hours12);
		
			// 处理分钟（mm补零两位）
			const minutes = date.getMinutes().toString().padStart(2, '0');
			format = format.replace(/mm/g, minutes);
		
			// 处理秒（ss补零两位）
			const seconds = date.getSeconds().toString().padStart(2, '0');
			format = format.replace(/ss/g, seconds);
		
			// 处理毫秒（SSS补零三位）
			const milliseconds = date.getMilliseconds().toString().padStart(3, '0');
			format = format.replace(/SSS/g, milliseconds);
		
			return format;
		}




	/*
		输出无时间部分的日期
		*/
	function $DateNoTime(date){
		var d = new Date();
		d = date;
		var dateStr = $DateFormatDateTime(date);
		dateStr = dateStr.substring(0,dateStr.indexOf(" "));
		return new Date(dateStr);
	}
	/*
		输出正常习惯下用（年，月，日，时，分，秒）创建的Date
		*/
	function $DateWithYMDHMS(y,mm,d,h,m,s){
		if(y==null||!$strIsNumber(y)){
			y=0;
		}else if("Number"==$ObjectGetType(y)&&y<1000&&y>=0){
			y+=2000;
			y=Math.floor(y);
		}
		if(mm==null||!$strIsNumber(mm)){
			mm=0;
		}else if($strIsNumber(mm)){
			mm-=1;
		}
		if(d==null||!$strIsNumber(d)){
			d=1;
		}
		if(h==null||!$strIsNumber(h)){
			h=0;
		}
		if(m==null||!$strIsNumber(m)){
			m=0;
		}
		if(s==null||!$strIsNumber(s)){
			s=0;
		}
		return new Date(y,mm,d,h,m,s);
	}
	/*
		返回Date的星期几，mod：1：数字，2：星期几，3：几
		*/
	function $DateToWeekday(date,mod) {
		const weekdays1 = [7, 1, 2, 3, 4, 5, 6];
		const weekdays2 = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
		const weekdays3 = ['七', '一', '二', '三', '四', '五', '六'];
		const dayIndex = date.getDay(); // 获取星期几的数字
		if(null==mod){
			return weekdays1[dayIndex];
		}else if(2==mod||"2"==mod){
			return weekdays2[dayIndex]; // 返回对应的中文星期
		}else if(3==mod||"3"==mod){
			return weekdays3[dayIndex]; // 返回对应的中文星期
		}else{
			return weekdays1[dayIndex];
		}
		
	}

	function $DateExcelValueIsDate(value) {
		function isExcelDateSerial(value) {
			// 尝试将字符串转换为数字
			let num = Number(value);
		  
			// 检查是否为有效数字
			if (isNaN(num)) return false;
		  
			// 检查是否在合理范围内（1900年1月1日之后）
			let minExcelDate = 1; // Excel 日期序列号 1 对应 1900 年 1 月 1 日
			let maxExcelDate = 2958465; // 2958465 对应 9999 年 12 月 31 日
		  
			return num >= minExcelDate && num <= maxExcelDate;
		  }

		function isDateStringValid(dateString) {
			// 尝试使用多种日期格式进行解析
			let formats = [
			  'yyyy-MM-dd', // 2024-11-30
			  'MM/dd/yyyy', // 11/30/2024
			  'dd/MM/yyyy'  // 30/11/2024
			];
		  
			for (let format of formats) {
			  let date = parseDateString(dateString, format);
			  if (date && !isNaN(date.getTime())) {
				return true;
			  }
			}
		  
			return false;
		  }
		  
		  function parseDateString(dateString, format) {
			let parts = dateString.split(/[\-\/]/);
			let [year, month, day] = format === 'yyyy-MM-dd'
			  ? [parts[0], parts[1], parts[2]]
			  : format === 'MM/dd/yyyy'
				? [parts[2], parts[0], parts[1]]
				: [parts[2], parts[1], parts[0]];
		  
			// 将月份和日期转换为数字
			year = parseInt(year, 10);
			month = parseInt(month, 10) - 1; // 月份从0开始
			day = parseInt(day, 10);
		  
			// 创建 Date 对象
			let date = new Date(year, month, day);
		  
			// 检查是否为有效日期
			if (isNaN(date.getTime())) return null;
		  
			// 验证解析后的日期是否与原始输入匹配
			if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
			  return null;
			}
		  
			return date;
		  }





		// 检查是否为 Excel 序列号
		if (isExcelDateSerial(value)) {
		  return true;
		}
	  
		// 检查是否为有效的日期字符串
		if (isDateStringValid(value)) {
		  return true;
		}
	  
		return false;
	  }


	/*
		Excel日期转换为jsDate日期。不含时间
	*/
	function $DateExcelToJS(excelDate) {
		// 将 excelDate 转换为数字
		excelDate = Number(excelDate);
	  
		// 创建一个代表 1899 年 12 月 30 日的 Date 对象
		let baseDate = new Date(1899, 11, 30);
	  
		if (excelDate < 60) {
		  // 对于 1900 年 3 月 1 日之前的日期，直接计算
		  let date = new Date(baseDate.getTime() + excelDate * 24 * 60 * 60 * 1000);
		  // 设置时间为 08:00:00
		  date.setHours(8, 0, 0, 0);
		  return date;
		} else {
		  // 对于 1900 年 3 月 1 日及之后的日期，减去 1 天来修正 1900 年 2 月 29 日的错误
		  let date = new Date(baseDate.getTime() + (excelDate + 1) * 24 * 60 * 60 * 1000);
		  // 设置时间为 08:00:00
		  date.setHours(8, 0, 0, 0);
		  return date;
		}
	  }
	  /*
		jsDate日期转换为Excel日期。不含时间
	*/
	function $DateJsToExcel(jsDate) {
		// 创建一个只包含日期部分的Date对象（时间部分设置为00:00:00）
		let dateOnly = new Date(jsDate.getFullYear(), jsDate.getMonth(), jsDate.getDate());
	  
		// 创建一个代表1899年12月30日的Date对象
		let baseDate = new Date(1899, 11, 30);
	  
		// 计算两个日期之间的毫秒差
		let msDifference = dateOnly.getTime() - baseDate.getTime();
	  
		// 将毫秒差转换为天数，并加上1来得到Excel的序列号
		let excelDate = Math.floor(msDifference / (24 * 60 * 60 * 1000)) ;
	  
		// 处理1900年2月29日的错误
		if (dateOnly < new Date(1900, 2, 1)) {
		  // 如果日期在1900年3月1日之前，减去1天来修正1900年2月29日的错误
		  excelDate -= 1;
		}
	  
		return excelDate;
	  }

	  /*
		jsDate日期转换为Excel日期。包含时间
	*/
	function $DateTimeJsToExcel(jsDate) {
		// 创建一个代表1899年12月30日的Date对象
		let baseDate = new Date(1899, 11, 30);
	  
		// 计算两个日期之间的毫秒差
		let msDifference = jsDate.getTime() - baseDate.getTime();
	  
		// 将毫秒差转换为天数（包括小数部分），并加上1来得到Excel的序列号
		let excelDate = (msDifference / (24 * 60 * 60 * 1000)) 
	  
		// 处理1900年2月29日的错误
		if (jsDate < new Date(1900, 2, 1)) {
		  // 如果日期在1900年3月1日之前，减去1天来修正1900年2月29日的错误
		  excelDate -= 1;
		}
	  
		return excelDate;
	}

	//Standard_Number
	
	/*
		检查字符串是否是数字
		*/
	function $NumberStrIs(str) {
		return !isNaN(Number(str));
	}
	
	/*
		返回数字的精度（小数点后最大位数）
		*/
	function $NumberPrecision(num) {
	// 将数字转换为字符串
		let numStr = num.toString();
		// 使用正则表达式匹配小数部分的长度
		let match = numStr.match(/\.(\d+)/);
		if (match) {
			return match[1].length;
		} else {
			// 如果没有小数点，则精度为0
			return 0;
		}
	}
	/*
		四舍五入，参数：数字，最大位数
		*/
	function $NumberRound(num, decimalPlaces) {
		let factor = Math.pow(10, decimalPlaces);
		return parseFloat((Math.round(num * factor) / factor).toFixed(decimalPlaces));
	}
	
			
	//Standard_String
	
	/*
		检查字符串是否是数字
		*/
	function $StringIsNumber(str) {
		return !isNaN(Number(str));
	}
	/*
		字符串全部替换
	*/
	function $StringReplaceAll(str,oldStr,newStr) {
		return str.replace(new RegExp(oldStr, 'g'), newStr);
	}


	//Standard_Array
	
	/*
		创建数字数组，参数为起始、结束和偏移量
		*/
	function $ArrayGetNum(start,end,offset){
		if(null==start||null==end){
			return null;
		}else if("Number"!=$ObjectGetType(start)||"Number"!=$ObjectGetType(end)){
			return null;
		}
		if(null==offset){
			offset=1;
		}else if("Number"!=$ObjectGetType(offset)){
			offset=1;
		}
		if(start>end&&offset>0){
			return null;
		}else if(start<end&&offset<0){
			return null;
		}
		//判断最大精度
		var startPrecision = $numberPrecision(start);
		var endPrecision = $numberPrecision(end);
		var offsetPrecision = $numberPrecision(offset);
		var numberPrecision = Math.max(endPrecision,startPrecision,offsetPrecision);
		
		var itemCount = Math.floor((end-start)/offset)+1;
		var arrayStr = "[";
		if(start<=end){
			for(var i=start;i<=end;i+=offset){
				arrayStr += $numberRound(i,numberPrecision);
				arrayStr += ",";
			}
		}else{
			for(var i=start;i>=end;i+=offset){
				arrayStr += $numberRound(i,numberPrecision);
				arrayStr += ",";
			}
		}
		
		arrayStr.substring(0,arrayStr.length-1);
		arrayStr += "]";
		return eval(arrayStr);
	}
	/*
		输出数组字符串，带中括号和字符串的双引号，但不包含多维数组
		
		*/
	function $ArrayToString4(arr){
		if("Array"==$ObjectGetType(arr)){
			var arrStr = "[";
			for(var i=0;i<arr.length;i++){
				if("String"==$ObjectGetType(arr[i])){
					arrStr+="\"";
					arrStr+=arr[i].toString();
					arrStr+="\"";
				}else{
					arrStr+=arr[i];
				}
				if(i!=arr.length-1){
					arrStr+=",";
				}
			}
			arrStr+="]";
			return arrStr;
		}else{
			return null;
		}
	}
	/*
		输出数组字符串，带中括号和字符串的双引号，包括数组中的数组
		
		*/
	function $ArrayToString2(arr) {
		if ("Array" == $ObjectGetType(arr)) {
			var arrStr = "[";
			var stack = [{ array: arr, index: 0, parentStr: arrStr }];
	
			while (stack.length > 0) {
				var current = stack[stack.length - 1];
				var array = current.array;
				var index = current.index;
				var parentStr = current.parentStr;
	
				if (index < array.length) {
					if (index > 0 && array[index - 1] !== undefined) {
						parentStr += ",";
					}
	
					if (array[index] !== undefined) {
						if ("String" == $ObjectGetType(array[index])) {
							parentStr += `"${array[index]}"`;
						} else if ("Array" == $ObjectGetType(array[index])) {
							parentStr += "[";
							stack.push({ array: array[index], index: 0, parentStr: parentStr });
						} else {
							parentStr += array[index];
						}
					}
	
					current.index++;
					current.parentStr = parentStr;
				} else {
					if ("Array" == $ObjectGetType(array)) {
						parentStr += "]";
					}
					stack.pop();
					if (stack.length > 0) {
						stack[stack.length - 1].parentStr = parentStr;
					} else {
						arrStr = parentStr; // 更新最终的字符串
					}
				}
			}
	
			return arrStr.replace(/,\s*\]/g, ']'); // 移除最后一个逗号及其后面的空格
		} else {
			return null;
		}
	}
	/*
		输出数组字符串，带中括号和字符串的双引号，包括数组中的数组
		再此基础上，空值显示为空
		*/
	function $ArrayToString3(arr) {
		if ("Array" == $ObjectGetType(arr)) {
			var arrStr = "[";
			var stack = [{ array: arr, index: 0, parentStr: arrStr }];
	
			while (stack.length > 0) {
				var current = stack[stack.length - 1];
				var array = current.array;
				var index = current.index;
				var parentStr = current.parentStr;
	
				if (index < array.length) {
					if (index > 0) {
						parentStr += ",";
					}
	
					if (array[index] !== undefined) {
						if ("String" == $ObjectGetType(array[index])) {
							parentStr += `"${array[index]}"`;
						} else if ("Array" == $ObjectGetType(array[index])) {
							parentStr += "[";
							stack.push({ array: array[index], index: 0, parentStr: parentStr });
						} else {
							parentStr += array[index];
						}
					}
	
					current.index++;
					current.parentStr = parentStr;
				} else {
					if ("Array" == $ObjectGetType(array)) {
						parentStr += "]";
					}
					stack.pop();
					if (stack.length > 0) {
						stack[stack.length - 1].parentStr = parentStr;
					} else {
						arrStr = parentStr; // 更新最终的字符串
					}
				}
			}
	
			return arrStr;
		} else {
			return null;
		}
	}
	/*
		输出数组字符串，带中括号和字符串的双引号，包括数组中的数组
		再此基础上，空值显示为""，常用
		*/
	function $ArrayToString(arr) {
		if ("Array" == $ObjectGetType(arr)) {
			var arrStr = "[";
			var stack = [{ array: arr, index: 0, parentStr: arrStr }];
	
			while (stack.length > 0) {
				var current = stack[stack.length - 1];
				var array = current.array;
				var index = current.index;
				var parentStr = current.parentStr;
	
				if (index < array.length) {
					if (index > 0) {
						parentStr += ",";
					}
	
					if (array[index] === undefined) {
						parentStr += `""`; // 显示空字符串 ""
					} else {
						if ("String" == $ObjectGetType(array[index])) {
							parentStr += `"${array[index]}"`;
						} else if ("Array" == $ObjectGetType(array[index])) {
							parentStr += "[";
							stack.push({ array: array[index], index: 0, parentStr: parentStr });
						} else {
							parentStr += array[index];
						}
					}
	
					current.index++;
					current.parentStr = parentStr;
				} else {
					if ("Array" == $ObjectGetType(array)) {
						parentStr += "]";
					}
					stack.pop();
					if (stack.length > 0) {
						stack[stack.length - 1].parentStr = parentStr;
					} else {
						arrStr = parentStr; // 更新最终的字符串
					}
				}
			}
	
			return arrStr;
		} else {
			return null;
		}
	}
	/*
		一维数组和二维数组转置
		*/
	function $ArrayTranspose(matrix) {
		if (!Array.isArray(matrix) || matrix.length === 0) {
			return [];
		}
	
		// 检查输入是否为二维数组
		const is2DArray = matrix.every(row => Array.isArray(row));
	
		if (is2DArray) {
			// 获取原数组的行数和列数
			const rowCount = matrix.length;
			const colCount = matrix[0].length;
	
			// 初始化转置后的数组
			const transposed = new Array(colCount).fill(null).map(() => new Array(rowCount));
	
			// 填充转置后的数组
			for (let i = 0; i < rowCount; i++) {
				for (let j = 0; j < colCount; j++) {
					transposed[j][i] = matrix[i][j];
				}
			}
	
			return transposed;
		} else {
			// 处理一维数组的情况
			const length = matrix.length;
			const transposed = new Array(length).fill(null).map(() => [matrix[length - 1]]);
	
			// 填充转置后的数组
			for (let i = 0; i < length; i++) {
				transposed[i][0] = matrix[i];
			}
	
			return transposed;
		}
	}
/*
	将二维数组的json字符串转为二维数组
*/
function $Array2DFromJsonStr(jsonString) {
	try {
	  // 解析 JSON 字符串为 JavaScript 对象
	  const jsonArray = JSON.parse(jsonString);
  
	  // 检查是否为数组
	  if (!Array.isArray(jsonArray)) {
		throw new Error('输入的 JSON 字符串不是一个数组');
	  }
	  // 检查数组是否为空
	  if (jsonArray.length === 0) {
		throw new Error('输入的 JSON 数组为空');
	  }
	  // 提取表头
	  const headers = Object.keys(jsonArray[0]);
	  // 检查每个对象是否具有相同的键
	  jsonArray.forEach((item, index) => {
		const itemKeys = Object.keys(item);
		if (itemKeys.length !== headers.length || !headers.every(key => itemKeys.includes(key))) {
		  throw new Error(`第 ${index + 1} 个对象的键与表头不一致`);
		}
	  });
	  // 初始化二维数组
	  const result = [headers];
	  // 遍历 JSON 数组，将每个对象的值放入二维数组
	  jsonArray.forEach(item => {
		const row = [];
		headers.forEach(key => {
		  let value = item[key];
		  // 如果值是对象或数组，将其转换为字符串
		  if (typeof value === 'object' && value !== null) {
			value = JSON.stringify(value);
		  }
		  row.push(value);
		});
		result.push(row);
	  });
	  return result;
	} catch (ex) {
	  alert('解析 JSON 字符串时发生错误:', ex);
	  $print("错误", ex);
	  return [];
	}
  }
/*
	  用$Array2DRowColCount(array2D).rows,$Array2DRowColCount(array2D).cols,分别返回二维数组所代表的行列数。（对象数+1为行数（+标题），每个对象的键数位列数）
*/
	  function $Array2DRowColCount(array2D) {
		// 检查输入是否为数组
		if (!Array.isArray(array2D)) {
		  throw new Error('输入的不是二维数组');
		}
		// 检查二维数组的有效性
		if (array2D.length === 0 || !Array.isArray(array2D[0])) {
		  throw new Error('输入的不是有效的二维数组');
		}
		// 计算行数
		const numRows = array2D.length;
		// 计算列数（假设所有行的列数相同）
		const numCols = array2D[0].length;
		// 返回行数和列数
		return { rows: numRows, cols: numCols };
	  }
/*
	  用于对二维数组按指定列和顺序排序
*/
	function $Array2DSort(array, column, order = 'asc') {
		// 参数校验：确保column为有效索引
		if (typeof column !== 'number' || column < 0 || column >= array[0].length){
			alert(`无效的列索引：${column}`);
			return [...array]; // 返回原数组副本避免错误
		}
		
		// 参数校验：标准化order参数（支持大小写）
		const validOrders = ['asc', 'desc'];
		const normalizedOrder = order.toLowerCase();
		if (!validOrders.includes(normalizedOrder)) {
			alert(`无效的排序顺序：${order}，将使用默认升序`);
			return sort2DArray(array, column, 'asc');
		}

		// 复制原数组避免修改原始数据
		const sortedArray = array.map(row => [...row]);
		
		sortedArray.sort((a, b) => {
			const valA = a[column];
			const valB = b[column];

			// 处理数值类型
			if (typeof valA === 'number' && typeof valB === 'number') {
				return normalizedOrder === 'asc' ? valA - valB : valB - valA;
			}

			// 处理字符串类型（区分大小写）
			const strCompare = String(valA).localeCompare(String(valB));
			return normalizedOrder === 'asc' ? strCompare : -strCompare;

			// 可扩展其他类型（如日期）的排序逻辑...
		});

		return sortedArray;
	}
/*
	  删除二维数组指定列
*/

	function $Array2DDeleteCol(A, B) {
		// 处理边界情况：空输入
		if (!A || !B || A.length === 0 || B.length === 0) return A;
	
		// 列索引去重（避免重复删除同一列）
		const uniqueCols = [...new Set(B)];
	
		// 获取原始数组的列数（取第一行的长度作为基准）
		const originalCols = (A[0] && A[0].length) || 0;
	
		// 过滤无效的列索引（负数或超过最大列数的索引）
		const validCols = uniqueCols.filter(col => 
			Number.isInteger(col) && col >= 0 && col < originalCols
		);
	
		// 遍历每一行，过滤掉需要删除的列
		return A.map(row => {
			return row.filter((_, colIndex) => !validCols.includes(colIndex));
		});
	}



	/*
	  用于筛选二维数组指定列符合正则表达式
*/
function $ArrayByRegex(arr, column, regexStr) {
    // 参数校验：处理空输入
    if (!arr || !arr.length || column < 0) return [];
    
    // 参数校验：检查列是否存在（以第一行长度为基准）
    const maxColumn = (arr[0] && arr[0].length) || 0;
    if (column >= maxColumn) {
        alert(`错误：列索引${column}超出数组最大列数${maxColumn-1}`);
        return [];
    }

    try {
        // 编译正则表达式
        const regex = new RegExp(regexStr);
        
        // 遍历数组筛选符合条件的行
        return arr.filter(row => {
            // 处理行长度不足的情况（跳过该行）
            if (row.length <= column) return false;
            // 将列值转为字符串后匹配正则（避免数值类型无法匹配）
            const cellValue = String(row[column] || '');
            return regex.test(cellValue);
        });
    } catch (error) {
        alert(`正则表达式错误：${error.message}`);
        return [];
    }
}

	//Standard_Range
	
	/*
		字符串是否为有效的单元格或连续单元格范围
		*/
	function $RangeStringIsValid(rangeStr) {
		// 正则表达式匹配有效的单元格范围或单个单元格
		const rangePattern = /^([A-Z]+\d+)(:[A-Z]+\d+)?$/;
		// 检查范围是否符合模式
		return rangePattern.test(rangeStr);
	}	
	/*
		连续单元格范围字符串返回Range对象
	*/
	function $RangeOfString(rangeStr) {
		if($RangeStringIsValid(rangeStr)){
			return Application.ActiveSheet.Range(rangeStr);
		}else{
			return null;
		}
	}
	/*
		选中字符串代表的单元格或连续单元格
	*/
	function $RangeSelect(rangeStr){
		if($RangeStringIsValid(rangeStr)){
			$RangeOfString(rangeStr).Select();
		}
	}
	/*
		输出选中单元格Range
	*/
	function $RangeSelection(){
		var address = Application.Selection.Address();
		return Application.ActiveSheet.Range(address);
	}
	
	/*
		单元格范围有多少单元格
	*/
	function $RangeCellsCount(range){
		// 获取范围内的单元格数量
		return range.Cells.Count;
	}
	
	/*
		单个单元格是否为空
	*/
	function $RangeIsEmpty(range){
		if(null!=range&&1==$RangeCellsCount(range)){
			var type = $ObjectGetType(range.Value2)
			if(""==range.Value2 || "Undefined"==type){
				return true;
			}
		}
		return false;
	}
	/*
		单元格填充数组内容
	*/
	function $RangeInValue2(address,row,col,array,format,sheetName){
		var sheetItem = $SheetTheActive();
		if(sheetName!=""&&sheetName!=null){
			sheetItem = Application.ActiveWorkbook.Worksheets.Item(sheetName);
		}
		if(null!=format&&undefined!=format){
			sheetItem.Range(address).Resize(row,col).NumberFormatLocal = format;
		}
		sheetItem.Range(address).Resize(row,col).Value2=array;
		return true;
	}
	/*
		当前A1单元格填充数组内容
	*/
	function $RangeA1InValue2(array,sheetName){
		if(sheetName!=null&&sheetName!=""){
			$RangeInValue2("A1", $Array2DRowColCount(array).rows,$Array2DRowColCount(array).cols,array,"@",sheetName);
		}else{
			$RangeInValue2("A1", $Array2DRowColCount(array).rows,$Array2DRowColCount(array).cols,array,"@");
		}
		
		return true;
	}
	/*
		某表A列，在最下填充数组内容
	*/
	function $RangeInValue2ToSheetNewRow(sheetName,array){
		try{
			thisSheet = window.Application.ActiveWorkbook.Worksheets.Item(sheetName);
			const lastRow = thisSheet.Cells.Item(thisSheet.Rows.Count, 1).End(-4162).Row;
			$RangeInValue2("A"+(lastRow+1), $Array2DRowColCount(array).rows,$Array2DRowColCount(array).cols,array,"@",thisSheet.Name);
			return true;
		} catch (ex) {
			// $print("错误", ex.message);
			alert(ex.message);
			throw new Error(ex.message);
		}
		
	}
	/*
		某表A1填充数组内容
	*/
	function $RangeInValue2ToSheetA1(workbookName,sheetName,array){
		try{
			thisSheet = window.Application.Workbooks.Item(workbookName).Worksheets.Item(sheetName);
			$RangeInValue2("A1", $Array2DRowColCount(array).rows,$Array2DRowColCount(array).cols,array,"@",sheetName);
			return true;
		} catch (ex) {
			$print("错误", ex.message);
			alert(ex.message);
			throw new Error(ex.message);
		}
		
	}
/*
		某单元格的值
	*/
	function $RangeValue(range){
		//检查range是否为Range类型
		if(null==range){
			return null;
		}
		if("Object"!=$ObjectGetType(range)){
			$print($ObjectGetType(range));
			return null;
		}
		return range.Value();
	}

/*
		为某单元格写入值
	*/
function $RangeInValue(range,str){
		//检查range是否为Range类型
		if(null==range){
			return null;
		}
		if("Object"!=$ObjectGetType(range)){
			$print($ObjectGetType(range));
			return null;
		}
		if(null==str){
			return null;
		}
		if("String"!=$ObjectGetType(str)){
			$print($ObjectGetType(str));
			return null;
		}
		range.Value(null,str);
		return true;
	}
	
	//Standard_Sheet
	
	/*
		当前表字符串是否有相应名字的工作表Activate
		*/
	function $SheetsIsHave(sheetName){

		var sheetNames = [];
		for (var i = 1; i <= Application.Worksheets.Count; i++) {
			// if(null!=Application.Worksheets.Item(i)){
			// 	sheetNames.push(Application.Worksheets.Item(i).Name);
			// }
			sheetNames.push(Application.Worksheets.Item(i).Name);
		}
		
		// 检查指定的工作表是否存在于工作簿中
		return sheetNames.includes(sheetName);
	}
	/*
		激活当前表字符串相应名字的工作表
		*/
	function $SheetsActivate(sheetName){
		if($SheetsIsHave(sheetName)){
			Application.ActiveWorkbook.Worksheets.Item(sheetName).Activate();
		}
	}
	/*
		当前工作簿的工作表数量
	*/
	function $SheetsCount(){
		return Application.ActiveWorkbook.Worksheets.Count;
	}
	/*
		在最后创建新工作表
		*/
	function $SheetsLastAdd(sheetName){
		var beforeCount=$SheetsCount();
		Application.ActiveWorkbook.Worksheets.Add(undefined,Application.ActiveWorkbook.Worksheets.Item(beforeCount)).Name = sheetName;
		var afterCount=$SheetsCount();
		return afterCount==beforeCount+1;
	}
	/*
		在某工作表最后创建新工作表
		*/
		function $SheetsLastAddWorkBook(workbookName,sheetName){
			var beforeCount=$SheetsCount();
			Application.Workbooks.Item(workbookName).Worksheets.Add(undefined,Application.ActiveWorkbook.Worksheets.Item(beforeCount)).Name = sheetName;
			var afterCount=$SheetsCount();
			return afterCount==beforeCount+1;
		}
	/*
		当前工作表
	*/
	function $SheetTheActive(){
		return Application.ActiveWorkbook.ActiveSheet;
	}
	/*
		当前工作簿所有工作表名字数组
	*/
	function $SheetsNamesArray(){
		// 获取当前活动的工作簿
		var workbook = Application.ActiveWorkbook;
		// 获取所有工作表 
		var sheets = workbook.Worksheets;
		// 初始化一个数组来存储工作表的名字
		var sheetNames = [];
		// 遍历所有工作表并获取名字
		for (var i = 1; i <= sheets.Count; i++) {
			sheetNames.push(sheets.Item(i).Name);
		}
		// 输出结果
		return sheetNames;
	}

/*
		工作表，第col列最后一个非空单元格，的行数
	*/
	function $SheetsLastRowNum(sheetName,col){
		if($SheetsIsHave(sheetName)){
			return Application.ActiveWorkbook.Worksheets.Item(sheetName).Cells.Item(
				Application.ActiveWorkbook.Worksheets.Item(sheetName).Rows.Count, col).End(-4162).Row;
		}
		return 0;

	}

/*
		指定名字工作表
	*/
	function $SheetsByName(sheetName){
		if($SheetsIsHave(sheetName)){
			return Application.ActiveWorkbook.Worksheets.Item(sheetName);
		}
	}


	/*
		尝试删除指定名字工作表，若因共享未删除则对表进行标记待删除

	*/
	function $SheetDelOrSign(sheetName) {
		const activeWorkbook = window.Application.ActiveWorkbook;
		const theSheet = activeWorkbook.Worksheets.Item(sheetName);
		if (!theSheet) {
			alert("当前无活动工作表");
			return;
		}
		if($SheetsIsHave(sheetName)){
			// 提示确认
			const confirmMsg = `确认删除工作簿【${activeWorkbook.Name}】中的工作表【${theSheet.Name}】吗？`;
			if ($MsgBox(confirmMsg)) {
				// 获取当前工作簿对象（假设activeSheet.Parent指向工作簿）
				const workbook = theSheet.Parent;
				const isShared = workbook.MultiUserEditing; // 假设IsShared属性表示是否共享

				if (!isShared) {
					// 未共享时执行删除
					theSheet.Delete();
					alert("删除成功");
				} else {
					// 共享时重命名工作表并提示
					const originalName = theSheet.Name;
					const date = new Date();
					theSheet.Name = "待删除"+$DateFormat(date,"yyyyMMddhhmmss") + originalName;
					alert("共享工作簿模式下不能删除工作表，请确认无人使用时取消共享模式再进行删除。");
					// 注：根据需求，若需要恢复共享模式可在此处添加相关逻辑（如workbook.SetShared(false)等）
				}
			}
		}else{
			alert("工作表不存在");
		}
		
	}
/**
		 * 
		 * 
		 */
	function $SheetsVisible(sheetName,visible){
		var workbook = Application.ActiveWorkbook;
		var sheet = workbook.Worksheets.Item(sheetName);
		sheet.Visible = visible;
	}