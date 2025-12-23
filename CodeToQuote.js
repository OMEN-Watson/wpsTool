Attribute Module_Name = "Module11"


function exportDataToNewExcel() {
  try {


//ActiveSheet.PrintOut()

//alert(quoteFilePath)
//return
//获取 quantity file path 地址
var quantityFilePath=fullPath=ActiveWorkbook.FullName


 //quote 模板 地址：
 QuoteTemplateFilePath="C:\\Users\\GERRY\\Desktop\\Gavin\\001\\garage\\OneDrive_1_2025-7-28\\Quote-Template"
priceValue= findPriceTValue()
siteValue=findSiteValue()
builderValue=findBuilderValue()
  var valueDate =convertDateString(Date());
  var dataArray = [siteValue, builderValue, priceValue, valueDate];
  //判断 价格， builder，site数据 是否为 有效值。-1为无效值， 直接 终止 继续运行
for (var i = 0; i < dataArray.length; i++) {
	if ( dataArray[i]==-1 ) {
           
            return;
        }	
            
             
}
 
        
        
        //获取 算量表 有效 数据行
steelData= collectValidRows()
//获取 报价表 最后一行
lastRowNumber=FileNumberExcel(dataArray)
//获取报价 单号
  quoteNumber=    ActiveSheet.Range("A"+lastRowNumber).Value2
  //获取 报价 单号 中的 数字
quoteIndex=extractQuoteIndex(quoteNumber)
//获取 quote 目标地址
var quoteFilePath=FileNameChange(quantityFilePath,quoteIndex)
if(quoteFilePath==null) return

    //quote 表格 的相关 代码

//  var QuoteFilePath = "C:\\Users\\GERRY\\Desktop\\Gavin\\Huang\\2145Googong\\Quote-2145-GOOGTest"; // Change this path
//  var QuoteFilePath = "D:\\01Gan\\abroad\\study\\05ANU\\Job\\GeneralSteel\\02Measurement\\BLK3SEC6MACNAMARA\\B3S6MACNAMARA\\Quote-2145-GOOGTest"; // Change this path
//if (!QuoteTemplateFilePath) {
//  alert("Quote template path is not defined.");
//  return;
//}
   var wbQuote = Workbooks.Open(QuoteTemplateFilePath);
//插入 报价单号
   ActiveSheet.Range("D13").Value2 =  quoteNumber
//插入 日期
ActiveSheet.Range("E13").Value2="=TODAY()"
//插入 运送地址
ActiveSheet.Range("A18").Value2=siteValue
//插入 价格
SetQuotedPriceValue(priceValue)

//插入 工程量数据
//插入空白行
if(steelData.length==1)
{
	 ActiveSheet.Range("B"+24).Resize(1,3).Value2 = steelData[0];           
	 		 ActiveSheet.Range("B"+23).EntireRow.Delete (xlShiftUp)     
}

else{
	for (var i = 2; i < steelData.length; i++) {
 ActiveSheet.Range("A24").EntireRow.Insert(xlShiftDown, true);
        }
	
	    
    for(var i=0; i<steelData.length; i++){
    	cellNumber=i+23
        ActiveSheet.Range("B"+cellNumber).Resize(1,3).Value2 = steelData[i];            
    }
 
}

wbQuote.SaveAs(quoteFilePath)
//ActiveSheet.PrintOut()

    

  } catch (e) 
  {
   alert ("❌ Error: " + e.message);
  }
}

function openOrActivateWorkbook() {
    var filePath = "C:\\Users\\GERRY\\Desktop\\Gavin\\001\\FileNumber.xlsx"; // Full path including file name
    var fileName = "FileNumber.xlsx"; // Just the file name

 
    var alreadyOpen = false;

    // Loop through open workbooks
    for (var i = 0; i < Workbooks.Count; i++) {
        var wb = Workbooks.Item(i);
        if (wb.Name.toLowerCase() === fileName.toLowerCase()) {
            wb.Activate();
            alreadyOpen = true;
            break;
        }
    }

    // If not already open, open it
    if (!alreadyOpen) {
        var wb = Workbooks.Open(filePath);
        wb.Activate();
    }
    return wb;
}

 function FileNameChange(fullPath,quoteIndex){
	
	// Split into folder path and file name
        var parts = fullPath.split("\\");
        var fileName = parts.pop();
        var folderPath = parts.join("\\");

        // Replace 'Quantity-' (case-insensitive) with 'Quote-'
//        var newFileName = fileName.replace(/^(quantity)/i, "Quote");

        var   newFileName = fileName.replace(/^(quantity)/i, "Quote " + quoteIndex);


        if (newFileName === fileName) {
            alert ("The file name does not start with 'Quantity-' (case-insensitive). No changes made.");
            return null;
        }


        // Construct new full path
        var newFullPath = folderPath + "\\" + newFileName;
//        isExist= FileSystem.Exists(newFullPath)
//if(isExist){
//	var newFilePath=quoteFilePath+" "+priceValue
//	return  newFilePath
//}
        
        return newFullPath
}

  //获取 报价 单号 中的 数字
function extractQuoteIndex(quoteNumber) {
    // "EST 1026" → "1026"
    var match = String(quoteNumber).match(/(\d+)/);
    if (match) return match[1];
    alert("Cannot extract number from quoteNumber: " + quoteNumber);
    return null;
}

 function FileNumberExcel(dataArray){
   
//    var filePath = "C:\\Users\\GERRY\\Desktop\\Gavin\\001\\FileNumber - 副本"; // Change this path
   var filePath = "C:\\Users\\GERRY\\Desktop\\Gavin\\001\\FileNumber"; // Change this path
//   检查 是否已 激活
var wb=openOrActivateWorkbook()

//        var wb = Workbooks.Open(filePath);
lastRowNumber= findLastRowCustom(ActiveSheet)

     // === Insert array values into columns B-E ===
        for (var i = 0; i < 4; i++) {
            ActiveSheet.Cells(lastRowNumber, i + 2).Value2 = dataArray[i]; // B=2, C=3, D=4, E=5
        }

        // === Save and close ===
        wb.Save();
        
        return 	lastRowNumber

}

function collectValidRows() {
    try {

  
        var collectedData = []; // Box for valid rows
        var row = 2; // Start from second row
        var maxRows = 10000; // Safety limit

        while (row <= maxRows) {
            var isEmptyRow = true;
            var isValidRow = true;
            var rowData = [];
 rowData.push(row);
            // Check columns B to G (2 to 7)
            for (var col = 2; col <= 7; col++) {
                var val = ActiveSheet.Cells(row, col).Value2;

                if (val != null && val!== "") {
                    isEmptyRow = false; // This cell has value
                    rowData.push(val);
                } else {
                    rowData.push(""); // Keep structure
                    isValidRow = false; // One empty cell => not valid row
                }
            }

            // Stop if all B-G are empty
            if (isEmptyRow) {
                break;
            }

            // If it's a valid row (all B-G filled)
            if (isValidRow) {
//            	newRowData=[rowData[0],rowData[1],rowData[2],rowData[4]]
newRowData=[ rowData[1],rowData[2],rowData[4]]
                collectedData.push(newRowData);
            }

            row++;
        }
// Show collected data
//        if (collectedData.length > 0) {
//            var message = "Valid rows found:\n";
//            for (var i = 0; i < collectedData.length; i++) {
//                message += "Row :" + collectedData[i].join(", ") + "\n";
//            }
//            alert(message);
//        } else {
//            alert("No valid rows found.");
//        }
//   alert(collectedData.length);
        return collectedData; // Returns the 2D array (box)

    } catch (err) {
        alert("Error: " + err.message);
    }
}


function MyFun(){
    //选中B4单元格
    Range("B4").Select();
    //圆括号里面就是选择B4单元格的文字
    (obj=>{
        //改变这个字体的颜色
        obj.Color = 2;        
    })(Selection.Font);
    //圆括号里面就是选择B4单元格的内部背景
    (obj=>{
        //改变这个背景的颜色
        obj.Color = 65536;
    })(Selection.Interior);
    (obj=>{
        //改变这个背景的颜色
        obj.Color = 65536;
    })();

}
function convertDateString(dateString) {
    // Example input: "Fri Jul 25 2025 09:27:45 GMT+1000 (Australian Eastern Standard Time)"
    var parts = dateString.split(" "); // ["Fri", "Jul", "25", "2025", "09:27:45", "GMT+1000", "(Australian", "Eastern", "Standard", "Time)"]

    var monthMap = {
        Jan: "01",
        Feb: "02",
        Mar: "03",
        Apr: "04",
        May: "05",
        Jun: "06",
        Jul: "07",
        Aug: "08",
        Sep: "09",
        Oct: "10",
        Nov: "11",
        Dec: "12"
    };

    var month = monthMap[parts[1]];
    var day = parts[2];
    var year = parts[3];

    return  day+ "/" + month + "/" + year; // Format: MM/DD/YYYY
}


function findLastRowCustom(sheet) {
    var lastRow = 1; // Start from the first row
    var row = 1;

    while (true) {
        var firstCell = sheet.Cells(row, 1).Value2;  // Column A
        var next4Empty = true;

        // Check if columns B-E are empty
        for (var col = 2; col <= 5; col++) {
//        	if()
            if (sheet.Cells(row, col).Value2 != null && sheet.Cells(row, col).Value2 != "") {
                next4Empty = false;
                break;
            }
        }

        if (firstCell != null && firstCell != "" && next4Empty) {
            lastRow = row;
            break; // Found the last row
        }

        row++;

        // Prevent infinite loop (assume max 10,000 rows)
        if (row > 100000) {
           alert("No last row found by the given logic.");
            break;
        }
    }

    return lastRow;
}


function findPriceTValue() {
    try {
    	lastRow=1000
          var headerRow = 1;
          var valueFound=-1;
        // Step 1: Find the row containing "Price /t" in column E
        for (var row = 1; row <= lastRow; row++) {
            var cellValue = ActiveSheet.Cells(row, 5).Value2; // Column E (5th column)
            
            if (cellValue == "Price /t") {
                headerRow = row;
                break;
            }
        }

        if (headerRow == -1) {
            alert("Header 'Price /t' not found in column E!");
            return   valueFound;;
        }

        // Step 2: Find the first non-empty cell below "Price /t"
        for (var row = headerRow + 1; row <= lastRow; row++) {
            var val = ActiveSheet.Cells(row, 5).Value2;
            if (val != null && val != "") {
                valueFound = val;
                break;
            }
        }


        if (valueFound != -1) {
//            alert("First value under 'Price /t' is: " + valueFound);
        } else {
            alert("No value found under 'Price /t'!");
        }
    return valueFound;
    } catch (err) {
        alert("Error: " + err.message);
    }
}

function findSiteValue() {
    try {
    	lastRow=1000
          var headerRow = -1;
           var valueFound=-1;
        // Step 1: Find the row containing "Price /t" in column E
        for (var row = 1; row <= lastRow; row++) {
            var cellValue = ActiveSheet.Cells(row, 3).Value2; // Column E (5th column)
            
            if (cellValue == "Site:") {
                headerRow = row;
                break;
            }
        }

        if (headerRow == -1) {
            alert("Header 'Site' not found in column C!");
            return valueFound;
        }

        // Step 2: Find the first non-empty cell below "Price /t"
        for (var row = headerRow + 1; row <= lastRow; row++) {
            var val = ActiveSheet.Cells(row, 3).Value2;
            if (val != null && val != "") {
                valueFound = val;
                break;
            }
        }


        if (valueFound != -1) {
//            alert("First value under 'Site' is: " + valueFound);
        } else {
            alert("No value found under 'Site'!");
        }
    return valueFound;
    } catch (err) {
        alert("Error: " + err.message);
    }
}

function findBuilderValue() {
    try {
    	lastRow=1000
          var headerRow = -1;
          valueFound=-1
        // Step 1: Find the row containing "Price /t" in column E
        for (var row = 1; row <= lastRow; row++) {
            var cellValue = ActiveSheet.Cells(row, 3).Value2; // Column E (5th column)
            
            if (cellValue == "Builder:") {
                headerRow = row;
                break;
            }
        }

        if (headerRow == -1) {
            alert("Header 'Builder:' not found in column C!");
            return valueFound;
        }

        // Step 2: Find the first non-empty cell below "Price /t"
        for (var row = headerRow + 1; row <= lastRow; row++) {
            var val = ActiveSheet.Cells(row, 3).Value2;
            if (val != null && val != "") {
                valueFound = val;
                break;
            }
        }


        if (valueFound != -1) {
//            alert("First value under 'Builder:' is: " + valueFound);
        } else {
            alert("No value found under 'Builder:'!");
        }
    return valueFound;
    } catch (err) {
        alert("Error: " + err.message);
    }
} 


function SetQuotedPriceValue(finalPrice) {
    try {
    	lastRow=1000
          var headerRow = -1;
        // Step 1: Find the row containing "Quoted Price " in column E
        for (var row = 1; row <= lastRow; row++) {
            var cellValue = ActiveSheet.Cells(row, 5).Value2; // Column E (5th column)
            
            if ( cellValue == "Quoted Price"||String(cellValue).includes("Quoted Price")) {
                headerRow = row;
                break;
            }
     
           
        }

        if (headerRow == -1) {
            alert("Header 'Quoted Price' not found in column C!");
            return;
        }
        valueRow=headerRow+1
  ActiveSheet.Range("E"+valueRow).Value2=finalPrice + "+GST"
        // Step 2: Find the first non-empty cell below "Price /t"
        for (var row = headerRow + 1; row <= lastRow; row++) {
            var val = ActiveSheet.Cells(row, 5).Value2;
            if (val != null && val != "") {
                valueFound = val;
             
                break;
            }
        }


        if (valueFound != null) {
//            alert("First value under 'Builder:' is: " + valueFound);
        } else {
            alert("No value found under 'Builder:'!");
        }
    return valueFound;
    } catch (err) {
        alert("Error: " + err.message);
    }
}
