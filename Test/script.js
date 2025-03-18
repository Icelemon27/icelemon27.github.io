// 全局变量用于跟踪是否已经添加了事件监听器
let isControlRangeListenerAdded = false;
let isVisibilityHandlerAdded = false;
let objGlobalFormulasAddress = null; //是一个对象，保存最初的变量名和变量地址的对应
let strGlobalFormulasCell = null; // 一个单元格地址，公式里变量名代替变量地址的存放单元格

let strGlobalLabelRange = null; // 保存在pivotTable下面一行不带sum of 的地址

let strGlbBaseLabelRange = null; //保存在工作表Process中Base部分，变量的名字对应的LabelRange
//-----------------------------Process Range 全局变量------------------------
let StrGlobalProcessRange = null; // 保存在Process工作表中的记录combine 过来的新的Range
let StrGlobalPreviousProcessRange = null; // 在ProcessRange往右移动之的前一个ProcessRange地址
let StrGlbProcessSolveStartRange = null; // 在Process的Base中求解变量放的第一行的公式地址。
let StrGblProcessDataRange = null; // 在Process中Base中的dataRange
let NumVarianceReplace= 0; // 记录变量被替换的次数
let NumMaxVariance = null; // 全部的变量个数
let StrGblBaseProcessRng = null; // BaseRange 地址
let StrGblTargetProcessRng = null; //TargetRange 地址å
let NumImpact = 0; // 记录有多少个result 需要计算impact 
let StrGlbIsDivided = false;
let StrGlbDenominator = null; //除法的分母，contribution的时候调用
let ContributionEndCellAddress = null; //Process表中Contribution的结束单元格再往右移动一格的地址，为后面variance 表格做为基础地址使用

let checkFormulaGlobalVar = false; //全局变量，判断是否是在检查formula的错误中，如果是，则不检查表头等差异，直接删除全部其余表，重新run

//------Bridge Data Temp 全局变量--------------
// let StrGblProcessSumCell = null;
let ResultSumType = null; // 判断Result的结果是SumY还是SumN

//--------------FormulasBreakdown 全局变量--------------
let ArrVarPartsForPivotTable = []; //存储变量筛选透视表

let GblComparison = false; //检测是否表头已经被检测过是否一致，避免runProgram调用循环

let checkType2Var = []; //存储SumN + SumN 新增的变量

//-------建立一个数组包含所有变量元素和运算符的对象，
//包含公式里的变量名，变量对应的单元格地址，是否是运算符号，是否是SumY, 是否是replace, others 表示其他情况，其他的自定义变量？？
let FormulaTokens = [];

/////////////---------初始化全局变量----Start---/////////////

function InitializeGlobalVariable(){

  // 全局变量用于跟踪是否已经添加了事件监听器
  isControlRangeListenerAdded = false;
  isVisibilityHandlerAdded = false;
  objGlobalFormulasAddress = null; //是一个对象，保存最初的变量名和变量地址的对应
  strGlobalFormulasCell = null; // 一个单元格地址，公式里变量名代替变量地址的存放单元格

  strGlobalLabelRange = null; // 保存在pivotTable下面一行不带sum of 的地址

  strGlbBaseLabelRange = null; //保存在工作表Process中Base部分，变量的名字对应的LabelRange
  //-----------------------------Process Range 全局变量------------------------
  StrGlobalProcessRange = null; // 保存在Process工作表中的记录combine 过来的新的Range
  StrGlobalPreviousProcessRange = null; // 在ProcessRange往右移动之的前一个ProcessRange地址
  StrGlbProcessSolveStartRange = null; // 在Process的Base中求解变量放的第一行的公式地址。
  StrGblProcessDataRange = null; // 在Process中Base中的dataRange
  NumVarianceReplace= 0; // 记录变量被替换的次数
  NumMaxVariance = null; // 全部的变量个数
  StrGblBaseProcessRng = null; // BaseRange 地址
  StrGblTargetProcessRng = null; //TargetRange 地址
  NumImpact = 0; // 记录有多少个result 需要计算impact 
  StrGlbIsDivided = false;
  StrGlbDenominator = null; //除法的分母，contribution的时候调用
  ContributionEndCellAddress = null; //Process表中Contribution的结束单元格再往右移动一格的地址，为后面variance 表格做为基础地址使用

  //------Bridge Data Temp 全局变量--------------
  // let StrGblProcessSumCell = null;

  //--------------FormulasBreakdown 全局变量--------------
  ArrVarPartsForPivotTable = []; //存储变量筛选透视表

  // GblComparison = false; //>>>>>>>>>这个用来检测runProgram 是否曾经运行检测过，不能再次初始化  检测是否表头已经被检测过是否一致，避免runProgram调用循环

  let checkType2Var = []; //存储SumN + SumN 新增的变量

  //-------建立一个数组包含所有变量元素和运算符的对象，
  //包含公式里的变量名，变量对应的单元格地址，是否是运算符号，是否是SumY, 是否是replace, others 表示其他情况，其他的自定义变量？？
  FormulaTokens = [];

}

/////////////---------初始化全局变量----End---/////////////

(function() {
    if (window.consoleLogModified) return;  // 如果已经修改过 console.log，则不再执行修改
    var originalConsoleLog = console.log;  // 保存原始的 console.log 函数

    console.log = function(message) {
        originalConsoleLog(message);  // 继续在控制台输出日志
        logMessage(message);  // 同时输出到界面上的日志区域
    };

    window.consoleLogModified = true;  // 设置一个标志，表明 console.log 已被修改
})();

//----------------下拉菜单的样式---------------
(async () => {
  // 加载资源后初始化Select2
  await loadResources();
  initializeSelect2();

  // 异步加载所需的外部资源，如 jQuery 和 Select2
  async function loadResources() {
    return new Promise((resolve, reject) => {
      // 动态加载 jQuery
      const jqueryScript = document.createElement("script");
      jqueryScript.src = "https://code.jquery.com/jquery-3.6.0.min.js";
      jqueryScript.onload = () => {
        
        // 动态加载 Select2 CSS 样式
        const select2Css = document.createElement("link");
        select2Css.rel = "stylesheet";
        select2Css.href = "https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css";
        document.head.appendChild(select2Css);

        // 动态加载 Select2 JS 脚本
        const select2Script = document.createElement("script");
        select2Script.src = "https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js";
        select2Script.onload = resolve;
        select2Script.onerror = reject;
        document.head.appendChild(select2Script);
      };
      jqueryScript.onerror = reject;
      document.head.appendChild(jqueryScript);
    });
  }

})();

//----------------下拉菜单的样式---end------------

let isInitializing = null; // 用于标记初始化状态

Office.onReady(async(info) => {

    console.log("Enter onReady");
    // Check that we loaded into Excel
    if (info.host === Office.HostType.Excel) {
      isInitializing = true;
      //初始化检查，是否有Data 工作表

      // await TaskPaneStart("Data"); // 没有工作表则生成新的Data 工作表
      await createSourceData("Data");

      // let CheckBridgeDataSheet = await TaskPaneStart("Bridge Data");
      // //
      // if (CheckBridgeDataSheet === 1){   //如果已经存在了Bridge Data工作表
      //   await ClearBridgeData(); //////清除BridgeData工作表//////
      //   await CopyDataSheet(); //////拷贝DataSheet工作表///////
      // }

      //若数据没有变化，则生成下拉菜单, 若有变化，则提示是否要生成新的waterfall
      let GenerateWaterfall = await handleCompareFieldType();
      console.log("GenerateWaterfall is " + GenerateWaterfall);
      if(GenerateWaterfall){
          await runProgramHandler();

      }
      
        // ----------初始化按钮绑定-----------------------
        // 确认按钮点击事件
        // document.querySelector("#confirmKeyWarningButton").addEventListener("click", () => {
        //   hideKeyWarning();
        //   // CheckKey(); // 再次检查 暂时不用多次检查
        // });
        // ----------初始化按钮绑定---End--------------------

        document.getElementById("runProgram").onclick = runProgramHandler;
        // document.getElementById("refreshWaterfall").onclick = refreshBridge;

        // document.getElementById("refreshWaterfall").onclick = checkBridgeDataHeadersAndValues;
        // document.getElementById("refreshWaterfall").onclick = WaterfallVarianceTable;
        document.getElementById("refreshWaterfall").onclick = () =>CheckFormula("O3");
        
        // 确保Waterfall工作表事件处理程序已添加
        // await ensureWaterfallEventHandler();

        Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            // 确保能够读取单元格范围的地址
            range.load("address");
            await context.sync();
    
            // 显示初始选中范围
            document.getElementById("selectedRange").value = range.address;
        });

        
        //监控Data 数据表的变化
        // Excel.run(async (context) => {
        //     const sheet = context.workbook.worksheets.getItem("Data");
        //     sheet.onChanged.add(onChange);
            
        //     await context.sync();
        //     console.log("Worksheet onChanged event handler has been added.");
        // }).catch(function(error) {
        //     console.error("Error: " + error);
        // });

        //监控Waterfall数据表的变化
      //   Excel.run(async (context) => {
      //     const sheet = context.workbook.worksheets.getItem("Waterfall");
      //     sheet.onChanged.add(monitorRangeChanges);
          
      //     await context.sync();
      //     console.log("Waterfall onChanged event handler has been added.");
      // }).catch(function(error) {
      //     console.error("Error: " + error);
      // });

      document.getElementById("restoreOptions").onclick = async (event) => {
        // 检查并清空工作表 SelectedValue1 和 SelectedValue2
        await clearWorksheetDataIfExists("SelectedValue1");
        await clearWorksheetDataIfExists("SelectedValue2");
    
        // 确保 CreateDropList 是异步函数，调用前使用 await
        await CreateDropList(event);
        // isInitializing = false;
        await refreshBridge();
    };

        setUpEventHandlers();
        isInitializing = false;
    }
});

//刷新waterfall
async function refreshBridge() {
  console.log("refreshBridge 1");
  isInitializing = true; // 设为初始化状态，避免waterfall工作表 中循环更新
  await Excel.run(async (context) => {
    
    const result = await compareFieldType();
    console.log("refreshBridge 2");
    // const result = 0;

    // 这里需要增加更多的检测条件，例如是否全部需要的工作表都存在
    if (result === 0) {
        console.log("No changes detected.");

      // 调用更新数据透视表的函数
      console.log("refreshBridge 3");
      await updatePivotTableFromSelectedOptions("dropdown-container1", "BasePT");

      await updatePivotTableFromSelectedOptions("dropdown-container2", "TargetPT");
      console.log("refreshBridge 4");
      // 调用 DrawBridge 函数
      await BridgeCreate();
      await CreateContributionTable(); //
      await DrawBridge();
      console.log("refreshBridge 5");
    }

  });
  isInitializing = false; // 结束初始化状态，避免waterfall工作表 中循环更新
}

// 函数：检查工作表是否存在，如果存在则清空内容
async function clearWorksheetDataIfExists(sheetName) {
  try {
      await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          const sheet = sheets.getItemOrNullObject(sheetName);
          await context.sync(); // 同步以加载 isNullObject

          if (!sheet.isNullObject) {
              // 如果工作表存在，清空其数据
              const range = sheet.getUsedRange(); // 获取已用范围
              range.clear(); // 清空内容
              console.log(`Cleared data in worksheet: ${sheetName}`);
          } else {
              console.log(`Worksheet ${sheetName} does not exist.`);
          }
      });
  } catch (error) {
      console.error(`Error clearing worksheet ${sheetName}:`, error);
  }
}

// 显示进度条
async function showProgressBar() {
  const progressContainer = document.getElementById("progressContainer");
  progressContainer.style.display = "block";
  updateProgressBar(0);
}

// 更新进度条
async function updateProgressBar(percentage) {
  const progressBar = document.getElementById("progressBar");
  progressBar.style.width = `${percentage}%`;
  progressBar.textContent = `${percentage}%`;
}

// 隐藏进度条
async function hideProgressBar() {
  const progressContainer = document.getElementById("progressContainer");
  progressContainer.style.display = "none";
}

//创建BridgeData的工作表
async function CreateBridgeData() {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    // 检查是否存在同名的工作表
    let BridgeSheet = workbook.worksheets.getItemOrNullObject("Bridge Data");
    await context.sync();

    if (BridgeSheet.isNullObject) {
      // 工作表不存在，创建新工作表
      BridgeSheet = context.workbook.worksheets.add("Bridge Data");
      await context.sync();
      console.log("创建了新工作表：" + "Bridge Data");
      await CopyDataSheet(); //////拷贝DataSheet工作表///////
    }else{
      //如果已经存在Bridge Data，则从Data中拷贝数据
      await ClearBridgeData(); //////清除BridgeData工作表//////
      await CopyDataSheet(); //////拷贝DataSheet工作表///////

    }

      await DoNotChangeCellWarning("Bridge Data");
  });
}


async function runProgramHandler() {

      console.log("Enter runProgramHandler");

      InitializeGlobalVariable(); // 初始化全局变量

      //判断Bridge Data工作表是否存在，如不存在则建立，
      // let CheckBridgeDataSheet = await TaskPaneStart("Bridge Data");
      // //
      // if (CheckBridgeDataSheet === 1){   //如果已经存在了Bridge Data工作表，则重新拷贝Data工作表的内容
      //   await ClearBridgeData(); //////清除BridgeData工作表//////
      //   await CopyDataSheet(); //////拷贝DataSheet工作表///////
      // }

      // await ClearBridgeData(); ///////清除BridgeData工作表///////
      // await CopyDataSheet(); //////拷贝DataSheet工作表//////
      // 隐藏 .prompt-container 容器
      const promptContainer = document.querySelector(".prompt-container");
      if (promptContainer) {
          promptContainer.style.display = "none"; // 隐藏提示容器
      }

      console.log("Initializing runProgram...");

      isInitializing = true; // 设置初始化标记

      //检查比较数据表头和维度类型,GblComparison检测是否已经对比过，避免循环调用
      console.log("GblComparison is ");
      console.log(GblComparison);
      // console.log("checkFormulaGlobalVar is ");
      // console.log(checkFormulaGlobalVar);

      //直接运行的时候，不需要比较表头的变化信息？？？/////////////////////////
      // if((!GblComparison)){
      //   let shouldRestart = await handleCompareFieldType();
      //   console.log("shouldRestart is " + shouldRestart);
      //   // 如果用户确认重新生成 Waterfall，则继续处理后续逻辑，否则返回结束
      //   if (!shouldRestart) {
      //     return;
      //   }
      //   console.log("RunProgram - CheckDimension");
      // }
      ////////////////////////////////////////////////////////////////////

      // 检查是否存在指定的工作表
      let sheetsExist = await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        let sheetNames = sheets.items.map(sheet => sheet.name);

        let requiredSheets = ["FormulasBreakdown", "Process", "Waterfall"];

        let existingSheets = requiredSheets.filter(name => sheetNames.includes(name));

        return existingSheets.length > 0;
      });

      if (sheetsExist) {
        // 如果工作表存在，显示提示框

        let userConfirmed = await showWaterfallPrompt();
        if (!userConfirmed) {
            // 用户选择不重新生成，退出函数
            isInitializing = false;
            await hideProgressBar(); // 隐藏进度条
            await hideWarning(); //隐藏警告不要修改excel
            return;
        }
      }
      //下面可以放置各种检查条件

      // //检查是否有Key
      // let hasKey = await CheckKey();
      // if(!hasKey){
      //   return;
      // }
      console.log("RunProgram 0");
    // }


      //Data 第一行是否有含有必须的全部标题
      const hasRequiredHeaders = await Excel.run(async (context) => {
        return await checkRequiredHeaders(context);
        });

      if (hasRequiredHeaders) {
          return;
      }

      //Data 第一行是否有重复的Key值
      const hasDuplicateKey = await Excel.run(async (context) => {
        return await hasDuplicateKeyInFirstRow(context);
        });

      if (hasDuplicateKey) {
          return;
      }


      console.log("RunProgram 1")
      //Data 第一行是否有重复的Result值
      const hasDuplicateResult = await Excel.run(async (context) => {
        return await hasDuplicateResultInFirstRow(context);
        });

      if (hasDuplicateResult) {
          return;
      }


      //----检查第三行开始的数据类型是否是正确的----
      const hasCorrectDataType = await Excel.run(async (context) => {
        return await checkBridgeDataHeadersAndValues(context);
        });

      if (hasCorrectDataType) {
          return;
      }

      console.log("RunProgram 2");
      console.log("hasDuplicateKey");
      console.log(hasDuplicateKey);
      console.log("hasDuplicateResult");
      console.log(hasDuplicateResult);
      // 如果上面的前提条件成立，则不继续执行后面的代码
      // if (hasDuplicateKey || hasDuplicateResult) {
      //     return;
      // }
      console.log("RunProgram 3");
      
      
      //判断是否有加减乘除以外不合适的运算符
      let ValidOperator = await CheckValidOperator(); 
      console.log("isInValidOperator is " + ValidOperator);
      if(!ValidOperator){
        return;
      }

      //判断Data第二行是否有重复的字段名，如果重复数据透视表不能生成
      let DuplicateTitle = await CheckDuplicateHeaders();
      console.log("DuplicateTitle is " + DuplicateTitle);
      if(!DuplicateTitle){
        return;
      }

      await showWarning(); //警告不要修改excel
      await showProgressBar(); // 显示进度条

      let progress = 0;
      const totalSteps = 25; // 根据脚本的执行步骤总数设置这个值

      function incrementProgress() {
          progress += 100 / totalSteps;
          updateProgressBar(Math.min(progress, 100)); // 确保进度不会超过100%
      }


      //在不同阶段添加进度更新
      incrementProgress();
      const startTime = new Date(); // Start timer
      await Excel.run(async (context) => {
          // let HideSheetNames = ["Base", "连除"];
          // await hideSheets(context, HideSheetNames); // 隐藏工作表以防止用户操作
          // console.log("RunProgram 4")
          // await protectSheets(context, HideSheetNames); // 保护工作表以防止用户交互
          // console.log("RunProgram 5")
          // await disableScreenUpdating(context); // 添加 await 以正确等待挂起
          // console.log("RunProgram 6")
          
          incrementProgress();
          //--------------程序开始--------------------
          await CreateBridgeData(); // 创建新的Bridge Data表 ***********这样的话每一次按钮都要拷贝一次数据，那么下面的deleteProcessSum()也不需要了
          // await deleteProcessSum();
          const sheetsToDelete = ["FormulasBreakdown", "Bridge Data Temp", "TempVar","BasePT","TargetPT","Combine","Analysis","Process","Waterfall","SelectedValue1","SelectedValue2"];
          await deleteSheetsIfExist(sheetsToDelete);
          
          incrementProgress();

          await CreateTempVar();
          let result = await FormulaBreakDown2();
          if (result === "Error") {
            console.log("FormulaBreakDown 返回 Error，终止 main 函数执行");
            return; // main 函数提前退出
          }
          incrementProgress();
          await createPivotTableFromBridgeData("BasePT");
          incrementProgress();
          await createPivotTableFromBridgeData("TargetPT");
          incrementProgress();
          await createPivotTableFromBridgeData("Combine");
          incrementProgress();
          NumVarianceReplace = 0; //中间有中断的可能，每次都需要清零重新计数，初始化以便按钮任意点击
          
          await copyAndModifySheet("FormulasBreakdown","Bridge Data Temp"); //********* */ 从Breakdown 中复制，用最新的公式，后面看是否需要删掉，直接用Breakdown
          incrementProgress();
          await CreateAnalysisSheet("BasePT","Analysis");
          incrementProgress();
          await CreateAnalysisSheet("Combine","Process");
          incrementProgress();
          await fillProcessRange("TargetPT");
          //await runProcess();
          //await GetFormulasAddress("Bridge Data Temp", strGlobalFormulasCell ,"Process", strGlbBaseLabelRange);
          //await CopyFormulas();
          incrementProgress();
          await ResolveLoop();
          incrementProgress();
          //如果result最后是除法则需要用公式，不用SumIf公式
          await ResultDivided();
          incrementProgress();
          await copyProcessRange(); // ProcessRange 平移
          incrementProgress();
          await fillProcessRange("BasePT");
          incrementProgress();
          //await GetFormulasAddress("Bridge Data Temp", strGlobalFormulasCell ,"Process", strGlbBaseLabelRange);
          //await CopyFormulas();
          await ResolveLoop();
          incrementProgress();
          //如果result最后是除法则需要用公式，不用SumIf公式
          await ResultDivided();
          incrementProgress();
          let VarFormulasObjArr = await GetBridgeDataFieldFormulas();
          await VarStepLoop(VarFormulasObjArr);

          incrementProgress();
          await BridgeCreate();
          incrementProgress();
          await Contribution();
          incrementProgress();
          //创建用户使用的Contribution Table
          await CreateVarianceTable();
          await CreateContributionTable();
          incrementProgress();

          // await WaterfallVarianceTable();
          await DrawBridge();
          await setFormat("Waterfall");
          incrementProgress();

          console.log("StrGlobalProcessRange is " + StrGlobalProcessRange);
          console.log("ContributionEndCellAddress is " + ContributionEndCellAddress);
          console.log("strGlbBaseLabelRange is " + strGlbBaseLabelRange);
          console.log("StrGblBaseProcessRng is " + StrGblBaseProcessRng);
          console.log("StrGblProcessDataRange is " + StrGblProcessDataRange);

          //创建下拉菜单
          await CreateDropList();
          incrementProgress();

          let HideSheetNames = ["FormulasBreakdown","Bridge Data", "Bridge Data Temp", "TempVar", "BasePT", "TargetPT", "Combine", "Analysis", "SelectedValue1", "SelectedValue2"];
          await hideSheets(context, HideSheetNames); // 隐藏工作表以防止用户操作

        // await enableScreenUpdating(context); // 添加 await 以正确等待恢复
        // await unprotectSheets(context, HideSheetNames); // 操作完成后取消保护工作表
        // incrementProgress();
        // await unhideSheets(context, HideSheetNames); // 操作完成后取消隐藏工作表
        console.log("RunProgram 6");

        await createFieldTypeMapping();
        incrementProgress();
        // let sheet = context.workbook.worksheets.getItem("Waterfall");
        // console.log("RunProgram 7")
        // sheet.onChanged.add(monitorRangeChanges); // 加入监控
        // console.log("RunProgram 8")
      });
      console.log("RunProgram End")
      ////--------------程序结束--------------------
      await hideProgressBar(); // 隐藏进度条
      await hideWarning(); //隐藏警告不要修改excel
      isInitializing = false; // 解除初始化标记    
      GblComparison = false;      


      const endTime = new Date(); // End timer
      const elapsedTimeInSeconds = Math.floor((endTime - startTime) / 1000);  // Calculate elapsed time in seconds
  
      // Convert seconds to MM:SS format
      const minutes = Math.floor(elapsedTimeInSeconds / 60);
      const seconds = elapsedTimeInSeconds % 60;
      const formattedTime = `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;

      // Write the formatted time to the Waterfall worksheet in cell L1
      await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem("Waterfall");
          const range = sheet.getRange("L1");
          range.values = [[`Execution Time: ${formattedTime}`]];
          await context.sync();
      });

      console.log(`Execution Time: ${formattedTime}`);
};


function showWaterfallPrompt() {
  return new Promise((resolve, reject) => {
      // 显示提示框
      document.getElementById("waterfallPrompt").style.display = "block";
      // 显示遮罩层
      document.getElementById("modalOverlay").style.display = "block";

      // 禁用其他交互
      document.querySelector('.container').classList.add('disabled');

      // 获取按钮元素
      let confirmButton = document.getElementById("confirmWaterfall");
      let cancelButton = document.getElementById("cancelWaterfall");

      // 移除之前的事件监听器
      confirmButton.onclick = null;
      cancelButton.onclick = null;

      // 设置事件监听器
      confirmButton.onclick = async function () {
          // 用户点击了“是”
          try {
              await Excel.run(async (context) => {
                  const workbook = context.workbook;

                  // 定义要删除的工作表名称数组 *********这里没有删除Bridge Data 工作表
                  const sheetsToDelete = ["FormulasBreakdown", "Bridge Data","Bridge Data Temp","TempVar","BasePT","TargetPT","Combine","Analysis","Process","Waterfall"]; // 可根据需要添加更多工作表

                  const sheets = workbook.worksheets;
                  sheets.load("items/name");
                  await context.sync();

                  // 遍历要删除的工作表名称数组
                  sheetsToDelete.forEach(sheetName => {
                      if (sheets.items.some(sheet => sheet.name === sheetName)) {
                          // 如果工作表存在，则删除
                          const sheet = sheets.getItem(sheetName);
                          sheet.delete();
                          console.log(`Worksheet ${sheetName} has been deleted.`);
                      } else {
                          console.log(`Worksheet ${sheetName} does not exist.`);
                      }
                  });

                  // deleteProcessSum();

                  await context.sync();
              });
          } catch (error) {
              console.error("Error deleting worksheets:", error);
          }

          // 隐藏提示框和遮罩层
          document.getElementById("waterfallPrompt").style.display = "none";
          document.getElementById("modalOverlay").style.display = "none";
          // 重新启用交互
          document.querySelector('.container').classList.remove('disabled');
          // 解析Promise
          resolve(true);
      };

      cancelButton.onclick = function () {
          // 用户点击了“否”
          // 隐藏提示框和遮罩层
          document.getElementById("waterfallPrompt").style.display = "none";
          document.getElementById("modalOverlay").style.display = "none";
          // 重新启用交互
          document.querySelector('.container').classList.remove('disabled');
          // 解析Promise
          resolve(false);
      };
  });
}





//------------------------------------Waterfall 监听事件-----------------------------
// 创建或更新 Waterfall 工作表并添加事件处理程序 ----------目前没有地方引用
// async function createOrUpdateWaterfallSheet() {
//   await Excel.run(async (context) => {
//       const sheets = context.workbook.worksheets;

//       // 检查工作表是否已存在
//       let sheet = sheets.getItemOrNullObject("Waterfall");
//       await context.sync();

//       if (sheet.isNullObject) {
//           // 创建新的 "Waterfall" 工作表
//           sheet = sheets.add("Waterfall");
//           console.log("Waterfall sheet created.");
//       } else {
//           console.log("Waterfall sheet already exists.");
//       }

//       // 添加事件处理程序
//       await addWaterfallEventHandler(sheet,context);

//       await context.sync();
//   }).catch(function(error) {
//       console.error("Error in createOrUpdateWaterfallSheet:", error);
//   });
// }

// 确保 Waterfall 工作表的事件处理程序已添加
// async function ensureWaterfallEventHandler() {
//   // 添加工作表添加事件的监听器
//   Excel.run(async (context) => {
//       context.workbook.worksheets.onAdded.add(onSheetAdded);
//       context.workbook.worksheets.onDeleted.add(onSheetDeleted);
//       await context.sync();
//       console.log("Worksheet onAdded and onDeleted event handlers have been added.");

//       // 初始检查是否存在 Waterfall 工作表
//       const sheet = context.workbook.worksheets.getItemOrNullObject("Waterfall");
//       await context.sync();

//       if (sheet.isNullObject) {
//         console.log("No existing Waterfall sheet found. Awaiting addition...");
//       } else {
//           console.log("Waterfall sheet exists. Adding event handler.");
//           await addWaterfallEventHandler(sheet,context);
//       }
//   }).catch(function(error) {
//       console.error("Error ensuring worksheet event handlers:", error);
//   });
// }

// 当工作表被添加时的事件处理程序
async function onSheetAdded(event) {
  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(event.worksheetId);
      sheet.load("name");
      await context.sync();

      console.log("event.worksheetId is " + event.worksheetId);
      console.log("Waterfall sheet is " + sheet.name);
      if (sheet.name === "Waterfall") {
          console.log("OnSheetAdded Here");
          await addWaterfallEventHandler(sheet,context);
          console.log("Event handler added to new Waterfall sheet.");
      }
  }).catch(function(error) {
      console.error("Error in onSheetAdded:", error);
  });
}

// 当工作表被删除时的事件处理程序
async function onSheetDeleted(event) {
  // 您可以在这里处理工作表删除的逻辑
  console.log(`Worksheet with ID ${event.worksheetId} has been deleted.`);

  // 如果需要，清理与该工作表相关的资源
  // 由于事件处理程序会自动解除绑定，无需额外处理
}

// 添加 Waterfall 工作表的事件处理程序
async function addWaterfallEventHandler(sheet,context) {
  console.log("Enter addWaterfallEventHandler");
  
  sheet.load("name");
  console.log("addWaterfallEventHandler step 1");
  await context.sync();
  console.log("addWaterfallEventHandler step 2");
  console.log("addWaterfallEventHandler sheet is " + sheet.name);
  await sheet.onChanged.add(monitorRangeChanges);
  console.log("Attempting to add event handler.");
  await context.sync();
  console.log("Event handler added to Waterfall sheet.");
}

// 监控 Waterfall 工作表中指定范围的更改，并在发生变化时重新绘制图表
async function monitorRangeChanges(event) {

  //在 monitorRangeChanges 中检查 isInitializing 标志，如果为 true，直接返回，避免处理事件。
  if (isInitializing) {
    console.log("Skipping event handling during initialization.");
    return;
  }
  try {
      await Excel.run(async (context) => {
          // 获取 Waterfall 工作表
          console.log("Enter monitorRangeChanges");
          const sheet = context.workbook.worksheets.getItemOrNullObject("Waterfall");
          await context.sync();

          if (sheet.isNullObject) {
              console.log("Waterfall sheet no longer exists. Event handling skipped.");
              return;
          }

          // 获取被改变的 Range 地址
          let changedRange = event.address; // e.g., "Sheet1!$A$1:$B$2"
          console.log("Changed range: " + changedRange);

          // 您的全局变量，指定监控的目标范围

          let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
          let BridgeRangeVar = TempVarSheet.getRange("B6");
          BridgeRangeVar.load("values");
          await context.sync();
      
          let BridgeRangeAddress = BridgeRangeVar.values[0][0];
          let targetRange = BridgeRangeAddress; // e.g., "Waterfall!$A$1:$B$10"
          if (!targetRange) {
              console.error("BridgeRangeAddress is not defined.");
              return;
          }

          console.log("changedRange is " + changedRange);
          console.log("targetRange is " + targetRange);

          if (isRangeIntersecting(changedRange, targetRange)) {
              console.log("Target range changed, updating chart...");
              await DrawBridge_onlyChart(); // 调用更新函数
          } else {
              console.log("Changed range does not affect target range.");
          }

      });
  } catch (error) {
      console.error("Error in monitorRangeChanges:", error);
  }
}

// 检查两个范围是否有交集
function isRangeIntersecting(changedRange, targetRange) {
  // 将范围解析为工作表和地址部分
  const [changedSheet, changedAddress] = splitRange(changedRange);
  const [targetSheet, targetAddress] = splitRange(targetRange);

  // 检查是否在同一工作表
  // if (changedSheet !== targetSheet) {
  //     return false;
  // }

  // 解析范围地址为行列索引
  const changedBounds = parseRangeBounds(changedAddress);
  const targetBounds = parseRangeBounds(targetAddress);

  console.log("")


  if (!changedBounds || !targetBounds) {
      return false;
  }

  // 检查是否有交集
  return rangesIntersect(changedBounds, targetBounds);
}

// 辅助函数：拆解范围为工作表和地址部分
function splitRange(range) {
  const parts = range.split("!");
  return parts.length === 2 ? parts : [null, parts[0]]; // 处理无工作表前缀的情况
}

// 辅助函数：解析范围地址为行列索引
function parseRangeBounds(address) {
//   const regex = /(\$?)([A-Z]+)(\$?)(\d+)(:)?(\$?)([A-Z]*)(\$?)(\d*)/;
    const regex = /(\$?)([A-Za-z]+)(\$?)(\d+)(:)?(\$?)([A-Za-z]*)(\$?)(\d*)/;
  const match = address.match(regex);
  if (!match) return null;

  const [, , startCol, , startRow, colon, , endCol, , endRow] = match;

  return {
      startRow: parseInt(startRow),
      endRow: endRow ? parseInt(endRow) : parseInt(startRow),
      startCol: colToIndex(startCol),
      endCol: endCol ? colToIndex(endCol) : colToIndex(startCol),
  };
}

// 辅助函数：将列字母转换为数字索引
function colToIndex(col) {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - "A".charCodeAt(0) + 1);
  }
  return index;
}

// 辅助函数：判断两个范围是否有交集
function rangesIntersect(bounds1, bounds2) {
  return (
      bounds1.startRow <= bounds2.endRow &&
      bounds1.endRow >= bounds2.startRow &&
      bounds1.startCol <= bounds2.endCol &&
      bounds1.endCol >= bounds2.startCol
  );
}


//------------------------------------Waterfall 监听事件 End-----------------------------


function runProgram() {
    const option = document.getElementById('options').value;
    const isEnabled = document.getElementById('check1').checked;
    // alert(`Running with option ${option} and feature enabled: ${isEnabled}`);

    Excel.run(context => {

        // Insert text 'Hello world!' into cell A1.
        context.workbook.worksheets.getActiveWorksheet().getRange("A5").values = [[`Running with option ${option} and feature enabled: ${isEnabled}`]];
        context.workbook.worksheets.getActiveWorksheet().getRange("A2").values =[['Hello world 0519!']];
        // sync the context to run the previous API call, and return.
        return context.sync();
    });

}


async function createPivotTable() {

    // try {
    //     const headers = await getHeaders();
    //     logMessage(headers);
    // } catch (error) {
    //     console.error('Error logging headers:', error);
    // }
    return Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        worksheets.load("items/name");
        //console.log("Step1")
        await context.sync();

        // 检查是否已存在同名的工作表
        const existingSheet = worksheets.items.find(ws => ws.name === "Pivot Table Sheet");
        //console.log("Step2")
        if (existingSheet) {
            document.getElementById('prompt').style.display = 'block';
            //console.log("Step3")
            return;
        }
        //console.log("Step4")
        await createAndFillPivotTable(context); // 如果没有同名工作表直接创建
    }).catch(handleError);
}

async function createAndFillPivotTable(context) {
    const workbook = context.workbook;
    const selectedRange = workbook.getSelectedRange();
    selectedRange.load("address");

    const newSheet = workbook.worksheets.add("Pivot Table Sheet");
    newSheet.activate();

    await context.sync();
    //console.log(selectedRange.address)
    const pivotTable = newSheet.pivotTables.add( "PivotTable", selectedRange, "A1");
    //console.log("Step5")
    // pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Column1"));
    // console.log("Step6")
    // pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Column2"));
    // console.log("Step7")
    // pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Column3"));
    // console.log("Step8")

    await context.sync();
    //console.log("Step9")
    console.log("PivotTable created on new sheet.");
}

function deleteExistingSheet() {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Pivot Table Sheet");
        sheet.delete();
        await context.sync();
        document.getElementById('prompt').style.display = 'none';
        // 需要创建新的Excel.run 会话来确保上下文正确
        Excel.run(async (newContext) => {
            await createAndFillPivotTable(newContext);
        }).catch(handleError);
    }).catch(handleError);
}

function hidePrompt() {
    document.getElementById('prompt').style.display = 'none';
}

function handleError(error) {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}


// ------------------------------------------------------------------End Pivot Table ---------------------------------------------------------


function sayHello() {
    Excel.run(context => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A1");
        range.values = [['Hello world 0512!']];
        console.log(" THis is log"   )
        logMessage("test")
        return context.sync();
    });
}
// ------------------------------------文本框显示地址--------------------------------------------
function setUpEventHandlers() {
    Excel.run(async (context) => {
        const workbook = context.workbook;
        // 添加工作表激活事件监听器
        workbook.worksheets.onActivated.add(handleWorksheetActivated);
        // 初始设置，确保加载时也能监听当前活动工作表的选区变化
        addSelectionChangedListenerToActiveWorksheet(context);
        await context.sync();
    }).catch(function (error) {
        console.error("Error setting up event handlers: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

function addSelectionChangedListenerToActiveWorksheet(context) {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.onSelectionChanged.add(handleSelectionChange);
    return context.sync();
}

async function handleWorksheetActivated(eventArgs) {
    Excel.run(async (context) => {
        // 移除先前工作表的事件监听器
        context.workbook.worksheets.getItem(eventArgs.worksheetId).onSelectionChanged.remove(handleSelectionChange);
        // 为新激活的工作表添加选区变更事件监听器
        addSelectionChangedListenerToActiveWorksheet(context);
    }).catch(function (error) {
        console.error("Error in handleWorksheetActivated: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}

async function handleSelectionChange(eventArgs) {
    await Excel.run(async (context) => {
        // 获取新选区
        const newRange = context.workbook.getSelectedRange();
        // 加载新选区的地址
        newRange.load("address");
        await context.sync();
        // 更新HTML中的文本框显示新选区的地址
        document.getElementById("selectedRange").value = newRange.address;
    }).catch(function (error) {
        console.error("Error in handleSelectionChange: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}


function logMessage(message) {
    const logOutput = document.getElementById("logOutput");
    const timeNow = new Date().toTimeString().split(" ")[0]; // 获取当前时间的时分秒

    // 检查消息类型，如果是对象或数组，则尝试转换为字符串
    if (typeof message === 'object') {
        try {
            message = JSON.stringify(message, null, 2); // 美化输出
        } catch (error) {
            message = "Error in stringifying object: " + error.message; // 转换失败的处理
        }
    }

    let formattedMessage = message;
    if (Array.isArray(message)) {
        formattedMessage = message.join(", ");
    }

    const newLogEntry = `<div>[${timeNow}] ${formattedMessage}</div>`;

    // 添加新日志到输出区域
    logOutput.innerHTML += newLogEntry;

    // 保持日志条目数量不超过10个
    let logEntries = logOutput.querySelectorAll('div');
    if (logEntries.length > 5000) {
        logEntries[0].remove(); // 移除最旧的日志条目
    }

    // 滚动到最新的日志条目
    logOutput.scrollTop = logOutput.scrollHeight;
}



function isString(value) {
    return typeof value === 'string';
}


// ----------------------------------------------获取表头 -----------------------------------------------------------
async function getHeaders(RowNo) {
    return Excel.run(async (context) => {
        // 获取当前选中的范围

        //const selectedRange = workbook.getSelectedRange(); // 获取当前选中的范围
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        
        // 获取 "Data" 工作表
        const sheet = worksheets.getItem("Data");
        const rowRangeAddress = `${RowNo}:${RowNo}`;
        //const RowRange = sheet.getRange(rowRangeAddress).getUsedRange();
        
        // 获取第一行的范围
        //const rangeAddress = RowRange.load("address");
        const selectedRange = sheet.getRange(rowRangeAddress).getUsedRange();
        //const selectedRange = context.workbook.worksheet("Data").range("RowNo:RowNo");
        // 加载选中范围的行信息
        selectedRange.load('address');
        selectedRange.load('rowCount');
        selectedRange.load('columnCount');

        await context.sync();

        // 获取选中范围第一行的数据范围
        let firstRowAddress = selectedRange.address.split("!")[1].replace(/(\d+):(\d+)/, (match, p1, p2) => `1:${p2}`);
        //let firstRowAddress = selectedRange.offset(RowNo, 0, 1, selectedRange.columnCount).address.split("!")[1].replace(/(\d+):(\d+)/, (match, p1, p2) => `1:${p2}`);
        logMessage(firstRowAddress)
        const headerRange = selectedRange.worksheet.getRange(firstRowAddress);
        headerRange.load("values");  // 请求加载选中范围第一行的值

        await context.sync();

        // 检查选中范围第一行是否加载了值
        if (headerRange.values && headerRange.values.length > 0) {
            let headers = headerRange.values[0].filter(value => value !== "");
            return headers.length > 0 ? headers : ["No headers found or empty first row."];
        } else {
            return ["No headers found or empty first row."]; // 没有找到数据时的返回信息
        }
    }).catch(error => {
        console.error("Error: " + error);
        return ["Error fetching headers: " + error.toString()]; // 返回错误信息
    });
}


//-------------------------------建立datasource 表格-----------------------------------------
async function createSourceData(SheetName) {
  return Excel.run(async (context) => {
        console.log("createSourceData 开始")
        const workbook = context.workbook;
        // const sheetName = "Data";
        //const sheetName = "Data"; 
        const sheets = workbook.worksheets;
        sheets.load("items/name");  // 加载所有工作表的名称

        await context.sync();

        // 检查是否存在同名工作表
        if (sheets.items.some(sheet => sheet.name === SheetName)) {
            // 显示对话框
            // document.getElementById('promptSource').style.display = 'block';
            // // 暂停执行，等待用户响应
            // 这里先注销掉，因为Data工作表有的话就继续执行，没有则创建，不需要询问用户
            return 0;
        } else {
            // 直接创建工作表和设置
            await setupWorksheet(SheetName);
            return 1;
        }
        console.log("createSourceData 完成")
    }).catch(error => {
        console.error("Error: " + error);
    });
}

// ---------------------------------创建数据第一行的各种字段类型选项----------------------------------------
async function setupWorksheet(sheetName) {
    return Excel.run(async (context) => {
        console.log("setupWorksheet 开始");
        const sheet = context.workbook.worksheets.add(sheetName);
        sheet.activate();

        sheet.getRange("A1").values = [["Data Type"]];
        sheet.getRange("A2").values = [["Header"]];
        sheet.getRange("A3").values = [["Data"]];
        
        const options = ["Dimension","SumY","SumN","Result","Key"];
        const validationRule = {
            list: {
                inCellDropDown: true,
                source: options.join(",")
            }
        };
        const dataRange = sheet.getRange("B1:AAA1");
        dataRange.dataValidation.rule = validationRule;

        // 自动调整 A 列宽度以适应内容
        const columnARange = sheet.getRange("A:A");
        columnARange.format.autofitColumns();

        await context.sync();
        console.log("Worksheet and validation setup complete.");

        // 显示提示信息
        await showTaskPaneMessage("请在第一行选择相应的数据类型\n第二行输入数据的标题\n第三行往下输入原始数据。");
        console.log("setupWorksheet 完成");
    });
}







async function showTaskPaneMessage(message) {
  console.log("showTaskPaneMessage 开始");
  const promptContainer = document.getElementById("taskPanePrompt");
  const messageContent = document.getElementById("messageContent");
  const confirmButton = document.getElementById("confirmButton");

  // 替换换行符 \n 或自定义标记 [break] 为 <br> 标签
  const formattedMessage = message.replace(/\n/g, '<br>').replace(/\[break\]/g, '<br>');

    // 设置提示内容
    messageContent.innerHTML = formattedMessage;
    promptContainer.style.display = "block";

    return new Promise((resolve) => {
        confirmButton.onclick = () => {
            promptContainer.style.display = "none"; // 隐藏提示容器
            console.log("showTaskPaneMessage 完成");
            resolve(); // 继续执行后续代码
        };
    });
}



function deleteExistingSheetSource() {
    Excel.run(async (context) => {
        context.workbook.worksheets.getItem("Data").delete();
        await context.sync();
        // 隐藏对话框
        document.getElementById('promptSource').style.display = 'none';
        // 创建新工作表
        setupWorksheet("Data");
    }).catch(error => {
        console.error("Error: " + error);
    });
}

function hidePromptSource() {
    document.getElementById('promptSource').style.display = 'none';
    // 可以在这里添加退出 Office Add-in 的逻辑，如果适用
    // 例如，通过 Office Add-ins API 关闭任务窗格
    console.log("Operation cancelled by the user.");
    // 如果在 Excel Online 中使用，可以考虑使用某种方法来关闭窗格或通知用户操作已取消
    // 如果在桌面应用中，可能需要通过其他方式通知用户
}

//-------------------------------End  建立datasource 表格-----------------------------------------


//-----------------------------------------从 RawData 建立数据透视表------------------------------------------------








//-----------------------------------------获取每个字段的唯一值 ----- 单纯获得不重复的值
async function GetUniqFieldValue() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Data");
      // 获取工作表的已用范围
      const usedRange = sheet.getUsedRange();
      usedRange.load('rowCount, columnCount');
      await context.sync();
  
      // 读取字段名，假设字段名在第二行
      const headerRange = sheet.getRangeByIndexes(1, 1, 1, usedRange.columnCount - 1);
      headerRange.load('values');
      await context.sync();
  
      let headers = headerRange.values[0];
      let uniqueValues = {};
  
      // 初始化每个字段的Set
      headers.forEach(header => {
        uniqueValues[header] = new Set();
      });
  
      // 读取数据，从第三行开始直到最后
      const dataRange = sheet.getRangeByIndexes(2, 1, usedRange.rowCount - 2, usedRange.columnCount - 1);
      dataRange.load('values');
      await context.sync();
  
      // 遍历每一列
      for (let colIndex = 0; colIndex < headers.length; colIndex++) {
        // 使用map提取每一列的值，并应用Set去重
        let columnData = dataRange.values.map(row => row[colIndex]);
        uniqueValues[headers[colIndex]] = new Set(columnData);
      }
  
      // 将每个字段的Set转换为数组
      let results = {};
      for (let header of headers) {
        results[header] = Array.from(uniqueValues[header]);
      }
  
      console.log(results);
      return results;
    }).catch(error => {
      console.error("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}


//判断目前active 的是否是Waterfall 工作表，如果不是则设置
async function activateWaterfallSheet() {
  await Excel.run(async (context) => {
      // 获取当前活动的工作表
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();

      // 加载工作表的名称
      activeSheet.load("name");
      await context.sync();

      // 判断当前活动工作表是否为“Waterfall”
      if (activeSheet.name !== "Waterfall") {
          // 如果不是，则激活名为“Waterfall”的工作表
          const waterfallSheet = context.workbook.worksheets.getItem("Waterfall");
          waterfallSheet.activate();
      }
  });
}

//将列索引（从 0 开始）转换为 Excel 列字母 (A, B, ..., Z, AA, AB, ...)
function toColumnLetter(colIndex) {
  let letter = "";
  let index = colIndex + 1; // Excel 列号是从1开始的，比如 A=1, B=2...
  
  while (index > 0) {
    const remainder = (index - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter; // 65 -> 'A'
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

// 全局对象，用于存储每个容器的选中选项
const selectedOptionsMapContainer1 = {}; // For 'dropdown-container1'
const selectedOptionsMapContainer2 = {}; // For 'dropdown-container2'

// 全局数组，用于存储所有下拉菜单实例
const dropdownInstances = [];

//---------------------------------------获取每个字段的唯一值 ----- 并创建HTML 下拉菜单----------------------------
async function CreateDropList(event = null) {
  console.log("test test test");
  await Excel.run(async (context) => {
    // 获取 "Data" 工作表
    const sheet = context.workbook.worksheets.getItem("Data");
    // 获取工作表的已用范围
    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount, columnCount");
    await context.sync();



    // 读取字段名，假设字段名在第二行
    const headerRange = sheet.getRangeByIndexes(1, 1, 1, usedRange.columnCount - 1);
    headerRange.load("values");
    await context.sync();

    // 读取控制信息（第一行）
    const controlRange = sheet.getRangeByIndexes(0, 1, 1, usedRange.columnCount - 1);
    controlRange.load("values");
    controlRange.load("address");
    await context.sync();

    let headers = headerRange.values[0];
    let uniqueValues = {};
    let controls = controlRange.values[0];

    // 初始化每个字段的Set
    headers.forEach((header) => {
      uniqueValues[header] = new Set();
    });

    // 读取数据，从第三行开始直到最后
    const dataRange = sheet.getRangeByIndexes(2, 1, usedRange.rowCount - 2, usedRange.columnCount - 1);
    dataRange.load("values");
    await context.sync();

    // 遍历每一列，提取唯一值
    for (let colIndex = 0; colIndex < headers.length; colIndex++) {
      // 使用map提取每一列的值，并应用Set去重
      let columnData = dataRange.values.map((row) => row[colIndex]);
      uniqueValues[headers[colIndex]] = new Set(columnData);
    }

    // 将Set转换为数组
    for (let header of headers) {
      uniqueValues[header] = Array.from(uniqueValues[header]);
    }

    const worksheetNames = context.workbook.worksheets.load("items/name");
    await context.sync();

    const sheetNames = worksheetNames.items.map(ws => ws.name);
    let hasSelectedValue1 = sheetNames.includes("SelectedValue1");
    let hasSelectedValue2 = sheetNames.includes("SelectedValue2");
    console.log("Create DropList 1");

    //如果SelectedValue1 和 SelectedValue2 有一个工作表不存在，或者 restoreOptions 按钮按下的时候，执行下面的还原下拉菜单代码，否则不执行
    //确保 event 对象包含 target 属性，防止直接访问 event.target.id 抛出错误。
    if (!hasSelectedValue1 || !hasSelectedValue2 || (event && event.target && (event.target.id === "restoreOptions"))) { 
      console.log("Create DropList 2");
        // console.log("event.target.id is " + event.target.id);
        //当时按键restoreOptions按下的时候，不执行下面重新生成da
      // if (!(event && event.target && (event.target.id === "restoreOptions"))) {   //确保 event 对象包含 target 属性，防止直接访问 event.target.id 抛出错误。
          // 检查并创建 "SelectedValue1" 和 "SelectedValue2" 工作表
          // const worksheetNames = context.workbook.worksheets.load("items/name");
          // await context.sync();

          // const sheetNames = worksheetNames.items.map(ws => ws.name);
          let selectedValue1Sheet, selectedValue2Sheet;

          if (!sheetNames.includes("SelectedValue1")) {
            selectedValue1Sheet = context.workbook.worksheets.add("SelectedValue1");
            await context.sync();
            await DoNotChangeCellWarning("SelectedValue1");
          } else {
            selectedValue1Sheet = context.workbook.worksheets.getItem("SelectedValue1");
            selectedValue1Sheet.getUsedRange().clear(); // 清空工作表数据
          }

          if (!sheetNames.includes("SelectedValue2")) {
            selectedValue2Sheet = context.workbook.worksheets.add("SelectedValue2");
            await context.sync();
            await DoNotChangeCellWarning("SelectedValue2");
          } else {
            selectedValue2Sheet = context.workbook.worksheets.getItem("SelectedValue2");
            selectedValue2Sheet.getUsedRange().clear(); // 清空工作表数据
          }

          await context.sync();

          // 将字段名和唯一值写入 "SelectedValue1" 和 "SelectedValue2" 工作表，仅当控制值为 "dimension"
          let colIndex = 0;
          for (let index = 0; index < headers.length; index++) {
            // 仅当 controls[index] === "Dimension" 时写入
            if (controls[index] === "Dimension") {
              // 调用通用函数，将 0-based colIndex 转为列字母
              const columnLetter = toColumnLetter(colIndex); 
        
              // 1) 写入字段名到第一行 (无 sync)
              selectedValue1Sheet.getRange(`${columnLetter}1`).values = [[headers[index]]];
              selectedValue2Sheet.getRange(`${columnLetter}1`).values = [[headers[index]]];
        
              // 2) 写入唯一值从第二行开始 (无 sync)
              const uniqueValuesLength = uniqueValues[headers[index]].length;
              const startAddress = `${columnLetter}2`;
              const endAddress   = `${columnLetter}${uniqueValuesLength + 1}`;
              const fullAddress  = `${startAddress}:${endAddress}`;
              
              selectedValue1Sheet.getRange(fullAddress).values =
                uniqueValues[headers[index]].map(value => [value]);
        
              selectedValue2Sheet.getRange(fullAddress).values =
                uniqueValues[headers[index]].map(value => [value]);
        
              colIndex++; // 仅当写入数据时增加列索引
            }
          }
          // ★ 最后统一一次提交
          await context.sync();
          console.log("所有写入已一次性同步到 Excel。");
    } else {
      // 如果任一工作表不存在，打印日志并跳过此逻辑段
      console.log("工作表 'SelectedValue1' 或 'SelectedValue2' 不存在，跳过此逻辑段");
    }
      await context.sync();


    // 清空旧的下拉菜单内容
    const dropdownContainer1 = document.getElementById("dropdown-container1");
    const dropdownContainer2 = document.getElementById("dropdown-container2");

    dropdownContainer1.innerHTML = ""; // 清空第一个容器内容
    dropdownContainer2.innerHTML = ""; // 清空第二个容器内容

    // 在这里调用创建下拉菜单的函数，并传递 controls 数组
    await createDropdownMenus(uniqueValues, headers, controls);

    // 激活工作表（如果需要）
    // WaterfallSheet.activate();
    await context.sync();
  });
}


// 封装函数用于创建下拉菜单
async function createDropdownMenus(uniqueValues, headers, controls) {
  // 创建映射，将每个字段对应的选项数据准备好
  console.log("createDropdownMenus 0");
  const optionsDataMap = {};
  headers.forEach((header) => {
    optionsDataMap[header] = uniqueValues[header].map((value) => ({
      value: value,
      label: value,
    }));
  });

  console.log("createDropdownMenus 1");
  // 获取页面上的两个容器，分别用于存放两个下拉菜单
  const dropdownContainer1 = document.getElementById("dropdown-container1");
  const dropdownContainer2 = document.getElementById("dropdown-container2");

  // **移除旧的容器标签（如果存在）**
  const oldLabel1 = document.querySelector("label[for='dropdown-container1']");
  const oldLabel2 = document.querySelector("label[for='dropdown-container2']");
  if (oldLabel1) oldLabel1.remove();
  if (oldLabel2) oldLabel2.remove();

  // **动态创建并添加容器标签**
  const containerLabel1 = document.createElement("label");
  containerLabel1.setAttribute("for", "dropdown-container1");
  containerLabel1.classList.add("container-label");
  containerLabel1.textContent = "Base";

  const containerLabel2 = document.createElement("label");
  containerLabel2.setAttribute("for", "dropdown-container2");
  containerLabel2.classList.add("container-label");
  containerLabel2.textContent = "Target";

  // 将标签插入到容器之前
  dropdownContainer1.parentNode.insertBefore(containerLabel1, dropdownContainer1);
  dropdownContainer2.parentNode.insertBefore(containerLabel2, dropdownContainer2);

  console.log("createDropdownMenus 2");
  // 遍历每个字段，为两个容器创建相同的下拉菜单
  headers.forEach((header, index) => {
    // 仅当 controls[index] === "Dimension" 时才创建下拉菜单
    if (controls[index] === "Dimension") {
      console.log("createDropdownMenus 3");
      const optionsData = optionsDataMap[header]; // 获取当前字段的选项数据
      console.log("createDropdownMenus 4");
      // 在第一个容器中创建下拉菜单

      createDropdown(dropdownContainer1, optionsData, header, selectedOptionsMapContainer1);

      // 在第二个容器中创建相同的下拉菜单
      createDropdown(dropdownContainer2, optionsData, header, selectedOptionsMapContainer2);

    }
  });
}

let isDropdownOpening = false; // 初始化 不然点击选项框以外部分不能关闭选项框

// 假设在 createDropdown 中增加如下逻辑：
// 创建下拉菜单实例时，附加额外信息（header, containerId, optionsData, 以及更新UI函数）
function createDropdown(container, optionsData, header, selectedOptionsMap) {
  console.log("createDropdown 1");
  //let isDropdownOpening = false;

  const customSelect = document.createElement("div"); // 创建自定义选择框容器
  customSelect.classList.add("custom-select"); // 添加样式类

  const dropdownLabel = document.createElement("label"); // 创建标签显示字段名
  dropdownLabel.classList.add("dropdown-label");
  dropdownLabel.textContent = header; // 将字段名作为标签文本

  const selectBox = document.createElement("input"); // 创建输入框作为下拉选择框的显示区域
  selectBox.type = "text";
  selectBox.classList.add("select-box");
  selectBox.placeholder = "全选"; // 默认占位文本
  selectBox.readOnly = true; // 设置为只读，防止弹出键盘（在移动设备上）

  const dropdown = document.createElement("div"); // 创建下拉选项容器
  dropdown.classList.add("dropdown");

  const dropdownHeader = document.createElement("div"); // 创建下拉菜单头部
  dropdownHeader.classList.add("dropdown-header");

  const confirmBtn = document.createElement("button"); // 创建确认按钮
  confirmBtn.classList.add("confirm-btn");
  confirmBtn.textContent = "确认";
  confirmBtn.disabled = true; // 默认禁用，直到有选项被选择

  const cancelBtn = document.createElement("button"); // 创建取消按钮
  cancelBtn.classList.add("cancel-btn");
  cancelBtn.textContent = "取消";

  dropdownHeader.appendChild(confirmBtn); // 将确认按钮添加到头部
  dropdownHeader.appendChild(cancelBtn); // 将取消按钮添加到头部

  const optionsList = document.createElement("ul"); // 创建选项列表
  optionsList.classList.add("options-list");

  dropdown.appendChild(dropdownHeader); // 将头部添加到下拉菜单
  dropdown.appendChild(optionsList); // 将选项列表添加到下拉菜单

  customSelect.appendChild(selectBox); // 将输入框添加到自定义选择框
  customSelect.appendChild(dropdown); // 将下拉菜单添加到自定义选择框

  container.appendChild(dropdownLabel); // 将标签添加到指定的容器
  container.appendChild(customSelect); // 将自定义选择框添加到指定的容器
  console.log("createDropdown 2");
  // 初始化选项值为字符串形式，避免数字和字符串类型问题
  let selectedOptions = optionsData.map((option) => String(option.value)); // 初始化选中的选项数据
  let tempSelectedOptions = [...selectedOptions]; // 临时存储选中的选项数据
  console.log("createDropdown 3");
  // **新增代码标识：添加更新UI的方法，便于后续批量更新选中状态**
  function setSelection(newSelection) {
    tempSelectedOptions = [...newSelection];
    selectedOptions = [...newSelection];
    updateCheckboxes();
    updateSelectBoxText();
  }
  // **新增代码结束**

  // 创建下拉选项内容
  function createOptions() {
    optionsList.innerHTML = ""; // 清空选项列表内容

    const selectAllOption = document.createElement("li"); // 创建全选选项
    selectAllOption.innerHTML = `
            <label>
                <input type="checkbox" class="option-checkbox" value="selectAll">
                全选
            </label>
        `;
    optionsList.appendChild(selectAllOption); // 将全选选项添加到列表

    optionsData.forEach((option) => {
      // 遍历每个选项数据，生成对应的列表项
      const li = document.createElement("li");
      li.innerHTML = `
                <label>
                    <input type="checkbox" class="option-checkbox" value="${String(option.value)}">
                    ${option.label}
                </label>
            `;
      optionsList.appendChild(li); // 将生成的选项添加到选项列表
    });

    updateCheckboxes(); // 更新复选框状态
  }
  console.log("createDropdown 4");

  
  // 更新选项复选框状态
  function updateCheckboxes() {
    const checkboxes = optionsList.querySelectorAll(".option-checkbox");

    // 初始状态上面已经定义了 tempSelectedOptions.length === optionsData.length;
    checkboxes.forEach((checkbox) => {
      if (checkbox.value === "selectAll") {
        // 全选复选框
        checkbox.checked = tempSelectedOptions.length === optionsData.length;
      } else {
        // 其他复选框
        checkbox.checked = tempSelectedOptions.includes(String(checkbox.value));
      }
    });

    updateConfirmButton(); // 更新确认按钮状态
  }

  // 更新选择框的显示文本
  function updateSelectBoxText() {
    if (selectedOptions.length === optionsData.length) {
      selectBox.placeholder = "全选"; // 全选状态
    } else if (selectedOptions.length === 1) {
      const selectedOption = optionsData.find((option) => String(option.value) === selectedOptions[0]);
      selectBox.placeholder = selectedOption.label; // 单选状态
    } else if (selectedOptions.length > 1) {
      selectBox.placeholder = "Multiple Selection"; // 多选状态
    } else {
      selectBox.placeholder = ""; // 无选择
    }
  }

  // 更新确认按钮状态
  function updateConfirmButton() {
    confirmBtn.disabled = tempSelectedOptions.length === 0; // 当没有选择项时禁用确认按钮
  }

  // 重置选项列表的显示状态
  function resetOptionsDisplay() {
    const options = optionsList.querySelectorAll("li");
    options.forEach((option) => {
      option.style.display = ""; // 恢复所有选项的显示
    });
  }

  // 在创建选项或下拉菜单展开后调用。在下拉菜单展开时，调用函数检查内容是否溢出，动态设置 overflow-y 属性。
  function checkOverflow() {
    const optionsList = dropdown.querySelector(".options-list");
    if (optionsList.scrollHeight > optionsList.clientHeight) {
      optionsList.style.overflowY = "auto";
    } else {
      optionsList.style.overflowY = "hidden";
    }
  }

  function openDropdown() {
    // 如果下拉菜单已经是打开状态，则不需要再次打开
    if (dropdown.classList.contains("show")) {
      return;
    }

    isDropdownOpening = true; // 设置标志位，表示正在打开下拉菜单

    // 移除可能干扰的任何内联 max-height 样式
    dropdown.style.maxHeight = "";

    // 延迟执行下拉菜单的打开逻辑
    setTimeout(() => {
      // 添加 'show' 类，使下拉菜单可见
      dropdown.classList.add("show");

      // 强制重绘确保浏览器应用样式变化
      dropdown.offsetHeight; // 强制重绘

      dropdown.style.visibility = "visible";
      // 设定最大允许高度（例如 200px）
      const maxAllowedHeight = 200;

      // 计算内容实际高度
      const contentHeight = dropdown.scrollHeight;

      // 设置 'max-height' 为内容高度和最大允许高度的较小值
      const finalHeight = Math.min(contentHeight, maxAllowedHeight);
      dropdown.style.maxHeight = finalHeight + "px";
      // 在展开下拉菜单后，滚动页面以使下拉菜单完全可见
      dropdown.scrollIntoView({ block: "nearest", inline: "nearest", behavior: "smooth" });

      // 检查溢出
      checkOverflow();

      // 重置滚动条位置到顶部
      dropdown.scrollTop = 0;

      // 设置 z-index 确保下拉菜单在最上层
      dropdown.style.zIndex = "9999";

      // 下拉菜单已打开，重置标志位
      isDropdownOpening = false;
    }, 0); // 使用 0 延迟，确保页面滚动完成后再执行
  }

  // 关闭下拉菜单
  function closeDropdown() {
    tempSelectedOptions = [...selectedOptions]; // 恢复选中的选项
    updateCheckboxes(); // 更新复选框状态
    dropdown.classList.remove("show"); // 隐藏下拉菜单
    selectBox.value = ""; // 清空输入框
    resetOptionsDisplay(); // 重置选项列表显示
    dropdown.style.zIndex = ""; // 清除 z-index
    // 重置 'max-height' 为 0
    dropdown.style.maxHeight = "0";
    dropdown.style.visibility = "hidden";
  }

  function closeOtherDropdowns() {
    dropdownInstances.forEach((instance) => {
      if (instance !== dropdownInstance && instance.isOpen()) {
        instance.closeDropdown();
      }
    });
  }

  // 使用 mousedown 事件
  selectBox.addEventListener("mousedown", function (e) {
    e.preventDefault(); // 阻止默认行为
    e.stopPropagation(); // 阻止事件冒泡

    // 关闭其他下拉菜单
    closeOtherDropdowns();

    // 打开下拉菜单
    openDropdown();
  });

  // 阻止 selectBox 的 focus 事件
  selectBox.addEventListener("focus", function (e) {
    e.preventDefault();
    e.stopPropagation();
  });

  // 输入框的输入事件，用于过滤选项
  selectBox.addEventListener("input", function () {
    const filter = selectBox.value.toLowerCase();
    const options = optionsList.querySelectorAll("li");

    options.forEach((option) => {
      const label = option.textContent.toLowerCase();
      option.style.display = label.includes(filter) ? "" : "none"; // 根据输入过滤选项显示
    });
  });

  // 选项列表的变化事件，用于更新选择状态
  optionsList.addEventListener("change", function (e) {
    const checkbox = e.target;
    if (checkbox.classList.contains("option-checkbox")) {
      if (checkbox.value === "selectAll") {
        // 全选复选框
        tempSelectedOptions = checkbox.checked ? optionsData.map((option) => String(option.value)) : [];
      } else {
        // 单个选项复选框
        if (checkbox.checked) {
          tempSelectedOptions.push(String(checkbox.value));
        } else {
          tempSelectedOptions = tempSelectedOptions.filter((value) => value !== String(checkbox.value));
        }

        const selectAllCheckbox = optionsList.querySelector('input[value="selectAll"]');
        selectAllCheckbox.checked = tempSelectedOptions.length === optionsData.length; // 更新全选状态
      }
      updateCheckboxes(); // 更新复选框状态
    }
  });

  // 确认按钮的点击事件
  confirmBtn.addEventListener("click", async function () {
    if (confirmBtn.disabled) return;
    selectedOptions = [...tempSelectedOptions]; // 更新选中的选项
    selectedOptionsMap[header] = [...selectedOptions]; // 更新全局的选项映射
    updateSelectBoxText(); // 更新选择框的显示文本
    closeDropdown();
    console.log(`已确认选择（${header}）：`, selectedOptions); // 输出确认选择的结果

    // **修改代码标识：根据 container 的 id 来决定使用哪个工作表**
    let sheetName = (container.id === "dropdown-container1") ? "SelectedValue1" : "SelectedValue2";

    // 调用自定义函数SaveSelectedValue 将数据存储到相应的工作表中
    await SaveSelectedValue(header, selectedOptionsMap, sheetName);
    // **修改代码结束**

    await refreshBridge();
  });
  console.log("createDropdown 5");

  // 取消按钮的点击事件
  cancelBtn.addEventListener("click", function () {
    closeDropdown(); // 关闭下拉菜单
  });

  // 点击下拉菜单内部时，阻止事件冒泡，防止关闭下拉菜单
  dropdown.addEventListener("mousedown", function (e) {
    e.stopPropagation();
  });


  // **新增代码标识：创建下拉菜单实例并添加到全局数组，增加更新UI方法的引用**
  const dropdownInstance = {
    customSelect: customSelect,
    closeDropdown,
    isOpen: () => dropdown.classList.contains("show"),
    header,
    containerId: container.id,
    optionsData,
    setSelection, // 新增的更新UI方法
    getAllOptions: () => optionsData.map(o => String(o.value)),
  };
  console.log("createDropdown 5.4");
  dropdownInstances.push(dropdownInstance);
  // **新增代码结束**
  console.log("createDropdown 5.5");
  createOptions(); // 创建选项内容
  console.log("createDropdown 5.6");
  updateSelectBoxText(); // 更新选择框文本
  console.log("createDropdown 5.7");
}

// 全局点击事件监听器，当点击页面其他区域时，关闭所有下拉菜单
document.addEventListener("mousedown", function (e) {
  if (isDropdownOpening) {
    // 正在打开下拉菜单，忽略此次点击事件
    isDropdownOpening = false; // 重置标志位
    return;
  }

  // 遍历所有下拉菜单实例，关闭点击区域外的下拉菜单
  dropdownInstances.forEach((instance) => {
    if (!instance.customSelect.contains(e.target)) {
      instance.closeDropdown();
    }
  });
});


// **新增代码标识：根据两个工作表更新所有下拉菜单的选中状态**
async function updateDropdownsFromSelectedValues() {
  console.log("Enter updateDropdownsFromSelectedValues 1");
  await Excel.run(async (context) => {
    let selectedValueSheet1, selectedValueSheet2;
    try {
      selectedValueSheet1 = context.workbook.worksheets.getItem("SelectedValue1");
      selectedValueSheet1.load("name");
    } catch (e) {
      return; // 不存在则直接返回，不执行后续操作
    }

    try {
      selectedValueSheet2 = context.workbook.worksheets.getItem("SelectedValue2");
      selectedValueSheet2.load("name");
    } catch (e) {
      return; // 不存在则直接返回，不执行后续操作
    }

    await context.sync();
    console.log("Enter updateDropdownsFromSelectedValues 2");
    // 如果能执行到这里，说明 SelectedValue1 和 SelectedValue2 都存在
    const usedRange1 = selectedValueSheet1.getUsedRangeOrNullObject();
    const usedRange2 = selectedValueSheet2.getUsedRangeOrNullObject();
    usedRange1.load("values,rowCount,columnCount,address");
    usedRange2.load("values,rowCount,columnCount,address");
    await context.sync();

    console.log("usedRange1 address is" + usedRange1.address);
    console.log("usedRange2 address is" + usedRange2.address)
    console.log("Enter updateDropdownsFromSelectedValues 3");

    let headersSV1 = [];
    let dataSV1 = {};
    if (!usedRange1.isNullObject && usedRange1.rowCount > 0) {
      headersSV1 = usedRange1.values[0];
      headersSV1.forEach((h, idx) => {
        let colData = usedRange1.values.slice(1).map(r => r[idx]).filter(v => v !== null && v !== undefined && v !== "");  // 过滤掉 null, undefined 和空字符串
        dataSV1[h] = colData;
      });
    }
    console.log("Enter updateDropdownsFromSelectedValues 4");
    console.log("headersSV1 is ");
    console.log(JSON.stringify(headersSV1, null, 2));

    console.log("dataSV1 is ")
    console.log(JSON.stringify(dataSV1, null, 2));

    let headersSV2 = [];
    let dataSV2 = {};
    if (!usedRange2.isNullObject && usedRange2.rowCount > 0) {
      headersSV2 = usedRange2.values[0];
      headersSV2.forEach((h, idx) => {
        let colData = usedRange2.values.slice(1).map(r => r[idx]).filter(v => v !== null && v !== undefined  && v !== ""); // 过滤掉 null, undefined 和空字符串
        dataSV2[h] = colData;
      });
    }
    console.log("headersSV2 is ");
    console.log(JSON.stringify(headersSV2, null, 2));

    console.log("dataSV2 is ")
    console.log(JSON.stringify(dataSV2, null, 2));

    console.log("Enter updateDropdownsFromSelectedValues 5");
    // 全选 dropdown-container1 和 dropdown-container2 中所有Dimension类型的下拉菜单
    const container1Dropdowns = dropdownInstances.filter(d => d.containerId === "dropdown-container1");
    const container2Dropdowns = dropdownInstances.filter(d => d.containerId === "dropdown-container2");

    container1Dropdowns.forEach(d => {
      const allOptions = d.getAllOptions();
      d.setSelection(allOptions);
    });

    container2Dropdowns.forEach(d => {
      const allOptions = d.getAllOptions();
      d.setSelection(allOptions);
    });

    console.log("Enter updateDropdownsFromSelectedValues 5");
    // 根据SelectedValue1的数据更新dropdown-container1
    headersSV1.forEach(h => {
      let dropdown = container1Dropdowns.find(d => d.header === h);
      console.log("dropdown is ")
      console.log(JSON.stringify(dropdown, null, 2));
      console.log("dataSV1[h] is ")
      console.log(JSON.stringify(dataSV1[h], null, 2));
      console.log("dropdown && dataSV1[h] is ");
      console.log(JSON.stringify(dropdown && dataSV1[h], null, 2));
      // if (dropdown && dataSV1[h]) {
      //   dropdown.setSelection(dataSV1[h]);
      // }
      if (dropdown && dataSV1[h]) {
        // 过滤出合法值
        const validSelection = dataSV1[h].filter(value => 
          dropdown.optionsData.some(option => option.value === value)
        );
        console.log(`Setting selection for ${h}: `, validSelection);
        dropdown.setSelection(validSelection); // 设置选中项
      }
    });


    console.log("Enter updateDropdownsFromSelectedValues 6");
    // 根据SelectedValue2的数据更新dropdown-container2
    headersSV2.forEach(h => {
      let dropdown = container2Dropdowns.find(d => d.header === h);
      console.log("dropdown is ")
      console.log(JSON.stringify(dropdown, null, 2));
      console.log("dataSV2[h] is ")
      console.log(JSON.stringify(dataSV2[h], null, 2));
      console.log("dropdown && dataSV2[h] is ");
      console.log(JSON.stringify(dropdown && dataSV2[h], null, 2));
      if (dropdown && dataSV2[h]) {
        dropdown.setSelection(dataSV2[h]);
      }
    });

    console.log("Enter updateDropdownsFromSelectedValues 7");
    await context.sync();

  });
}
// **新增代码结束**

// 修改后的SaveSelectedValue函数，可根据sheetName参数动态创建或写入指定工作表
async function SaveSelectedValue(header, selectedOptionsMap, sheetName) {
  await Excel.run(async (context) => {
    let selectedValueSheet;

    // 尝试获取指定sheetName的工作表
    try {
      selectedValueSheet = context.workbook.worksheets.getItem(sheetName);
      selectedValueSheet.load("name");
      await context.sync();
    } catch (err) {
      // 如果不存在则新建
      selectedValueSheet = context.workbook.worksheets.add(sheetName);
      await context.sync();
    }

    // 读取第一行，用于确定 header 所在列
    const usedRange = selectedValueSheet.getUsedRangeOrNullObject();
    usedRange.load("rowCount, columnCount, values");
    await context.sync();

    let headersInSheet = [];
    let colCount = 0;
    if (!usedRange.isNullObject) {
      colCount = usedRange.columnCount;
      if (usedRange.rowCount > 0) {
        headersInSheet = usedRange.values[0]; // 第一行
      }
    }

    // 在 headersInSheet 中查找当前 header 的位置
    let headerIndex = headersInSheet.indexOf(header);

    // 如果没有找到该 header，则在末尾添加一列
    if (headerIndex === -1) {
      headerIndex = colCount; // 新列索引
      // 在第一行的 headerIndex 列写入 header
      selectedValueSheet.getRangeByIndexes(0, headerIndex, 1, 1).values = [[header]];
    }

    // 确保加载最新的 usedRange
    const updatedUsedRange = selectedValueSheet.getUsedRange();
    updatedUsedRange.load("rowCount");
    await context.sync();

    // 清空该 header 列（第一行以下）的已有数据
    if (updatedUsedRange.rowCount > 1) {
      const oldDataRange = selectedValueSheet.getRangeByIndexes(1, headerIndex, updatedUsedRange.rowCount - 1, 1);
      oldDataRange.clear();
    }

    // 写入新的数据，从第二行开始写
    const newData = selectedOptionsMap[header].map(value => [value]);
    if (newData.length > 0) {
      selectedValueSheet.getRangeByIndexes(1, headerIndex, newData.length, 1).values = newData;
    }

    await context.sync();
  });
}


// async function onChange(event) {

//         await Excel.run(async (context) => {
//             if (isFirstRow(event.address)) {
//                 CreateDropList();
//                 createCombinePivotTable();
//                 await context.sync();
//             }
//         });
// }  
          
// //根据用户按钮选择是否根据bridge data 的变化执行代码
// async function onChange(event) {
//   await Excel.run(async (context) => {
//       console.log("Enter onChange");
//       // 判断是否有重复的 "Key"
//       if (await hasDuplicateKeyInFirstRow(context)) {
//           // 如果有重复的 "Key"，已在函数内部处理了警告逻辑，直接返回
//           return;
//       }

//       // 判断是否有重复的 "Result"
//       if (await hasDuplicateResultInFirstRow(context)) {
//         // 如果有重复的 "Result"，已在函数内部处理了警告逻辑，直接返回
//         return;
//     }

//       // 如果上面的前提条件没有发生
//       let changeResult = await isFirstRow(event.address);
//       if (changeResult) {
//           // 显示 waterfall 提示
//           const waterfallPrompt = document.getElementById("waterfallPrompt");
//           const modalOverlay = document.getElementById("modalOverlay");
//           const container = document.querySelector(".container");

//           // 显示模态遮罩和提示框
//           waterfallPrompt.style.display = "flex"; //必须改成flex才能使用对应的样式
//           modalOverlay.style.display = "block";
//           container.classList.add("disabled"); // 禁用其他容器
//           // 禁用其他容器，但保留 waterfallPrompt
//           // waterfallPrompt.style.pointerEvents = "auto"; // 启用交互
//           // waterfallPrompt.style.zIndex = "1100"; // 保证提示框层级

//           // 滚动到提示并聚焦到 "Yes" 按钮
//           waterfallPrompt.scrollIntoView({ behavior: "smooth", block: "center" });
//           document.getElementById("confirmWaterfall").focus(); // Set focus on the "Yes" button
        
//           // 处理 "Yes" 按钮点击
//           document.getElementById("confirmWaterfall").onclick = async function () {
//               await Excel.run(async (context) => {
//                   await CreateDropList();
//                   await createCombinePivotTable();
//                   await context.sync();
//               });
//               // 隐藏提示框
//               waterfallPrompt.style.display = "none";
//               modalOverlay.style.display = "none";
//               container.classList.remove("disabled"); // 恢复其他容器交互
//           };

//           // 处理 "No" 按钮点击
//           document.getElementById("cancelWaterfall").onclick = function () {
//               waterfallPrompt.style.display = "none";
//               modalOverlay.style.display = "none";
//               container.classList.remove("disabled"); // 恢复其他容器交互
//           };
//       }
//   }).catch(function (error) {
//       console.error("Error: " + error);
//   });
// }


//-------检查第三行开始的数据类型是否是正确的--------
async function checkBridgeDataHeadersAndValues(context) {
  console.log("checkBridgeDataHeadersAndValues start")
  const workbook = context.workbook;
  const sheet = workbook.worksheets.getItem("Data");
  let range = sheet.getUsedRange();
  const firstRowRange = range.getRow(0); // 获取第一行
  const secondRowRange = range.getRow(1); // 获取第二行
  const thirdRowRange = range.getRow(2); // 获取第三行
  firstRowRange.load("values");//加载第一行的值
  secondRowRange.load("values"); // 加载第二行的值
  thirdRowRange.load("values"); // 加载第三行的值
  await context.sync(); // 确保加载完成
  console.log("Check Data Step 1");

  const firstRowValues = firstRowRange.values[0];
  const secondRowValues = secondRowRange.values[0];
  const thirdRowValues = thirdRowRange.values[0];
  
  // 验证第二行和第三行的值
  for (let i = 0; i < secondRowValues.length; i++) {

      const headerType = firstRowValues[i];
      const header = secondRowValues[i];
      const dataValue = thirdRowValues[i];

      if (["SumY", "Result", "SumN"].includes(headerType)) {
          console.log("Check Data Step 2");
          if (isNaN(dataValue)) {
              console.log("CheckData is NaN");
              const errorMessage = `${header} 的类型为数值相关，因此从第三行开始必须是数值类型，检测到非数值数据。`;

              // 显示提示框并等待用户点击确认按钮
              const modalOverlay = document.getElementById("modalOverlay");
              const keyWarningPrompt = document.getElementById("keyWarningPrompt");
              const container = document.querySelector(".container");

              const warningElement = document.querySelector("#keyWarningPrompt .waterfall-message");
              warningElement.textContent = errorMessage;

              modalOverlay.style.display = "block";
              keyWarningPrompt.style.display = "flex";
              container.classList.add("disabled");

              await new Promise((resolve) => {
                  const confirmButton = document.getElementById("confirmKeyWarning");
                  confirmButton.addEventListener(
                      "click",
                      function () {
                          keyWarningPrompt.style.display = "none";
                          modalOverlay.style.display = "none";
                          container.classList.remove("disabled");
                          resolve(); // 继续执行
                      },
                      { once: true } // 确保事件只触发一次
                  );
              });

              return true; // 返回 true 表示检测到非数值数据
          }
      }
  }

  console.log("验证通过，所有相关数据均为数值类型。");
  return false; // 返回 false 表示所有数据均符合要求
}


// 检查 Data 工作表第一行的值是否都存在"Dimension", "Key", "SumY", "Result"
async function checkRequiredHeaders(context) {
  const workbook = context.workbook;
  const sheet = workbook.worksheets.getItem("Data");
  let range = sheet.getUsedRange();
  const firstRowRange = range.getRow(0); // 获取第一行
  firstRowRange.load("values"); // 加载第一行的值
  await context.sync(); // 确保加载完成

  // 定义必需的标题
  const requiredHeaders = ["Dimension", "Key", "SumY", "Result"];
  const firstRowValues = firstRowRange.values[0];

  // 检查缺失的标题
  const missingHeaders = requiredHeaders.filter(header => !firstRowValues.includes(header));

  if (missingHeaders.length > 0) {
      // 显示第一个缺失标题的警告信息
      const missingHeadersList = missingHeaders.join(", ");
      const warningMessage = `在 "Data" 表的第一行中，缺少以下值：${missingHeadersList}。这些数据类型是必须的。`;
      console.error(warningMessage);

      // 显示缺失标题的提示框
      const modalOverlay = document.getElementById("modalOverlay");
      const keyWarningPrompt = document.getElementById("keyWarningPrompt");
      const container = document.querySelector(".container");

      const warningElement = document.querySelector("#keyWarningPrompt .waterfall-message");
      warningElement.textContent = warningMessage;

      modalOverlay.style.display = "block";
      keyWarningPrompt.style.display = "flex";
      container.classList.add("disabled");

      await new Promise((resolve) => {
          const confirmButton = document.getElementById("confirmKeyWarning");
          confirmButton.addEventListener(
              "click",
              function () {
                  keyWarningPrompt.style.display = "none";
                  modalOverlay.style.display = "none";
                  container.classList.remove("disabled");
                  resolve(); // 继续执行
              },
              { once: true } // 确保事件只触发一次
          );
      });

      return true; // 表示存在缺失的标题
  }

  console.log("所有必需的标题都存在。");
  return false; // 所有标题都存在
}


// 检查 Data 工作表第一行的 Key 值
//检查是否有两个key在bridge data 第一行
async function hasDuplicateKeyInFirstRow(context) {
  const workbook = context.workbook;
  const sheet = workbook.worksheets.getItem("Data");
  let range = sheet.getUsedRange();
  const firstRowRange = range.getRow(0); // 获取第一行
  firstRowRange.load("values"); // 加载第一行的值
  await context.sync(); // 确保加载了第一行的值

  // 获取第一行的值
  const firstRowValues = firstRowRange.values[0];
  const keyCount = firstRowValues.filter(value => value === "Key").length;

  // 输出 keyCount 以检查结果
  console.log("keyCount is " + keyCount);

  if (keyCount === 0 || keyCount > 1) {
      // 如果没有 "Key" 或有多个 "Key"，显示警告消息并等待用户确认
      console.log("Invalid number of 'Key' values in the first row.");
      const keyWarningPrompt = document.getElementById("keyWarningPrompt");
      const modalOverlay = document.getElementById("modalOverlay");
      const container = document.querySelector(".container");

      // 动态更新警告消息
      const warningMessage = document.querySelector("#keyWarningPrompt .waterfall-message");
      if (keyCount === 0) {
          warningMessage.textContent = "Data工作表第一行必须有一个单元格的值是Key。";
      } else {
          warningMessage.textContent = "Data工作表第一行只能有一个单元格的值是Key，修改并保留唯一的单元格值为Key。";
      }
    

      // 显示模态遮罩和提示框
      modalOverlay.style.display = "block";
      keyWarningPrompt.style.display = "flex";
      container.classList.add("disabled");

      // keyWarningPrompt.style.display = "block";

      // 等待用户点击确认按钮
      await new Promise((resolve) => {
        const confirmButton = document.getElementById("confirmKeyWarning");

        confirmButton.addEventListener(
            "click",
            function () {
                keyWarningPrompt.style.display = "none";
                modalOverlay.style.display = "none";
                container.classList.remove("disabled");
                resolve(); // 继续 Promise
            },
            { once: true } // 确保事件只触发一次
        );
      });

      return true; // 有重复的 "Key"
  }

  return false; // 没有重复的 "Key"
}

//检查是否有两个Result在bridge data 第一行
//检查 Data 工作表第一行的 Result 值
async function hasDuplicateResultInFirstRow(context) {
  const workbook = context.workbook;
  const sheet = workbook.worksheets.getItem("Data");
  let range = sheet.getUsedRange();
  const firstRowRange = range.getRow(0); // 获取第一行
  firstRowRange.load("values"); // 加载第一行的值
  await context.sync(); // 确保加载了第一行的值

  // 获取第一行的值
  const firstRowValues = firstRowRange.values[0];
  const ResultCount = firstRowValues.filter((value) => value === "Result").length;

  console.log("ResultCount is " + ResultCount);

  if (ResultCount === 0 || ResultCount > 1) {
      // 如果没有 "Result" 或有多个 "Result"，显示警告消息并等待用户确认
      console.log("Invalid number of 'Result' values in the first row.");
      const resultWarningPrompt = document.getElementById("ResultWarningPrompt");
      const modalOverlay = document.getElementById("modalOverlay");
      const container = document.querySelector(".container");

      // 动态更新警告消息
      const warningMessage = document.querySelector("#ResultWarningPrompt .waterfall-message");
      if (ResultCount === 0) {
          warningMessage.textContent = "Data工作表第一行必须有一个单元格的值是Result。";
      } else {
          warningMessage.textContent = "Data工作表第一行只能有一个单元格的值是Result，修改并保留唯一的单元格值为Result。";
      }

      // 显示模态遮罩和提示框
      modalOverlay.style.display = "block";
      resultWarningPrompt.style.display = "flex";
      container.classList.add("disabled");

      // 等待用户点击确认按钮
      await new Promise((resolve) => {
          const confirmButton = document.getElementById("confirmResultWarning");

          confirmButton.addEventListener(
              "click",
              function () {
                  resultWarningPrompt.style.display = "none";
                  modalOverlay.style.display = "none";
                  container.classList.remove("disabled");
                  resolve(); // 继续 Promise
              },
              { once: true } // 确保事件只触发一次
          );
      });

      return true; // 无法通过检查
  }

  return false; // 检查通过
}


async function isFirstRow(address) {
  return await Excel.run(async (context) => {
    console.log("isFirstRow address is " + address);

      let worksheet = context.workbook.worksheets.getItem("Data");
      let range = worksheet.getRange(address);
      range.load("values");
      await context.sync();

      let cellValue = range.values[0][0];
      // 定义需要检查的特定值
      let specificValues = ["Dimension", "Key", "SumY", "SumN", "Result"];

      // 去除可能的工作表名称前缀
      const cleanAddress = address.includes("!") ? address.split("!")[1] : address;

      // 正则表达式解释：
      // ^1:1$ 匹配完整的第一行
      // ^[A-Z]+1(:[A-Z]+1)?$ 匹配一个或多个列的第一行，如 A1, A1:A1, A1:B1
    //   const pattern = /^1:1$|^[A-Z]+1(:[A-Z]+1)?$/;
      const pattern = /^1:1$|^[A-Za-z]+1(:[A-Za-z]+1)?$/;
      console.log("isFirstRow address pattern.test is " + pattern.test(cleanAddress));
      console.log("isFirstRow address specificValues.includes is " + specificValues.includes(cellValue));
      let result = pattern.test(cleanAddress) && specificValues.includes(cellValue);
      console.log("isFirstRow result is " + result);
      return result;
  });
}




// -------------------------- 获取下拉菜单的值 -------参数为'dropdownContainer' 或 'dropdownContainer2'---------------------------
async function getSelectedOptions(containerId) {
  let selectedOptions = {};

  if (containerId === 'dropdown-container1') {
      // 
      selectedOptions = await getDropdownData("SelectedValue1");
  } else if (containerId === 'dropdown-container2') {
      // selectedOptions = selectedOptionsMapContainer2;
      selectedOptions = await getDropdownData("SelectedValue2");
  }

  console.log(selectedOptions);
  return selectedOptions;
}


//从SelectedValue1 和 SelectedValue2 工作表中获取数据
async function getDropdownData(sheetName) {
  try {
    // 获取当前Excel workbook的上下文
    return await Excel.run(async (context) => {
      // 获取指定工作表
      const sheet = context.workbook.worksheets.getItem(sheetName);

      // 获取工作表的UsedRange
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");

      await context.sync();

      // 获取UsedRange中的所有值
      const values = usedRange.values;

      // 确保第一行存在
      if (values.length === 0) {
        throw new Error("工作表为空或没有数据");
      }

      // 初始化结果对象
      const dropdownData = {};

      // 第一行为字段名
      const headers = values[0];

      // 遍历每个字段
      headers.forEach((header, columnIndex) => {
        if (header) { // 确保字段名不为空
          // 获取该列的值（从第二行开始）
          const columnValues = values.slice(1).map(row => row[columnIndex]).filter(value => value !== null && value !== undefined && value !== "");

          // 去重并存入对象
          dropdownData[header] = Array.from(new Set(columnValues));
        }
      });

      console.log(dropdownData);
      return dropdownData;
    });
  } catch (error) {
    console.error("Error retrieving dropdown data:", error);
    throw error;
  }
}


// ----------------------------------------- 根据下拉菜单的值 更新数据透视表 ---------------------------------------
async function updatePivotTableFromSelectedOptions(containerId,sheetName) {
    Excel.run(async (context) => {
        // 调用 getSelectedOptions 来获取选项
        //console.log("开始使用监听调用更新")
        
        const selectedOptions = await getSelectedOptions(containerId);  // 这应该是一个对象，键是字段名，值是选中的值数组
        
        // 遍历 selectedOptions 的每个键和值
        for (const [fieldName, fieldValues] of Object.entries(selectedOptions)) {
            // 调用 ControlPivotalTable 来更新数据透视表
            await ControlPivotalTable(sheetName, fieldName, fieldValues);
        }

        // 确保所有更改同步到工作簿
        await context.sync();
    }).catch(error => {
        console.error("Error:", error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info:", JSON.stringify(error.debugInfo));
        }
    });
}

// ------------------------- 使用这个操作数据透视表 --------------------------------
async function ControlPivotalTable(sheetName, fieldName, fieldValues) {
  Excel.run(async (context) => {
      // 获取名为"BasePT"的工作表上的名为"PivotTable"的数据透视表
      let pivotTable = context.workbook.worksheets.getItem(sheetName).pivotTables.getItem("PivotTable");
      
      // 根据传入的fieldName获取对应的层次和字段
      let fieldToFilter = pivotTable.hierarchies.getItem(fieldName).fields.getItem(fieldName);
      
      // 创建手动筛选对象，包含要筛选的值
      let manualFilter = { selectedItems: fieldValues };
      
      // 应用筛选
      fieldToFilter.applyFilter({ manualFilter: manualFilter });

      // 确保所有更改同步到工作簿
      await context.sync();
  }).catch(error => {
      console.error("Error:", error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info:", JSON.stringify(error.debugInfo));
      }
  });
}





  ////////////////////////////////////////////-----------------------------formula change --------------------------------------------------------------------///////////////////////////////////////
  const regLBraket=/^\($/
  const regRBraket=/^\)$/
  const regSignAdd=/^[+]$/
  const regSignSub=/^[-]$/
  const regSignMul=/^[*]$/
  const regSignDiv=/^[\/]$/
  const regEqual=/^\=$/
  const regComma=/^\,$/
  const regColon=/^\:$/
  
  const regArg=/^(\$?[a-z]+)(\$?[1-9][0-9]*)$/i
  const regNum=/^([0-9][0-9]*)(\.?[0-9]*)$/
  //const regNum=/^([0-9][0-9]*)(\.?[0-9]*)$|^(?<=[-+*\/])[-]([0-9][0-9]*)(\.?[0-9]*)$/
  const regSum=/^sum(?=\()/i
  const regFun=/^[a-z]+(?=\()/i
  
  const dtype={
      LB:111,
      RB:112,
      COMMA:270,
      COLON:260,
  
      SignMul:220,
      SignDiv:221,
      SignAdd:230,
      SignSub:231,
      SignEqual:250,
      
      VAR:301,
      NUM:302,
      FUNC:303,
          FSUM:304
      }
  
  const priority={
      LB:111,
      RB:112,
      COMMA:270,
      COLON:260,
  
      SignMul:220,
      SignDiv:220,
      SignAdd:230,
      SignSub:230,
      SignEqual:250,
      
      VAR:301,
      NUM:302,
      FUNC:303,
  
      EXP:404
      }
  
  function parseToken(strformula){
  
  if(!strformula)return;
  
  let result=[],tempStr="",len=strformula.length
  
  strformula=strformula.toUpperCase();
  
  for(let i=0;i<len;i++){
     tempStr=`${tempStr}${strformula[i]}`
  
     if(regLBraket.test(tempStr)){result.push({value:tempStr,type:dtype.LB,priority:priority.LB});tempStr="";continue;}
     if(regRBraket.test(tempStr)){result.push({value:tempStr,type:dtype.RB,priority:priority.RB});tempStr="";continue;}
     if(regComma.test(tempStr)){result.push({value:tempStr,type:dtype.COMMA,priority:priority.COMMA});tempStr="";continue;}  
     if(regColon.test(tempStr)){result.push({value:tempStr,type:dtype.COLON,priority:priority.COLON});tempStr="";continue;} 
  
     if(regSignMul.test(tempStr)){result.push({value:tempStr,type:dtype.SignMul,priority:priority.SignMul});tempStr="";continue;}   
     if(regSignDiv.test(tempStr)){result.push({value:tempStr,type:dtype.SignDiv,priority:priority.SignDiv});tempStr="";continue;} 
     if(regSignAdd.test(tempStr)){result.push({value:tempStr,type:dtype.SignAdd,priority:priority.SignAdd});tempStr="";continue;}
     if(regSignSub.test(tempStr)){result.push({value:tempStr,type:dtype.SignSub,priority:priority.SignSub});tempStr="";continue;}
     if(regEqual.test(tempStr)){result.push({value:tempStr,type:dtype.SignEqual,priority:priority.SignEqual});tempStr="";continue;} 
  
     if(regArg.test(tempStr)){
        if(i==len-1){	result.push({value:tempStr,type:dtype.VAR,priority:priority.VAR});tempStr="";continue;}
        if(!regArg.test(`${tempStr}${strformula[i+1]}`)){
          result.push({value:tempStr,type:dtype.VAR,priority:priority.VAR});tempStr="";continue;
          }
      }
  
     if(regNum.test(tempStr)){
        if(i==len-1){	result.push({value:tempStr,type:dtype.NUM,priority:priority.NUM});tempStr="";continue;}
        if(!regNum.test(`${tempStr}${strformula[i+1]}`)){
          result.push({value:tempStr,type:dtype.NUM,priority:priority.NUM});tempStr="";continue;
          }
      }
  
    if(i<len-1){
        if(regSum.test(`${tempStr}${strformula[i+1]}`)){
          result.push({value:tempStr,type:dtype.FSUM,priority:priority.FUNC});tempStr="";continue;
          }
     }
  
    if(i<len-1){
        if(regFun.test(`${tempStr}${strformula[i+1]}`)){
          result.push({value:tempStr,type:dtype.FUNC,priority:priority.FUNC});tempStr="";continue;
          }
     }	
      
      }
     
      return result;
  }
  
  function modifyToken(token,target){
      len=token.length;        
      token.forEach((v,i,arr)=>{
          if(v.value==":"&&i>0&&i<len-1){
                          let tarr=[];
              tarr=resovleColonAddr(arr[i-1].value,arr[i+1].value,target);			
              if(tarr){
                  arr[i-1]=tarr[0];
                  arr[i]=tarr[1];
                  arr[i+1]=tarr[2];
                  }
          }
  
          })
          return token;
  
  }
  
  
  function buildTree(eleArray,target){
      len=eleArray.length;tsign="";
      if(len<1){return;}
      if(!regArg.test(target))	{return;}
      
      let stackV=[],stackToken=[],sign;
      let TreeNode={},left,right,parent,position,type,op;
      let targetNode;
  
      let regTarget=new RegExp(target.replace(/^\$?([a-z]+)\$?([1-9][0-9]*)$/gi,"^\\$?$1\\$?$2$"),"ig");
  
      let sp=-1;
      for(let i=0;i<len;i++){
  
          switch(eleArray[i].type){
              case dtype.LB:
                  sign="LB";break;
              case dtype.RB:
                  sign="RB";break;
              case dtype.COMMA:
                  sign="SIGN";break;
              case dtype.COLON:
                  sign="SIGN";break;
  
              case dtype.NUM:
              case dtype.VAR:
                  sign="CONST";break;
  
              case dtype.SignAdd:
              case dtype.SignSub:
              case dtype.SignMul:
              case dtype.SignDiv:
              case dtype.SignEqual:
                  sign="SIGN";break;
  
              case dtype.FUNC:
                  sign="FUNC";break;
  
              case dtype.FSUM:
                  sign="FUNC";break;
              }
          stackToken.push(sign);
          stackV.push(eleArray[i]);
          
          sp++;
          if(sp<2){continue;}
          while(sp>=2){
  
              if((stackToken[sp-2]=="CONST"||stackToken[sp-2]=="EXP")&&(stackToken[sp-1]=="SIGN")&&(stackToken[sp]=="CONST"||stackToken[sp]=="EXP")){
                  if((i==len-1)||(eleArray[i+1].type>200&&eleArray[i+1].type<300&&eleArray[i+1].priority>=stackV[sp-1].priority)||eleArray[i+1].type<200||eleArray[i+1].type>300){
  
                      TreeNode={}					
  
                      left=stackToken[sp-2]=="CONST"?{pos:"left",value:stackV[sp-2].value,type:"leaf",parent:TreeNode}:stackV[sp-2]
                      left.pos="left";left.parent=TreeNode;
                      right=stackToken[sp]=="CONST"?{pos:"right",value:stackV[sp].value,type:"leaf",parent:TreeNode}:stackV[sp]
                      right.pos="right";right.parent=TreeNode;
                      
                      if(!targetNode&&stackToken[sp-2]=="CONST"&&regTarget.test(stackV[sp-2].value)){targetNode=left};
                      if(!targetNode&&stackToken[sp]=="CONST"&&regTarget.test(stackV[sp].value)){targetNode=right};
  
                      TreeNode.left=left;
                      TreeNode.right=right;
                      TreeNode.op=stackV[sp-1].value;
                      TreeNode.priority=stackV[sp-1].priority;
                      TreeNode.type="nonleaf"
                      TreeNode.pos="";
                      TreeNode.parent="";
                      
                      stackV.pop();
                      stackV.pop();
                      stackV[sp-2]=TreeNode;
                      
                      stackToken.pop();
                      stackToken.pop();
                      stackToken[sp-2]="EXP"
                      
                      sp=sp-2;
                      continue;
                  }
              }			
  
              if(stackToken[sp-2]=="LB"&&(stackToken[sp-1]=="CONST"||stackToken[sp-1]=="EXP")&&stackToken[sp]=="RB"){
  
                      stackV[sp-2]=stackV[sp-1];
                      stackV.pop();
                      stackV.pop();
  
                      stackToken[sp-2]=stackToken[sp-1];
                      stackToken.pop();
                      stackToken.pop();
                      
                      sp=sp-2;
                      continue;
              }
  
              if((stackToken[sp-2]=="SIGN"||stackToken[sp-2]=="LB")&&stackToken[sp-1]=="SIGN"&&(stackToken[sp]=="CONST"||stackToken[sp]=="EXP")){
                  tsign=stackV[sp-1].value;
                  if(tsign=="+"||tsign=="-") {
                      switch(stackV[sp-2].value){
                          case "+":;
                          case "-":;						
                          case "*":;
                          case "/":;
                          case "(":;
                          case ",":;
  
                          stackV[sp-1]=stackV[sp];
                          stackV.pop();
  
                          stackToken[sp-1]=stackToken[sp];
                          stackToken.pop();							
                          
                          sp--;
                      }
  
                       if(tsign=="-"){
                          
                          if(stackV[sp].type==dtype.NUM){
                                stackV[sp-1].value=="+"?stackV[sp-1].value="-":stackV[sp-1].value=="-"?stackV[sp-1].value="+":stackV[sp].value=-stackV[sp].value
                                                        continue;
                             }						  
                          if(stackV[sp-1].value=="+"||stackV[sp-1].value=="-"){
                            stackV[sp-1].value=stackV[sp-1].value=="+"?"-":"+";
                            continue;
                          }						
                            stackV.push({value:"*",type:dtype.SignMul,priority:priority.SignMul});
                            stackToken.push("SIGN");
                            stackToken.push("CONST");
                            stackV.push({value:"-1",type:dtype.NUM,priority:priority.NUM});
                            
                            sp=sp+2;
                            continue;
                      }
                                           continue;   
                          
                   }
  
  
  
  
              }
  
              if(stackToken[sp-1]=="FUNC"&&(stackToken[sp]=="CONST"||stackToken[sp]=="EXP")){
  
                      TreeNode={}					
  
                      left=stackToken[sp]=="CONST"?{pos:"left",value:stackV[sp].value,type:"leaf",parent:TreeNode}:stackV[sp]
                      left.pos="left";left.parent=TreeNode;
  
                      if(!targetNode&&stackToken[sp]=="CONST"&&regTarget.test(stackV[sp].value)){targetNode=left};
  
                      TreeNode.left=left;
                      TreeNode.right=undefined;
                      TreeNode.op=stackV[sp-1].value;
                      TreeNode.priority=stackV[sp-1].priority;
                      TreeNode.type="nonleaf"
                      TreeNode.pos="";
                      TreeNode.parent="";
  
                      stackV[sp-1]=TreeNode;
                      stackV.pop();
  
  
                      stackToken[sp-1]="EXP"
                      stackToken.pop();
                      
                      sp=sp-1;
                      continue;
              }
  
              break;
          }
  
      }
      
  return {TreeNode,targetNode};
  
  }
  
  function dbuildFormula(tn){
      let formula="";
  
      if(tn.type=="nonleaf"){
        if(tn.priority==priority.FUNC){
           return `${tn.op}(${dbuildFormula(tn.left)})`
        }
            else
           {
  
        formula=!tn.left.priority||tn.left.priority==priority.FUNC?`${dbuildFormula(tn.left)}`:tn.left.priority<=tn.priority?`${dbuildFormula(tn.left)}`:(tn.op==","?`${dbuildFormula(tn.left)}`:`(${dbuildFormula(tn.left)})`);
        formula+=`${tn.op}`;	
            formula+=!tn.right.priority||tn.right.priority==priority.FUNC?`${dbuildFormula(tn.right)}`:tn.right.priority<tn.priority?`${dbuildFormula(tn.right)}`:`(${dbuildFormula(tn.right)})` ;
        return formula;
      
        }
      }	
      else{
         return tn.value;
      }
  }
  
  function ubuildFormula(tn){
      let parent=tn.parent;
  
      if(!parent){return;}
      let formula="",op="",uformula="";
  
      if(tn.pos=="left"){
         switch(parent.op){
          case '+':op='-';break;
          case '-':op='+';break;
          case '*':op='/';break;
          case '/':op='*';break;
          case '=':op='=';break;
          default:op=parent.op;
          }
          parent.op=op;
          
          if(parent.op!="="){
              formula=`(${dbuildFormula(parent.right)})` ;
              uformula=`(${ubuildFormula(parent)})`;
              return `(${uformula}${parent.op}${formula})`;
          }else{
              return `(${dbuildFormula(parent.right)})`;	
          }
          
  
      }
      else{
         switch(parent.op){
          case '+':op='-';
              parent.op=op;
              formula=`(${dbuildFormula(parent.left)})` ;
              uformula=`(${ubuildFormula(parent)})`;
              return `${uformula}${parent.op}${formula}`;
              
          case '*':op='/';
              parent.op=op;
              formula=`(${dbuildFormula(parent.left)})` ;
              uformula=`(${ubuildFormula(parent)})`;
              return `${uformula}${parent.op}${formula}`;
              
          case '-':op='-';
              parent.op=op;
              formula=`(${dbuildFormula(parent.left)})` ;
              uformula=`(${ubuildFormula(parent)})`;
              return `${formula}${parent.op}${uformula}`
              
          case '/':op='/';
              parent.op=op;
              formula=`(${dbuildFormula(parent.left)})` ;
              uformula=`(${ubuildFormula(parent)})`;
              return `${formula}${parent.op}${uformula}`
              
          case '=':return `(${dbuildFormula(parent.left)})`;
          
          }
  
      }	
  
  }
  
  function resovleColonAddr(addr1,addr2,target){
  
    let iregTarget=new RegExp(target.replace(/^\$?([a-z]+)\$?([1-9][0-9]*)$/gi,"^\\$?$1\\$?$2$"),"ig");
    let uregTarget=new RegExp(target.replace(/^\$?([a-z]+)\$?([1-9][0-9]*)$/gi,"^\\$?\($1\)\\$?([1-9][0-9]*)$"),"ig");
  
    let item=[],r=[],c=[],ci=[],bitem=[];
    
    bitem[0]=addr1;
    bitem[1]=addr2;
  
    c[0]=target.replace(/[$0-9]/gi,"");
    r[0]=parseInt(target.replace(/[$a-z]/gi,""));
    c[1]=bitem[0].replace(/[$0-9]/gi,"");
    c[2]=bitem[1].replace(/[$0-9]/gi,"");
    r[1]=parseInt(bitem[0].replace(/[$a-z]/gi,""));
    r[2]=parseInt(bitem[1].replace(/[$a-z]/gi,""));
  
    ci[0]=colToNum(c[0]);
    ci[1]=colToNum(c[1]);
    ci[2]=colToNum(c[2]);
    
    if(!(ci[1]!=c[2]&&r[1]!=r[2])) return;  
     
    if(r[0]==r[1]&&r[0]==r[2]&&c[0]==c[1]&&c[0]==c[2]){
      item[0]={value: '0',type:dtype.NUM,priority:priority.NUM};
          item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
          item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
      return item;
      }
  
    if(r[1]==r[2]&&r[0]==r[1]){
      if(ci[0]==ci[1]){
          item[0]={value: `${bitem[0].replace(/([a-z][a-z]*)/gi,numToCol(ci[0]+1))}:${bitem[1]}`,type:dtype.VAR,priority:dtype.VAR};
              item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
              item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
          return item;
          }
          
          if(ci[0]==ci[2]){
          item[0]={value: `${bitem[0]}:${bitem[1].replace(/([a-z][a-z]*)/gi,numToCol(ci[0]-1))}`,type:dtype.VAR,priority:dtype.VAR};
              item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
              item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
          return item;
          }
  
          if(ci[0]>ci[1]&&ci[0]<ci[2]){
          item[0]={value: `${bitem[0]}:${bitem[0].replace(/([a-z][a-z]*)/gi,numToCol(ci[0]-1))},${bitem[1].replace(/([a-z][a-z]*)/gi,numToCol(ci[0]+1))}:${bitem[1]}`,type:dtype.VAR,priority:dtype.VAR};
              item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
              item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
          return item;
          }
  
     }
  
    if(uregTarget.test(bitem[0])){
      if(r[0]==r[1]){
          item[0]={value: `${bitem[0].replace(/([1-9][0-9]*)/gi,r[0]+1)}:${bitem[1]}`,type:dtype.VAR,priority:dtype.VAR};
              item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
              item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
          return item;
          }
          
          if(r[0]==r[2]){
          item[0]={value: `${bitem[0]}:${bitem[1].replace(/([1-9][0-9]*)/gi,r[0]-1)}`,type:dtype.VAR,priority:dtype.VAR};
              item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
              item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
          return item;
          }
  
          if(r[0]>r[1]&&r[0]<r[2]){
          item[0]={value: `${bitem[0]}:${bitem[0].replace(/([1-9][0-9]*)/gi,r[0]-1)},${bitem[1].replace(/([1-9][0-9]*)/gi,r[0]+1)}:${bitem[1]}`,type:dtype.VAR,priority:dtype.VAR};
              item[1]={value:",", type:dtype.COMMA, priority:priority.COMMA};
              item[2]={value: target, type:dtype.VAR, priority:dtype.VAR};
          return item;
          }
  
     }		
      return;
  
  }
  
  function colToNum(colName){
  
  let chars=[3];
  if(!colName) return 0;
  if(colName&&colName.length>3) return 0;
  chars=colName.toUpperCase().padStart(3,"$");
  return (chars[0]!="$"?chars[0].charCodeAt(0)-64:0)*676+(chars[1]!="$"?chars[1].charCodeAt(0)-64:0)*26+(chars[2]!="$"?chars[2].charCodeAt(0)-64:0);
  
  }
  
  function numToCol(colIndex){
  
  let chars=[3],i=0;
  
  if(colIndex<1||colIndex>16384) return;
  chars.forEach((v,i,arr)=>chars[i]="");
  
  if(colIndex>702){
     i=Math.floor((colIndex-703)/676);
     chars[0]=toColumnLetter(i);
     colIndex-=((i+1)*676);
  }
  if(colIndex>26){
     i=Math.floor((colIndex-27)/26);
     chars[1]=toColumnLetter(i);
     colIndex-=((i+1)*26);
  }
  i=colIndex-1
  chars[2]=toColumnLetter(i)
  return chars.join("");
  
  }
  
  
  function moveTreeNode(targetNode){
  
  let tn=targetNode;
  let commaNode=[],funcNode=[],funcSumCount=0,tempNode={};
  let cNode,pNode,broNode,ppNode,lrNodePointer;
  
  while(tn){
  
      if(tn.op&&tn.op==","&&(!commaNode[funcSumCount]))commaNode.push(tempNode);
      if(tn.op=="SUM"){
          if(!commaNode[funcSumCount])commaNode.push(undefined)
          funcNode.push(tn);
          funcSumCount++;
          }
  
      tempNode=tn;
          tn=tn.parent||undefined;
      if(!tn)break;
  }
  
  commaNode.reverse();
  funcNode.reverse();
  
  funcNode.forEach((v,i,arr)=>{
      cNode=commaNode[i];
      if(!cNode){
           v.left.parent=v.parent; 
           v.pos=="left"?v.parent.left=v.left:v.parent.right=v.left;		
           v.left.pos=v.pos;
           return;
         }
  
      pNode=cNode.parent;
      ppNode=pNode.parent;
      //lrNodePointer=ppNode.pos=="left"?ppNode.left:ppNode.right;
      broNode=cNode.pos=="left"?pNode.right:pNode.left;
      
      if(ppNode.op==","||broNode.op==","||ppNode.op==":"||broNode.op==":"||broNode.value.indexOf(":")>-1){
  
         pNode.pos=="left"?ppNode.left=broNode:ppNode.right=broNode;
         broNode.parent=ppNode;
         broNode.pos=pNode.pos;
  
         tempNode={}
         tempNode.pos=v.pos;
         tempNode.op="+";
          tempNode.type="nonleaf";
         tempNode.priority=priority.SignAdd;
  
         tempNode.parent=v.parent;
         tempNode.left=v;
         tempNode.right=cNode;	   
         tempNode.pos=="left"?tempNode.parent.left=tempNode:tempNode.parent.right=tempNode;
  
         v.parent=tempNode;
         cNode.parent=tempNode;
  
         v.pos="left";
         cNode.pos="right";
         
         return;	   
      }
         pNode.pos=v.pos;
         pNode.op="+";
         pNode.type="nonleaf";
         pNode.priority=priority.SignAdd;	   
         pNode.parent=v.parent;
         pNode.pos=="left"?v.parent.left=pNode:v.parent.right=pNode;		
         return;
      })
  
  }
  
//////////////////////////////--------------------------------- 解出方程 -----------------------------------------------/////////////////////

  function resolveEquation(formula,target){
  
  let count=0,revolvedFormula="";
  let iregTarget=new RegExp(target.replace(/^\$?([a-z]+)\$?([1-9][0-9]*)$/gi,"^\\$?$1\\$?$2$"),"ig");
  let tokens=parseToken(formula);
      tokens=modifyToken(tokens,target);
      tokens.forEach((v,i,arr)=>iregTarget.test(v.value)?(count++):count);
  
      target=target.toUpperCase();
  
  let {TreeNode,targetNode}=buildTree(tokens,target);  
      moveTreeNode(targetNode);
  
      revolvedFormula=`${target}=${ubuildFormula(targetNode)}`;
      //console.log(revolvedFormula);
      tokens=parseToken(revolvedFormula);
      TreeNode=buildTree(tokens,target)["TreeNode"];
      //console.log(TreeNode);
      return dbuildFormula(TreeNode);
  }
  
//   let formula='d1=B1+sum(-sum(D1:D130,D1260,9*-20,+++++++++-------------------A10,round(20))*(X1-Y1)+-5+k(+30),sum(A10:A20)/30,C95:C100)*100+-E140';
//   let target="$D126"
  
//    console.log(`formula: ${formula}`);
//    console.log(resolveEquation(formula,target));


  async function GetFormulas() {
    await Excel.run(async (context) => {
      
        const workbook = context.workbook;
        const sheet = workbook.worksheets.getItem("formulas");
        const ResultRange = sheet.getRange("H2");
        ResultRange.load("values, text, address, rowCount, columnCount, formulas")
    
        await context.sync();
        let formula = ResultRange.address.split("!")[1] + ResultRange.formulas[0][0];
        let target = "C2"
    
        console.log(formula);
        console.log(resolveEquation(formula,target));
        
  
    });

    
  }

//---------------------从 Bridge Data 建立数据透视表 生成 Base / Target --------------------
async function createPivotTableFromBridgeData(NewSheetName) {
  console.log("step00000000")
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const bridgeDataSheet = workbook.worksheets.getItem("Bridge Data");
    console.log("step11111111")
    
    // 检查是否存在同名的工作表
    let basePTSheet = workbook.worksheets.getItemOrNullObject(NewSheetName);
    await context.sync();

    if (basePTSheet.isNullObject) {
      // 工作表不存在，创建新工作表
      basePTSheet = workbook.worksheets.add(NewSheetName);
      await context.sync();
      console.log("创建了新工作表：" + NewSheetName);
    } else {
      console.log("工作表已存在：" + NewSheetName);
    }

    console.log("Here");

    const fullUsedRange = bridgeDataSheet.getUsedRange();
    fullUsedRange.load("address");  // 加载范围的地址属性
    await context.sync();

    // 修改范围地址以从B列开始
    const newRangeAddress = fullUsedRange.address.replace(/^([^!]+)!A/, '$1!B');

    // 获取从B列开始的使用范围
    const usedRange = bridgeDataSheet.getRange(newRangeAddress);
    usedRange.load("address");
    usedRange.load("rowCount");
    await context.sync();

    console.log("The address of the usedRange is: " + usedRange.address)
    if (usedRange.rowCount < 2) {
      console.error("Not enough rows in used range to perform operation.");
      return;
    }

    // 读取第一行以确定字段应放在数据透视表的哪个部分
    const configRange = usedRange.getRow(0);
    configRange.load("values");
    await context.sync();

    // 读取第二行作为字段名
    const headerRange = usedRange.getRow(1);
    headerRange.load("values");
    await context.sync();

    const rangeAddress = fullUsedRange.address;
    const sheetName = rangeAddress.split('!')[0];
    const columnRow = rangeAddress.split('!')[1];
    const columns = columnRow.split(':')[1]; // 提取结束列信息
    const newRangeAddress2 = `${sheetName}!B2:${columns}`; // 设置从第三行开始的新范围

    const dataRange = bridgeDataSheet.getRange(newRangeAddress2);
    dataRange.load("address");
    await context.sync();

    console.log("The address of the range is: " + dataRange.address)

    // 激活工作表
    basePTSheet.activate();
    await context.sync();

    // 检查是否存在同名的数据透视表
    let pivotTable = basePTSheet.pivotTables.getItemOrNullObject("PivotTable");
    await context.sync();

    if (!pivotTable.isNullObject) {
      // 数据透视表已存在，删除原有的数据透视表
      pivotTable.delete();
      await context.sync();
      console.log("已删除原有的数据透视表 'PivotTable'。");
    }

    // 创建新的数据透视表
    pivotTable = basePTSheet.pivotTables.add("PivotTable", dataRange, "C3");
    pivotTable.refresh(); // 必须加 refresh，不然改了标题名字就不能刷新了

    await context.sync();
    console.log("step6")

    // 配置数据透视表字段
    const configValues = configRange.values[0];
    console.log("configValues is " + configValues)
    const headerValues = headerRange.values[0];
    for (let i = 0; i < headerValues.length; i++) {
        const fieldName = headerValues[i];

        const columnIndex = i + 1; // B列开始，索引偏移1
        const columnLetter = toColumnLetter(columnIndex); // ASCII for 'A' is 65
        const fullColumnName = `${columnLetter}:${columnLetter}`;

        switch (configValues[i]) {
          case "Key":
            pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
            break;
          case "Dimension":
            pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
            break;
          case "SumY":
          case "SumN":
          case "Result":
          case "ProcessSum":
            if(ArrVarPartsForPivotTable.includes(fieldName)){
              console.log("SumY is " + fieldName);
              const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
              dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
              dataHierarchy.name = `Sum of ${fieldName}`; // 将字段名改成英文的 "Sum of"
              break;
            }
        }
    }

    pivotTable.layout.layoutType = "Tabular"; // 设置数据透视表的展现格式
    pivotTable.layout.subtotalLocation = Excel.SubtotalLocationType.off;
    pivotTable.layout.showRowGrandTotals = false;
    pivotTable.layout.showColumnGrandTotals = true;
    pivotTable.layout.repeatAllItemLabels(true);
    console.log("step7")
    basePTSheet.activate();
    await context.sync();
    console.log("Data Pivot Table created successfully on '" + NewSheetName + "' sheet.");

    await CreateLabelRange(NewSheetName); // 在数据透视表下面加一行不带 Sum of 的标题
    await DoNotChangeCellWarning(NewSheetName); // 添加禁止修改的警告长方形

  }).catch(error => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}


//--------------监控Bridge Data的变化，实时生成新的combine数据透视表--------------------
// async function createCombinePivotTable() {
//     try {
//         await Excel.run(async (context) => {
//             const sheets = context.workbook.worksheets;
//             sheets.load("items/name");

//             await context.sync();

//             const sheetName = "Combine";
//             let sheet = sheets.items.find((worksheet) => worksheet.name === sheetName);

//             if (sheet) {
//                 sheet.delete();
//                 await context.sync();
//                 console.log(`Sheet "${sheetName}" has been deleted.`);
//             } else {
//                 console.log(`Sheet "${sheetName}" does not exist.`);
//             }

//             // 调用 createPivotTableFromBridgeData 函数
//             await createPivotTableFromBridgeData("Combine");


//             await context.sync();
//         });
//     } catch (error) {
//         console.error(error);
//     }
// }




  // -------------------获取数据透视表的数据部分-------------已测试----------------
async function GetPivotRange(SourceSheetName) {
    let RangeInfo = null;
    await Excel.run(async (context) => {
      
      let pivotTable = context.workbook.worksheets.getItem(SourceSheetName).pivotTables.getItem("PivotTable");
      console.log("GetPivotFunc")
      // 获取不同部分的范围
      let DataRange = pivotTable.layout.getDataBodyRange();
      let RowRange = pivotTable.layout.getRowLabelRange();
      let PivotRange = pivotTable.layout.getRange();
      let ColumnRange = pivotTable.layout.getColumnLabelRange();

      //let LabelRange = DataRange.getLastRow().getOffsetRange(1,0); // 在dataRange的最后一行的下一行
      //LabelRange.copyFrom(ColumnRange,Excel.RangeCopyType.values);
      
      console.log("GetPivotFunc 1")
      DataRange.load("address");
      RowRange.load("address");
      PivotRange.load("address");
      ColumnRange.load("address");
      //LabelRange.load("address");

      await context.sync();
      console.log("GetPivotFunc 2")
      // 加载它们的地址属性
      console.log(DataRange.address)
      console.log(RowRange.address)
      console.log(PivotRange.address)
      console.log(ColumnRange.address)
      //console.log("Label Range is " + LabelRange.address)
      //await CleanHeader(SourceSheetName,LabelRange.address); //需要传递LabelRange.address 而不是LabelRange


      await context.sync(); // 同步更改
      //return PivotRange.address
      //   返回这些地址
        RangeInfo= {
          dataRangeAddress: DataRange.address,
          rowRangeAddress: RowRange.address,
          pivotRangeAddress: PivotRange.address,
          columnRangeAddress: ColumnRange.address
      };
  
  
    });

    return RangeInfo;

  }


  // 创建Process 数据表，拷贝Combine数据, 并清空数据，保留Key 和 格式
async function CreateAnalysisSheet(SourceSheetName, TargetSheetName) {
    await Excel.run(async (context) => {
      const workbook = context.workbook; // 获取工作簿引用
      const analysisSheet = workbook.worksheets.add(TargetSheetName); // 添加新的工作表
      await context.sync()
  
      const pivotRanges = await GetPivotRange(SourceSheetName); // 确保异步获取完成
      let SourceRange = pivotRanges.pivotRangeAddress; // 整个pivotTable 的 Range
      console.log(SourceRange);
      const startRange = analysisSheet.getRange("B3");
      await context.sync()
      
      // 由于GetPivotRange返回的是包含地址的对象，需要在工作簿上使用这些地址
      //const dataRange = workbook.getRange(pivotRanges.pivotRangeAddress);
  
      startRange.copyFrom(SourceRange); // 使用copyFrom复制

      await context.sync(); // 同步更改
      let processRange = null;
      //如果是Process工作表，则传递新的Range给全局变量StrGlobalProcessRange
      if (TargetSheetName == "Process" ){
        console.log(" in if")
        let tempRange = context.workbook.worksheets.getItem(SourceSheetName).getRange(SourceRange);
        tempRange.load("address,columnCount,rowCount");
        await context.sync();

        //console.log(tempRange.rowCount)
        //console.log(tempRange.columnCount)
        let processRange = startRange.getAbsoluteResizedRange(tempRange.rowCount,tempRange.columnCount); //重新获取copy来的Range
        let firstRow = processRange.getRow(0);
        processRange.load("address");
        firstRow.load("address");
        await context.sync();
        //console.log(processRange.address)
        StrGlobalProcessRange = processRange.address;  // 传递给全局变量
        CleanHeader(TargetSheetName, firstRow.address); // 清除Sum of

        let dataStartRange = startRange.getOffsetRange(1,1); // ProcessRange 保留标题的起始地址
        let dataRange = dataStartRange.getAbsoluteResizedRange(tempRange.rowCount-1, tempRange.columnCount-1); // ProcessRange的dataRange
        dataRange.clear(Excel.ClearApplyTo.contents); // 只清除数据，保留格式

        //console.log("Global Range is" + TargetSheetName + StrGlobalProcessRange)

        // let nextProcessRange = processRange.getOffsetRange(0, tempRange.columnCount+1); // ProcessRange 平移
        // nextProcessRange.load("address, values");

        // await context.sync();

        // startRange.getOffsetRange(-2,0).values = [[nextProcessRange.address]];
        await DoNotChangeCellWarning(TargetSheetName);
        await context.sync();

      }

    });
  }
  

  //---------------------- 删除 sum of---------已测试------------------
  async function CleanHeader(SheetName, Range) {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getItem(SheetName);
      const HeaderRange = sheet.getRange(Range);
      HeaderRange.load("values, text, address, rowCount,columnCount");
  
      await context.sync();
  
      let ReplaceCriteria = {
        completeMatch: false,
        matchCase: false
      };
  
      HeaderRange.replaceAll("Sum of ", "", ReplaceCriteria);
      await context.sync();
    });
  }

  // -----------------获得Occ%=Room Revenue/ARR/Ava. Rooms 之中的每个变量对应标题的下一行的单元格地址,并赋值到新目标单元格-------已测试--------------
async function GetFormulasAddress(sourceSht, sourceRng, targetSht, targetRng) {
    console.log("objGlobalFormulasAddress is ");
    console.log(objGlobalFormulasAddress);
    return await Excel.run(async (context) => {
      const sourceSheet = context.workbook.worksheets.getItem(sourceSht);
      const sourceRange = sourceSheet.getRange(sourceRng);
      const targetSheet = context.workbook.worksheets.getItem(targetSht);
      const targetRange = targetSheet.getRange(targetRng);
      sourceRange.load("values, address");
  
      await context.sync();
      console.log("step1111");
      console.log("sourceRange.values is " + sourceRange.values[0][0]);
      console.log("sourceRange.address is " + sourceRange.address);


      const formula = sourceRange.values[0][0];
      //console.log(formulas);
      if (typeof formula === "string" && formula.includes("=")) {
        const parts = formula.split("=");
        const formulaName = parts[0].trim(); // 获取公式的名称并去除两端的空白
        const formulaContent =
          "=" +
          parts
            .slice(1)
            .join("=")
            .trim(); // 获取公式的内容，并确保等号和内容
        console.log("formulaContent is: ");
        console.log(formulaContent);
  
        await CleanHeader(targetSht, targetRng); //清除sum of, 必须要加await~!!!
        await context.sync();
        targetRange.load("values, address"); //这里要清除以后再load, 提前load 没有效果
        await context.sync(); //// 任何操作excel的都需要同步~！！！
  
        const values = targetRange.values[0];
        console.log("targetRange is: " + targetRange.address);
        const updatedFormulasAddress = {};
        console.log("step444");
        let CellTitles = objGlobalFormulasAddress;
        console.log("CellTitles is");
        console.log(CellTitles);
  
        // 加载并同步 targetRange 的起始行号
        const firstCell = targetRange.getCell(0, 0);
        firstCell.load("rowIndex"); // 所有的属性都需要加载~！！！
        await context.sync();
        const targetRangeStartRow = firstCell.rowIndex + 1;
  

        // 对比target Range 中新的title，获取公式中对应的新的对象，包含单元格地址
        for (const [key, originalAddress] of Object.entries(CellTitles)) {
          console.log("boject in");
          for (let colIndex = 0; colIndex < values.length; colIndex++) {
            console.log(key + "=" + values[colIndex]);
            if (values[colIndex] === key) {
              //const columnLetter = String.fromCharCode(65 + colIndex + 2); // colindex 从 0 开始，对应A列, //// 这里标题从C列开始，因此要+2, 这里需要做灵活变化~！！！
              let targetColumn = targetRange.getColumn(colIndex); // 直接从targetRange 中寻找列
              targetColumn.load("address");
              await context.sync();
              console.log("targetColumn is "+ targetColumn);
              
              //let columnLetter = targetColumn.address.split("!")[1][0];
              let columnLetter = getRangeDetails(targetColumn.address).leftColumn
              console.log("columnLetter is " + columnLetter);

              console.log("column");
              const newRow = targetRangeStartRow + 1; // 获取下一行的单元格地址
              console.log("Row");
              const newAddress = `${columnLetter}${newRow}`;
              console.log("Address");
              updatedFormulasAddress[key] = newAddress; // 是一个对象
            }
          }
        }
        console.log("updatedFormulasAddress is:  ");
        console.log(updatedFormulasAddress);
  
        // 获取对象的属性数组
        const entries = Object.entries(updatedFormulasAddress);
  
        // 按键的长度进行排序
        entries.sort((a, b) => b[0].length - a[0].length);
  
        // 构造一个新的排序后的对象
        const RankedFormulasAddress = {};
        for (const [key, value] of entries) {
          RankedFormulasAddress[key] = value;
        }
  
        console.log("RankedFormulasAddress is: ");
        console.log(RankedFormulasAddress);
  
        let newFormulaContent = formulaContent; // 准备将变量名替换成变量地址
        let targetVarAddress = null;
        console.log("Before newFormulaContent is");
        console.log(newFormulaContent);
  
        for (let key in RankedFormulasAddress) {
          if (RankedFormulasAddress.hasOwnProperty(key)) {
            let value = RankedFormulasAddress[key];
            let formattedValue = `{_${value}_}`; // 为 value 添加前后的字符串
            let regex = new RegExp(escapeRegExp(key), 'g'); // 创建一个全局匹配的正则表达式,需要escapeRegExp函数对 key 中的特殊字符进行转义，这样它们在正则表达式中将被视为普通字符
            console.log("key is :" + key);
            console.log("formattedValue is ");
            console.log(formattedValue);
            newFormulaContent = newFormulaContent.replace(regex, formattedValue); // 替换匹配的字符串
            console.log("In Loop newFormulaContent is:" + newFormulaContent);
          }
        }
        newFormulaContent = newFormulaContent.replace(/{_|_}/g, '').replace("=",""); // 把前面的等号去掉，下面加上=IFERROR
  
        console.log(" newFormulaContent is:" + newFormulaContent);
  
        let targetVar = Object.keys(CellTitles)[0]; // 要求的变量存在第一个属性
        console.log("targetVar is");
        console.log(JSON.stringify(targetVar,null,2));
        //找到求解变量需要对应的单元格
        const foundRange = targetRange.find(targetVar, {
          completeMatch: true,
          matchCase: true,
          searchDirection: "Forward"
        });
        // 往下一行，放公式
        const nextRowRange = foundRange.getOffsetRange(1, 0);
        console.log("GetFormulasAddress 2");
        nextRowRange.formulas = [[`=IFERROR(${newFormulaContent},0)`]]; // 加入IFERROR(),避免出现除于0等情况
        nextRowRange.load("address");
        await context.sync();
  
        StrGlbProcessSolveStartRange = nextRowRange.address // 将第一个带有求解公式的地址赋值给全局变量
        console.log("StrGlbProcessSolveStartRange is " + StrGlbProcessSolveStartRange);
        //return updatedFormulasAddress;
        //console.log("Formula Name:", formulaName);
        //console.log("Formula Content:", formulaContent);
  
        //return { formulaName, formulaContent };
      } else {
        console.error("The cell does not contain a valid formula.");
        return null;
      }
    });
  }


  
// --------------------获取单元格的公式，并形成对象------已测试----目前已经将求解后的公式放在了需要求解变量的单元格如 ADR, OCC%-------------
async function getFormulaCellTitles(sheetName, formulaAddress) {
    console.log("getFormulaCellTitles run")
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const formulaCell = sheet.getRange(formulaAddress);
      formulaCell.load("formulas, values, address");
      await context.sync();
      console.log("formulacell is " + formulaCell.address)
      //console.log("formulaCell.values is " + formulaCell.values[0][0])
      const cellValue = formulaCell.values[0][0];
  
      if (typeof cellValue !== "string") {
        console.error("The cell value is not a string or is empty????.");
  
        return {};
      }
  
      const formula = formulaCell.values[0][0].replace(/\$/g, ""); // 
  
      //console.log(formula);
      //const cellReferenceRegex = /([A-Z]+[0-9]+)/g;
        const cellReferenceRegex = /([A-Za-z]+\d+)/g;
      const cellReferences = formula.match(cellReferenceRegex);
  
      if (!cellReferences) {
        console.log("No cell references found in the formula.");
        return {};
      }
  
      const cellTitles = {}; // 创建一个对象
  
      for (const cellReference of cellReferences) {
        // const match = cellReference.match(/([A-Z]+)([0-9]+)/);
          const match = cellReference.match(/([A-Za-z]+)(\d+)/);
        if (match) {
          const column = match[1];
          const row = parseInt(match[2]);
          const titleCellAddress = `${column}${row - 1}`;
          const titleCell = sheet.getRange(titleCellAddress);
          titleCell.load("values");
          await context.sync();
          const title = titleCell.values[0][0];
          cellTitles[title] = cellReference;
        }
      }
      //console.log("getFormulaCellTitles end")
      console.log(cellTitles);
      return cellTitles;
    });
  }

  //// ----------------------------------将反算公式的title 输入表格---------------已测试---------------
async function replaceCellAddressesWithTitles(sheetName, formulaCellAddress, targetCellAddress, cellTitles) {
    //console.log("replaceCellAddressesWithTitles run")
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
  
      // 获取 cellTitles
      //const cellTitles = await getFormulaCellTitles(sheetName, formulaCellAddress);
      //console.log(cellTitles);
      // 获取目标单元格中的公式
      const targetCell = sheet.getRange(targetCellAddress);
      const sourceCell = sheet.getRange(formulaCellAddress);
      sourceCell.load("formulas");
      targetCell.load("formulas");
      await context.sync();
      let formula = sourceCell.formulas[0][0];
      //console.log("test"+ formula)
      // 替换公式中的单元格地址为对应的标题
      for (const title in cellTitles) {
        const cellAddress = cellTitles[title];
        const cellAddressRegex = new RegExp(cellAddress, "g");
        formula = formula.replace(cellAddressRegex, title);
      }
  
      // 将新的公式设置回目标单元格
      targetCell.values = [[`${formula}`]]; // 需要一个二维数组
      //console.log(formula)
      await context.sync();
  
      //console.log(`Updated formula in ${targetCellAddress}: ${formula}`);
    });
    //console.log("replaceCellAddressesWithTitles end")
  }

  //----------------------复制bridge data 作为temp-------已测试-------//
async function copyAndModifySheet(SourceSheet,TargetSheet) {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sourceSheetName = SourceSheet;
      const targetSheetName = TargetSheet;
  
      // Get the source sheet
      const sourceSheet = workbook.worksheets.getItem(sourceSheetName);
  
      // Copy the source sheet
      const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.after, sourceSheet);
      copiedSheet.name = targetSheetName;
  
      await context.sync();
  
      // Load the used range to determine the number of rows
      const usedRange = copiedSheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();
  
      // Determine the number of rows to delete
      const rowCount = usedRange.rowCount;
      if (rowCount > 3) {
        const rowsToDelete = copiedSheet.getRange(`4:${rowCount}`);
        rowsToDelete.delete(Excel.DeleteShiftDirection.up);
      }
    
        await DoNotChangeCellWarning(TargetSheet);
      await context.sync();
      //console.log(`Sheet '${targetSheetName}' created and modified successfully.`);
    });
  }

  //------------获取Bridge Data Temp 中 ProcessSum的地址，返回一个Cell的地址----------已测试-------------//
async function findResultCell(ProcessSumAddress) {
    console.log("findResultCell run")
    return await Excel.run(async (context) => {
      const sheetName = "Bridge Data Temp";
      // const searchKeyword = Keyword; // 搜索关键词
      // console.log("searchKeyword is"+ searchKeyword)
      let sheet = context.workbook.worksheets.getItem(sheetName);
      let ProcessSumCell = sheet.getRange(ProcessSumAddress);
      let SecondRowCell = ProcessSumCell.getOffsetRange(1,0);
      let ThirdRowCell = ProcessSumCell.getOffsetRange(2, 0);
      SecondRowCell.load("values");
      ThirdRowCell.load("address,formulas");
      await context.sync();

      let secondRowTitle = SecondRowCell.values[0][0];
      let thirdRowAddress = ThirdRowCell.address;
      let thirdRowFormula = ThirdRowCell.formulas[0][0];
      let resultDetails = [];
      resultDetails.push([secondRowTitle, thirdRowAddress, thirdRowFormula]);



      // // 获取工作表的使用范围
      // let usedRange = sheet.getRange(StrGblProcessSumCell).getAbsoluteResizedRange(3,1); //用了loop以后只拿到最高的单元格，因此必须要往下扩大
      // usedRange.load("address,values,formulas");
      // await context.sync();
      // console.log("usedRange is " + usedRange.address);
      // // 获取使用范围的第一行和第二行
      // // let firstRowRange = usedRange.getRow(0);
      // // let secondRowRange = usedRange.getRow(1);
      // // firstRowRange.load("values");
      // // secondRowRange.load("values");
      // // await context.sync();
  
      // const firstRowValues = usedRange.values[0];
      // const secondRowValues = usedRange.values[1];
      // let resultDetails = [];
      
      // // 搜索包含 "Keyword" 的单元格
      // for (let col = 0; col < firstRowValues.length; col++) {
      //   if (firstRowValues[col] === searchKeyword) {
      //       console.log("firstRowValues[col] is" + firstRowValues[col])
      //     // 获取第二行的标题
      //     let secondRowTitle = secondRowValues[col];
      //     // 获取第三行中对应列的单元格
      //     let thirdRowCell = usedRange.getCell(2, col); // Row index is 2 for third row
      //     thirdRowCell.load("address");
      //     thirdRowCell.load("formulas");
      //     await context.sync();

      //     console.log("thirdRowFormula1 is " + thirdRowCell.formulas[0][0]);
      //     thirdRowCell.formulas = [[thirdRowCell.formulas[0][0].replace(/\$/g, "")]]
      //     //await context.sync(); // 确保将修改同步到Excel
      //     thirdRowCell.load("formulas"); 

      //     await context.sync();

      //     console.log("thirdRowFormula3 is " + thirdRowCell.formulas[0][0]);

      //     let thirdRowAddress = thirdRowCell.address;
      //     let thirdRowFormula = thirdRowCell.formulas[0][0];
      //       console.log("thirdRowAddress is " + thirdRowAddress);
      //       console.log("thirdRowFormula2 is " + thirdRowFormula);
      //     // 将结果添加到数组中
      //     resultDetails.push([secondRowTitle, thirdRowAddress, thirdRowFormula]);
      //   }
      // }
  
      // if (resultDetails.length > 0) {
      //   //console.log("Found results:", resultDetails);
      // } else {
      //   console.log(`"${searchKeyword}" not found in the first row.`);
      // }
      // //console.log("findResultCell end")
      return resultDetails;
    });
  }

  //-------------------------- 找到在Result 公式中的 要解的变量单元格------已测试------------//
async function processResultFormulas(ProcessSumCell) {
    console.log("processResultFormulas run")
    const resultDetails = await findResultCell(ProcessSumCell);
    console.log(resultDetails)
    if (resultDetails.length === 0) {
      console.log("No results found.");
      return [];
    }
  
    //console.log("process:  " + resultDetails);
    return await Excel.run(async (context) => {
      const sheetName = "Bridge Data Temp";
      const sheet = context.workbook.worksheets.getItem(sheetName);
  
      let nonAdditiveAddresses = [];
  
      for (let [secondRowTitle, thirdRowAddress, thirdRowFormula] of resultDetails) {
        // let cellReferences = thirdRowFormula.match(/([A-Z]+[0-9]+)/g); // match 返回的是一个数组
          let cellReferences = thirdRowFormula.match(/([A-Za-z]+\d+)/g); // match 返回的是一个数组, 修改成可以匹配大小写
        //cellReferences = cellReferences.replace(/\$/g, ""); 不能直接将数组中的$替换
        
        // 将公式中的$固定符号替换
        if (cellReferences) {
          cellReferences = cellReferences.map(reference => reference.replace(/\$/g, ""));
        }
        if (!cellReferences) continue;
  
        for (let cellReference of cellReferences) {
        //   const match = cellReference.match(/([A-Z]+)([0-9]+)/);  // 解析result 中的公式
            const match = cellReference.match(/([A-Za-z]+)(\d+)/);  // 解析result 中的公式
          if (match) {
            const column = match[1];
            const row = parseInt(match[2]);
  
            if (row > 1) {
              const firstRowCell = sheet.getRange(`${column}1`);
              firstRowCell.load("values, address");
              await context.sync();
  
              if (firstRowCell.values[0][0] === "SumN") {   //根据第一行的标识找出要解的变量的地址
                nonAdditiveAddresses.push(cellReference);
              }
            }
          }
        }
      }
  
      //console.log("SumN addresses:", nonAdditiveAddresses);
      //console.log("processResultFormulas end")
      return [nonAdditiveAddresses, resultDetails];
    });
  }

//-------------------- 将Bridge Data Temp 整个单元格复制成值-------已测试---------------------
async function pasteSheetAsValues(SheetName) {
    //console.log("pasteSheetAsValues run")
    await Excel.run(async (context) => {
      const sheetName = SheetName; // 请根据需要修改工作表名称
      const sheet = context.workbook.worksheets.getItem(sheetName);
  
      // 获取工作表的使用范围
      const usedRange = sheet.getUsedRange();
      usedRange.load("address");
      await context.sync();
  
      // 复制使用范围并粘贴为值
      usedRange.copyFrom(usedRange, Excel.RangeCopyType.values);
  
      await context.sync();
  
      //console.log(`All cells in '${sheetName}' have been pasted as values.`);
    });
    //console.log("pasteSheetAsValues end")
  }


  ///-------------执行逆运算，根据Result 和 target的个数需要进行调整--------------------/////
async function runProcess(ProcessSumCell) {
    console.log("runProcess start");
  // const resultDetails = await findResultCell(ProcessSumCell);
  //   console.log(resultDetails);
  //   console.log("runProcess Step 1")
  //   if (resultDetails.length === 0) {
  //     console.log("No results found.");
  //     return [];
  //   }
  let [nonAdditiveAddresses, resultDetails] = await processResultFormulas(ProcessSumCell);
  // const nonAdditiveAddresses = await processResultFormulas(ProcessSumCell); //
  
    console.log("Target is  " + nonAdditiveAddresses)
    if (nonAdditiveAddresses.length === 0) {
      strGlobalFormulasCell = null;  //如果没有找到任何的SumN，则这里要设为空，不然会保留上一次求解得到的地址
      console.log("No non-additive addresses found.");

      return [];
    }
    console.log("runProcess Step 2")
    let results = [];
    let targets = [];
  
    //下面的循环只对应一个方程，如果有多个方程需要进一步调整目标单元格
    for (let [, thirdRowAddress, thirdRowFormula] of resultDetails) {
      //console.log(thirdRowAddress.split("!")[1] + thirdRowFormula, nonAdditiveAddresses[0])
      let result = resolveEquation(thirdRowAddress.split("!")[1] + thirdRowFormula, nonAdditiveAddresses[0]); // ***这里若有几个 target 需要求解，则需要利用循环等修改。nonAdditiveAddresses[0] 只求解一个ProcessSum里面第一个SumN
      console.log("result is " + result);
      //result = '=' + result.split('=')[1]; // 只保留公式部分
      results.push(result);
    }
    console.log(" runProcess Step 3")
    //console.log("Resolved equations results:", results);
    //console.log(nonAdditiveAddresses[0])
  
    return await Excel.run(async (context) => {
      console.log("runProcess Step 4")
      const sheet = context.workbook.worksheets.getItem("Bridge Data Temp");
      let targetRange = sheet.getRange(nonAdditiveAddresses[0]).getOffsetRange(1,0);//往下一行，不要覆盖原来的数据
      targetRange.load("address");
      await context.sync();
      console.log("runProcess Step 5")
      //await pasteSheetAsValues(); // 粘贴成值
      //const formulasArray = results.map(result => [result]); // 将一维数组转换为二维数组, 但目前只对一个单元格暂时不需要
      targetRange.values = [[results[0]]]; // 只使用第一个结果, 将解出后的公式放入目标单元格
      //console.log("end")
      //return results;
      console.log("runProcess Step 6")
      await context.sync(); ////////少了这一步，导致 targetRange.values = [[results[0]]]; 没有及时同步，后面的出错/////////////////////
  
      var cellTitles = await getFormulaCellTitles("Bridge Data Temp", targetRange.address);
      objGlobalFormulasAddress = cellTitles;
      // console.log("cellTitles in runprocess is ")
      // console.log(cellTitles)
      // console.log("objGlobalFormulasAddress in runprocess is ")
      // console.log(globalFormulasAddress)
      
      await context.sync();
      await replaceCellAddressesWithTitles(
        "Bridge Data Temp",
        targetRange.address,
        targetRange.address,
        cellTitles
      );
      console.log("runProcess Step 7")
      strGlobalFormulasCell = targetRange.address; // 处理结束后把保留变量名公式的地址传递给全局变量，以便使用。
      targetRange.load("address,values");
      await context.sync();
      console.log("test process");
      console.log("run process:  " + targetRange.values[0][0]);
    });
  }

// 创建数据透视表下一行不带Sum of 的标题列
  async function CreateLabelRange(SourceSheetName) {
    let RangeInfo = null;
    await Excel.run(async (context) => {
      
      let pivotTable = context.workbook.worksheets.getItem(SourceSheetName).pivotTables.getItem("PivotTable");
      console.log("GetPivotFunc")
      // 获取不同部分的范围
      let DataRange = pivotTable.layout.getDataBodyRange();
      let RowRange = pivotTable.layout.getRowLabelRange();
      let PivotRange = pivotTable.layout.getRange();
      let ColumnRange = pivotTable.layout.getColumnLabelRange();

      let LabelRange = DataRange.getLastRow().getOffsetRange(1,0); // 在dataRange的最后一行的下一行
      LabelRange.copyFrom(ColumnRange,Excel.RangeCopyType.values);
      
      console.log("GetPivotFunc 1")
      DataRange.load("address");
      RowRange.load("address");
      PivotRange.load("address");
      ColumnRange.load("address");
      LabelRange.load("address");

      await context.sync();
      console.log("GetPivotFunc 2")
      // 加载它们的地址属性
      console.log(DataRange.address)
      console.log(RowRange.address)
      console.log(PivotRange.address)
      console.log(ColumnRange.address)
      console.log("Label Range is " + LabelRange.address)
      await CleanHeader(SourceSheetName,LabelRange.address); //需要传递LabelRange.address 而不是LabelRange
      let strGlobalLabelRange = LabelRange.address; // 给全局变量赋值

      await context.sync(); // 同步更改
      //return PivotRange.address
      //   返回这些地址
        RangeInfo= {
          dataRangeAddress: DataRange.address,
          rowRangeAddress: RowRange.address,
          pivotRangeAddress: PivotRange.address,
          columnRangeAddress: ColumnRange.address
      };
  
  
    });

    return RangeInfo;

  }


// 填写sum of 到 process 的新的range里，从base 和 target 抓取数据
async function fillProcessRange(SourceSheetName) {
  await Excel.run(async (context) => {
    console.log("fill process 1")
    const sheet = context.workbook.worksheets.getItem("Process");
    let ProcessRange = sheet.getRange(StrGlobalProcessRange); // 从全局变量获取Process Range 地址
    console.log("fill process")
    ProcessRange.load("address,rowCount,columnCount");

    await context.sync();
    console.log("ProcessRange is " + ProcessRange.address);

    //给全局变量Base/Target 的range 赋值地址
    if(SourceSheetName =="BasePT"){
      StrGblBaseProcessRng = ProcessRange.address
      let TempSheet = context.workbook.worksheets.getItem("TempVar"); // 将全局变量存储在TempVar中
      let VarRange = TempSheet.getRange("B2");
      let VarTitle = TempSheet.getRange("B1");
      VarRange.values = [[StrGblBaseProcessRng]];
      VarTitle.values = [["BasePT"]];
      await context.sync();
    }else if(SourceSheetName =="TargetPT"){
      StrGblTargetProcessRng = ProcessRange.address
    }
    
    //----------------在数据的上一行标明BasePT或者TargetPT的来源-----------------//
    let dataSourceLabelRange = ProcessRange.getRow(0).getOffsetRange(-1,0);
    dataSourceLabelRange.load("address, values");
    await context.sync();

    dataSourceLabelRange.values = dataSourceLabelRange.values.map(row => row.map(() => SourceSheetName));
    // await context.sync();

    let startRange = ProcessRange.getCell(0,0); // 获取左上角第一个单元格
    startRange.load("address");
    // await context.sync();

    // console.log("dataSourceLabelRange.address is " + dataSourceLabelRange.address);
    // console.log("Row is " + ProcessRange.rowCount);

    let dataRowCount = ProcessRange.rowCount - 1; // data range 的行数
    let dataColumnCount = ProcessRange.columnCount -1; // data range 的列数

    let dataStartRange = startRange.getOffsetRange(1,1); // 获取data左上角第一个单元格, 往下和往右个移动一格格子
    let dataRange = dataStartRange.getAbsoluteResizedRange(dataRowCount,dataColumnCount); // 扩大到整个dataRange

    let labelRange = startRange.getOffsetRange(0,1).getAbsoluteResizedRange(1,dataColumnCount); // 先从startRange 右移动一格，然后再扩大范围获得labelRange
    let keyRange = startRange.getOffsetRange(1,0).getAbsoluteResizedRange(dataRowCount,1); // 先从startRange 下移动一格，然后再扩大范围获得keyRange

    let PTsheet = context.workbook.worksheets.getItem(SourceSheetName);
    let pivotTable = PTsheet.pivotTables.getItem("PivotTable"); //获得basePT 或者targetPT的PT
    let PTDataRange = pivotTable.layout.getDataBodyRange(); //获得PT 的dataRange 部分
    let PTDataLastRow = PTDataRange.getLastRow(); // 获得dataRange的最后一行
    let PTLabelRow = PTDataLastRow.getOffsetRange(1,0); // 下移一行获得basePT 或者targetPT 的 下一行不带sum of的Range
    let PTRowLabelRange = pivotTable.layout.getRowLabelRange(); //获得sumif 的 criteriaRange 部分

    //console.log("fill process 3");
    //startRange.load("address");
    dataStartRange.load("address, values");
    dataRange.load("address, values");
    labelRange.load("address, values");
    keyRange.load("address, values");
    PTLabelRow.load("address, values");
    PTRowLabelRange.load("address, values");

    await context.sync();
    console.log("dataSourceLabelRange.address is " + dataSourceLabelRange.address);
    console.log("Row is " + ProcessRange.rowCount);
    //console.log("startCell is " + startRange.address);
    //console.log("dataStart is " + dataStartRange.address);
    //console.log("dataRange is " + dataRange.address);
    //console.log("labelRange is " + labelRange.address);
    await CopyFliedType(); //先填写ProcessRange 最上面的数据Type
    StrGblProcessDataRange = dataRange.address // 将dataRange 地址赋值给全局变量

    // if (SourceSheetName == "BasePT"){
      strGlbBaseLabelRange = labelRange.address // 将base的变量标题Range传递给全局函数，做进一步公式替换values
      let VarTempSheet = context.workbook.worksheets.getItem("TempVar");
      let VarBaseLabelName = VarTempSheet.getRange("B12");
      let VarBaseLableAddress =  VarTempSheet.getRange("B13");
      VarBaseLabelName.values = [["strGlbBaseLabelRange"]];
      VarBaseLableAddress.values = [[strGlbBaseLabelRange]]; //保存到临时变量工作表以便调用
      console.log(SourceSheetName + " and " + strGlbBaseLabelRange );
    // }  
    let dataRangeAddress = await GetRangeAddress("Process", dataRange.address);
    let keyRangeAddress = await GetRangeAddress("Process", keyRange.address);
      
    // 遍历dataRange每一列,每一行,每个单元格
    // for (let colIndex = 0; colIndex < dataColumnCount; colIndex++) {
      // for (let rowIndex = 0; rowIndex < dataRowCount; rowIndex++) {    
          let dataCell = dataRangeAddress[0][0];
          let labelCell = labelRange.values[0][0];
          let keyCell = keyRangeAddress[0][0];
          // dataCell.load("address, values");
          // labelCell.load("address, values");
          // keyCell.load("address, values");

          // await context.sync();
          console.log("dataCell is "+ dataCell);
          // console.log("labelCell is " +labelCell.address);
          console.log("keyCell is "+ keyCell);
          console.log("PTLabelRow is" + PTLabelRow.address);


          // 在base 或者 target PT 下面不带sum of的一行找到对应变量名在的单元格
          let targetCell = PTLabelRow.find(labelCell, {
            completeMatch: true,
            matchCase: true,
            searchDirection: "Forward"
          });

          targetCell.load("address");

          // 获取整列范围
          //let columnRange = PTsheet.getRange(columnRangeAddress);
          //let PTusedRange = columnRange.getUsedRange(); // 获得usedRange 对应的整列信息
          let PTDataRangeRow = PTDataRange.getEntireRow(); // 获得dataRange的行信息，例如3:10

          //PTusedRange.load("address");
          PTDataRangeRow.load("address");

          await context.sync();
          console.log("targetCells is " + targetCell.address);

          // ------------- 拆解targeCell 的 列，并用在base 或者 target 的ProcessRange上----------------------

          let [sheetName, cellRef] = targetCell.address.split('!');
          //   let column = cellRef.match(/^([A-Z]+)/)[0];
          let column = cellRef.match(/^([A-Za-z]+)/)[0]; 
          let columnRangeAddress = `${column}:${column}`; // 得到整列信息

          console.log("fillProcessRange 4");
          // await context.sync();

          //console.log(`Used range in column ${column}: ${PTusedRange.address}`);
          //console.log("dataRangeRow is " + PTDataRangeRow.address);

          let PTDataStartRow = PTDataRangeRow.address.split("!")[1].split(":")[0]; //拆解成Row的最上面一行
          let PTDataEndRow = PTDataRangeRow.address.split("!")[1].split(":")[1]; //拆解成Row的最下面一行

          //console.log("dataStartRow is " + PTDataStartRow);
          console.log("dataEndRow is " + PTDataEndRow);

          let PTSumRange = `${SourceSheetName}!${column}$${PTDataStartRow}:${column}$${PTDataEndRow}`; // 组合成base 或 PT里需要对应的Sum if 中的SumRange
          console.log("PTSumRange is " + PTSumRange);
          console.log("PTRowLabelRange is " + PTRowLabelRange.address);
          await insertSumIfsFormula(dataCell, PTSumRange, PTRowLabelRange.address, keyCell);
          dataRange.copyFrom(dataStartRange,Excel.RangeCopyType.formulas);
          await context.sync();

          //--------这里加入对纯数字，例如1 和 2 这样的系数不能相加，需要直接变成1 和 2的处理----------//
          
          // console.log("labelRange.address 系数处理" + labelRange.address);
 
          // let LabelRangeAddress = await GetRangeAddress("Process", labelRange.address);
          // // 使用 for 循环遍历所有单元格
          // for (let i = 0; i < labelRange.values[0].length; i++) {
          //   let cellValue = labelRange.values[0][i];
          //   console.log("系数处理 cellValue is " + cellValue);
          //   // 在 FormulaTokens 数组中查找 TokenName 与 cellValue 相同且 isNumber 为 true 的对象
          //   let matchingTokenObj = FormulaTokens.find(tokenObj => tokenObj.TokenName === cellValue && tokenObj.isNumber === true);
            
          //   // 如果找到了匹配的对象，则获取其 Token 的值
          //   if (matchingTokenObj) {
          //     const token = matchingTokenObj.Token;
          //     console.log(`在单元格 ${cell.address} 中找到匹配项，其 Token 值为：${token}`);

          //     //找到对应的要处理的labelrange的地址的column
          //     // 获取当前单元格对象

          //     let cellAddress = LabelRangeAddress[0][i];
          //     let NumTopBottomRow = getRangeDetails(dataRange.address);
          //     let NumTopRow = NumTopBottomRow.topRow;
          //     let NumBottomRow = NumTopBottomRow.bottomRow;
          //     let NumColumn = getRangeDetails(cellAddress).leftColumn;
          //     let NumRange = `${NumColumn}$${NumTopRow}:${NumColumn}$${NumBottomRow}`;

          //     // 构造二维数组
          //     const newValues = [];
          //     for (let i = 0; i < dataRowCount; i++) {
          //       newValues.push([token]);
          //     }
          //     NumRange.values = newValues;
  
          //     await context.sync();

          //     //--------这里加入对纯数字，例如1 和 2 这样的系数不能相加，需要直接变成1 和 2的处理----------//
          //   }
          // }

      // }
    // }
  });
}

// --------------------sum if 函数 插入格子------------------------------
async function insertSumIfsFormula(targetCell,sumRange, criteriaRanges, criteria) {
  try {
      await Excel.run(async (context) => {
          let criteriaAddress = getRangeDetails(criteria);
          let criteriaLeft = criteriaAddress.leftColumn;
          let criteriaTop = criteriaAddress.topRow;
          let criteriaRangesSheet = criteriaRanges.split("!")[0];
          let criteriaRangesAddress = getRangeDetails(criteriaRanges);
          let criteriaRangesLeft = criteriaRangesAddress.leftColumn;
          let criteriaRangesTop = criteriaRangesAddress.topRow;
          let criteriaRangesBottom = criteriaRangesAddress.bottomRow;


          console.log("InsertSumif 1");
          const sheet = context.workbook.worksheets.getItem("Process");
          const selectedRange = sheet.getRange(targetCell);
          console.log("InsertSumif 2");
          // Construct the SUMIFS formula
          let formula = `=SUMIFS(${sumRange}, ${criteriaRangesSheet}!$${criteriaRangesLeft}$${criteriaRangesTop}:$${criteriaRangesLeft}$${criteriaRangesBottom}, $${criteriaLeft}${criteriaTop}`;


          // Set the formula for the selected cell
          selectedRange.formulas = [[formula]];
          selectedRange.format.autofitColumns();
          console.log("InsertSumif 3");
          await context.sync();
          console.log("InsertSumif 4");
      });
  } catch (error) {
      console.error("Error: " + error);
  }
}

//-------拷贝ProcessRange, 往右偏移----------
async function copyProcessRange() {
  //console.log("pasteSheetAsValues run")
  await Excel.run(async (context) => {
    const sheetName = "Process"; 
    const sheet = context.workbook.worksheets.getItem(sheetName);
    let processRange = sheet.getRange(StrGlobalProcessRange); // 获得最初的ProcessRange 
    processRange.load("address, values, columnCount, rowCount");
    let VarianceStartRange = processRange.getCell(0,0).getOffsetRange(-1,0); // 标有目前替换变量的单元格，判断不能是TargetPT 或 BasePT
    VarianceStartRange.load("values")
    // await context.sync();

    console.log("copyProcessRange 1111111111")
    let ProcessTypeRange = processRange.getRow(0).getOffsetRange(-2,0);
    ProcessTypeRange.load("values");
    await context.sync();
    let ProcessTypeValues = ProcessTypeRange.values;
    console.log("copyProcessRange 222222")

    //搜索之前的ProcessRange是否已经开始进入Step，条件是标题上一行放置当前替换变量的地方不是TargetPT和BasePT,并有Result
    let ResultCount = 0;
    if(VarianceStartRange.values != "TargetPT" &&VarianceStartRange.values != "BasePT"){
        for (let i = 0; i < ProcessTypeValues.length; i++) {
          for (let j = 0; j < ProcessTypeValues[i].length; j++) {
              if (ProcessTypeValues[i][j] === "Result") {
                ResultCount++;
              }
          }
        }
    }
    console.log("copyProcessRange 33333")
    let nextProcessRange = processRange.getOffsetRange(0, processRange.columnCount+1+ResultCount); // ProcessRange 平移，如果进入Step开始有Impact，这需要再右移动
    nextProcessRange.load("address, values, columnCount, rowCount");
    console.log("copyProcessRange 44444")
    await context.sync();

    nextProcessRange.copyFrom(processRange);

    let dataStartRange = nextProcessRange.getCell(0,0).getOffsetRange(1,1); // ProcessRange 保留标题的起始地址
    let dataRange = dataStartRange.getAbsoluteResizedRange(nextProcessRange.rowCount-1, nextProcessRange.columnCount-1); // ProcessRange的dataRange
    dataRange.clear(Excel.ClearApplyTo.contents); // 只清除数据，保留格式
    await context.sync();

    StrGlobalPreviousProcessRange = StrGlobalProcessRange; // 在ProcessRange 往右移动前保留之前的ProcessRange
    StrGlobalProcessRange = nextProcessRange.address // 重新给全局变量赋值，后面主要时TargetRange会使用这个行数

    console.log("Before Move is ");
    console.log("Before Move is adgfadgsdfg " );
    // let NewSolveStartRange = sheet.getRange(StrGlbProcessSolveStartRange);//.getOffsetRange(0, processRange.columnCount+1); //求解变量的单元格往右平移，为后面TargetRange需要使用
    // NewSolveStartRange.load("address");
    // //console.log(NewSolveStartRange);
    // await context.sync();

    // StrGlobalProcessRange = NewSolveStartRange.address; // 求解变量的单元格往右平移，为后面TargetRange需要使用
    console.log("After Move is " + StrGlbProcessSolveStartRange);
  });


}

// 在Process Range 中拷贝求解变量的公式，继续GetFormulasAddress 第一个单元格把反算公司赋值完后，放在第这一列的所有data 单元格
async function CopyFormulas() {
  await Excel.run(async (context) => {
    const sheetName = "Process"; 
    const sheet = context.workbook.worksheets.getItem(sheetName);

    let DataRangeAddress = getRangeDetails(StrGblProcessDataRange);
    let FirstRow = DataRangeAddress.topRow

    //let FirstRow = StrGblProcessDataRange.split("!")[1].split(":")[0][1] // 获取Process Data地址的第一行的行数，例如Process!C3:G10 中的行数3
    console.log("FirstRow is " + FirstRow);
    
    let EndRow = DataRangeAddress.bottomRow

    //let EndRow = StrGblProcessDataRange.split("!")[1].split(":")[1][1] // 获取Process Data地址的第一行的行数，例如Process!C3:G10 中的行数3

    console.log("EndRow is " + EndRow);

    let Column = getRangeDetails(StrGlbProcessSolveStartRange).leftColumn

    //let Column = StrGlbProcessSolveStartRange.split("!")[1][0] // 获取第一行带有公式的地址的列数，例如Process!F4 中的F

    console.log("Column is " + Column);


    // 结合行列得出要复制的范围
    let CopyFormulasAddress= `${Column}${FirstRow}:${Column}${EndRow}`;

    console.log("CopyFormulasAddress is ZZZZZZ " + CopyFormulasAddress);
    let CopyFormulasRange = sheet.getRange(CopyFormulasAddress);


    CopyFormulasRange.copyFrom(StrGlbProcessSolveStartRange,Excel.RangeCopyType.formulas,false,false); // 将求解公式拷贝到整一列

    await context.sync();

    console.log("CopyFormulas End");
  });

}

// 从Bridge Data 中拷贝Dimension,Key,Raw data 等类型到Process 对应的字段
async function CopyFliedType() {
  console.log("CopyFiledType Here")
  await Excel.run(async (context) => {
    let SourceSheet = context.workbook.worksheets.getItem("Bridge Data");
    //console.log("CopyFiledType 2")
    //let SourceDataType = SourceSheet.getUsedRange().getRow(0);
    let SourceRange = SourceSheet.getUsedRange(); // 获取Bridge Data中的标题范围
    //console.log("CopyFiledType 3")
    let SourceDataTitle = SourceRange.getRow(1); //获得source的Title
    let SourceDataType = SourceRange.getRow(0); //获得source的Type
    console.log("CopyFiledType 4")
    //SourceDataType.load("address");
    SourceDataTitle.load("address,values,rowCount,columnCount");
    SourceDataType.load("address,values,rowCount,columnCount");
    // await context.sync();
    //console.log(SourceDataType.address)
    

    let ProcessRange = context.workbook.worksheets.getItem("Process").getRange(StrGlobalProcessRange);
    let StartRange = ProcessRange.getCell(0,0);
    ProcessRange.load("address,rowCount,columnCount");
    await context.sync();
    
    console.log("SourceDataTitle.address is " + SourceDataTitle.address);
    console.log("SourceDataType.address is " + SourceDataType.address);
    console.log("ProcessRange.address is " + ProcessRange.address);
    // 往上移动两格，从最上一行开始获取最新的ProcessRange当前的Type到Title的Range，这时候Type还没有数据
    let ProcessTitle = StartRange.getOffsetRange(0,1).getAbsoluteResizedRange(1,ProcessRange.columnCount-1); 
    let ProcessType = StartRange.getOffsetRange(-2,1).getAbsoluteResizedRange(1,ProcessRange.columnCount-1); 
    ProcessTitle.load("address,values,rowCount,columnCount");
    ProcessType.load("address,values,rowCount,columnCount");
    await context.sync();
    console.log("ProcessTitle.address is " + ProcessTitle.address);
    
    let ProcessTypeTempValues = ProcessType.values; // 临时创建二维数组，然后再存回去，这样才可以正确整体赋值。单个赋值必须用getCell方法获得单元格，效率低
    // let ProcessTitleAddress = await GetRangeAddress(ProcessTitle.address);

    
    TitleColumnCount = ProcessTitle.columnCount;
    TitleRowCount = ProcessTitle.rowCount;
    console.log("TitleColumnCount is " + TitleColumnCount);
    console.log("TitleRowCount is " + TitleRowCount);
      // 和 Bridge Data里的Title Range 逐个对比
      // for (let rowIndex = 0; rowIndex < TitleRowCount; rowIndex++) {  
    console.log("CopyFiledType 5");
    // NextLoop:

    for (let colIndex = 0; colIndex < TitleColumnCount; colIndex++) { 

        const foundCell = SourceDataTitle.find(ProcessTitle.values[0][colIndex], {
          completeMatch: true,
          matchCase: false,
          searchDirection: "Forward",
        });
        console.log("CopyFiledType 5.2");
        let TypeCell = foundCell.getOffsetRange(-1,0);
        TypeCell.load("values");
        await context.sync();
        console.log("CopyFiledType 5.3");
        ProcessTypeTempValues[0][colIndex] = TypeCell.values[0][0];
        console.log("CopyFiledType 5.4");
          // let TitleCell = ProcessTitle.getCell(rowIndex,colIndex);
          // TitleCell.load("address, values");

          // await context.sync();

          // console.log("TitleCell value is " + TitleCell.values[0][0]);
          // //在Bridge Data中找到对应的Title单元格
          // let SourceTitleCell = SourceDataTitle.find(TitleCell.values[0][0], {
          //   completeMatch: true,
          //   matchCase: true,
          //   searchDirection: "Forward"
          // });
          
          // SourceTitleCell.load("address,values");

          // await context.sync();
          //在BridgeData最上面两行中循环，找到对应Title的Type
          // console.log("SourceDataTitle.values[0].length is " + SourceDataTitle.values[0].length);
          
 
          // //console.log("SourceTitleCell values is " + SourceTitleCell.values);
          // let ProcessTypeCell = TitleCell.getOffsetRange(-2,0);
          // let SourceTypeCell = SourceTitleCell.getOffsetRange(-1,0);
          // SourceTypeCell.load("address,values");
          // ProcessTypeCell.load("address,values");
          // await context.sync();

          // console.log("SourceTypeCell address is " + SourceTypeCell.address);
          // console.log("ProcessTypeCell address is " + ProcessTypeCell.address);
          // console.log("SourceTypeCell values[0][0] is " + SourceTypeCell.values[0][0] );
          

          // ProcessTypeCell.values = [[SourceTypeCell.values[0][0]]]; // values 是二维数组，只能对二维数组整体赋值
          //ProcessTypeCell.values = SourceTypeCell.values[0][0]; 这样的赋值方法是错误的
          //ProcessTypeCell.values[0][0] = SourceTypeCell.values[0][0] 这样赋值也是错误的
          if(TypeCell.values[0][0] == "Result" && NumVarianceReplace > 0){

            TitleColumnCount = TitleColumnCount -1 ; //若有一个Result 并且替换变量从第二个开始，则列数减一，否则在Bridge Data中的Title Range 列数会比Step中的少

          }

          // await context.sync();
          // console.log("ProcessTypeCell values[0][0] is " + ProcessTypeCell.values[0][0] );

      // }
    }
    ProcessType.values = ProcessTypeTempValues;
    console.log("CopyFiledType 6");
    await context.sync();


  });

}



//需要对 key 中的特殊字符进行转义，这样它们在正则表达式中将被视为普通字符
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& 表示匹配的整个字符串
}

// 0728 获得BridgeData中非Dimension字段的类型，以及判断是否有公式，存在对象中，在Process的Step中使用，并复制到一整列

async function GetBridgeDataFieldFormulas() {
  return await Excel.run(async (context) => {
    let DataSheetName = "Bridge Data Temp";
    let BridgeDataSheet = context.workbook.worksheets.getItem(DataSheetName);
    let BridgeUsedRange = BridgeDataSheet.getUsedRange();
    let BridgeTitleRange = BridgeUsedRange.getRow(1); // 获取第二行标题行

    BridgeTitleRange.load("address,values,rowCount,columnCount");
    await context.sync();

    console.log("BridgeTitleRange is " + BridgeTitleRange.address);

    // 获得Bridge工作表从第一行到第三行的数据
    let BridgeTitleStart = BridgeTitleRange.getCell(0, 0);
    let BridgeRange = BridgeTitleStart.getOffsetRange(-1, 0).getAbsoluteResizedRange(3, BridgeTitleRange.columnCount);

    BridgeRange.load("address,values,formulas");
    await context.sync();

    console.log("BridgeRange is " + BridgeRange.address);

    // 获取每个单元格的地址（若你已有函数GetRangeAddress，可以沿用）
    let BridgeRangeAddress = await GetRangeAddress(DataSheetName, BridgeRange.address);

    let TitleRowCount = BridgeTitleRange.rowCount; // 这里应是 1
    let TitleColumnCount = BridgeTitleRange.columnCount; // 列数
    console.log("row is " + TitleRowCount);
    console.log("column is " + TitleColumnCount);

    // 用来收集每列的相关数据
    let bridgeDataArray = [];

    // 遍历列
    for (let j = 0; j < TitleColumnCount; j++) {
      let TitleCell = BridgeRange.values[1][j]; // 第二行（索引1）的内容
      let TitleType = BridgeRange.values[0][j]; // 第一行（索引0）的内容
      let TitleValue = BridgeRange.values[2][j]; // 第三行（索引2）的内容
      let TitleValueFormulas = BridgeRange.formulas[2][j];
      let TitleValueAddress = BridgeRangeAddress[2][j];

      console.log("TitleCell is " + TitleCell);
      console.log("TitleType is " + TitleType);
      console.log("TitleValue is " + TitleValue);
      console.log("TitleValueFormulas is " + TitleValueFormulas);
      console.log("TitleValueAddress is " + TitleValueAddress);

      // 只有 RngFormulas 有实际公式时，我们才会改它
      let RngFormulas = null;
      // 可能要存储的公式-变量映射对象
      let FormulaVarTitle = null;

      // 如果 TitleType != Dimension / Key / Null，说明允许有公式, checkType2Var不能包含，因为这个变量时新生成的，因此有公式，公式里的变量已经被删除
      if (TitleType !== "Dimension" && TitleType !== "Key" && TitleType !== "Null" && ArrVarPartsForPivotTable.includes(TitleCell) && !checkType2Var.includes(TitleCell)) {
        // 判断单元格内是否实际有公式
        if (TitleValueFormulas !== TitleValue) {
          // 不相等 => 有公式
          console.log("there is formulas: " + TitleValueFormulas);

          // 去掉$符号
          RngFormulas = TitleValueFormulas.replace(/\$/g, "");

          // 获取公式中的变量-标题映射
          FormulaVarTitle = await getFormulaObj(DataSheetName, TitleValueAddress);
          console.log("FormulaVarTitle is:");
          console.log(JSON.stringify(FormulaVarTitle, null, 2));

          // 将RngFormulas里面的单元格地址替换为标题
          for (let title in FormulaVarTitle) {
            let cellAddress = FormulaVarTitle[title];
            let cellAddressRegex = new RegExp(cellAddress, "g");
            RngFormulas = RngFormulas.replace(cellAddressRegex, title);
          }
          console.log("RngFormulas is " + RngFormulas);
          // —— 将这一列的关键数据保存到数组中 ——
          bridgeDataArray.push({
            columnIndex: j, // 当前列索引（可选，方便后续识别是哪一列）
            TitleType,
            TitleCell,
            TitleValue,
            TitleValueFormulas, // 原始公式（可能为空或纯字符串）
            TitleValueAddress,
            RngFormulas, // 处理后的公式（若无公式则null）
            FormulaVarTitle // 公式中提取到的映射（若无则null）
          });
        }
      }
    }

    await context.sync();

    // 返回这次收集到的数据
    return bridgeDataArray;
  });
}


// 从Bridge Data工作表获得含有公式的变量对象，在
async function putFormulasToProcess(TitleFormulasArr) {
  await Excel.run(async (context) => {
    
    let ProcessSheetName = "Process";
    let ProcessSheet = context.workbook.worksheets.getItem(ProcessSheetName);
    let ProcessStepRange = ProcessSheet.getRange(StrGlobalProcessRange); // 获得全局变量中当前的Process中的Range,已经右移动
    let ProcessRange = ProcessStepRange.getRow(0);
    let ProcessStartRng = ProcessStepRange.getCell(0,0);
    let ProcessDataStartRng = ProcessStartRng.getOffsetRange(1,1);

    ProcessStepRange.load("address,values,rowCount,columnCount");
    ProcessRange.load("address");

    await context.sync();

    console.log("ProcessRange is " + ProcessRange.address);
    console.log("ProcessStepRange.rowCount is " + ProcessStepRange.rowCount);
    console.log("ProcessStepRange.column is " + ProcessStepRange.columnCount);
    //获得Bridge工作表从第一行到第三行的数据

    let ProcessDataRng = ProcessDataStartRng.getAbsoluteResizedRange(ProcessStepRange.rowCount-1,ProcessStepRange.columnCount-1); //扩大到整个目前的DataRang
    ProcessDataRng.load("address");
    console.log("0728 here");
    await context.sync();

    console.log("ProcessDataRng is " + ProcessDataRng.address)
    console.log("TitleFormulasArr is:")

    console.log(TitleFormulasArr);

          for (const TitleFormulasObj of TitleFormulasArr) {
            // let TitleCell = BridgeTitleRange.getCell(0,j); // 获取字段名
              let FormulaVarTitle = TitleFormulasObj.FormulaVarTitle;
              console.log("Before FormulaVarTitle is ")
              console.log(FormulaVarTitle);
              //下面要先对公式里的变量排序，不然可能会导致Test 和 Test 2 重复替换错误
              const sortedKeys = Object.keys(FormulaVarTitle).sort((a, b) => b.length - a.length);

              // 重新构建排序后的对象
              const sortedFormulaVarTitle = {};
              for (let key of sortedKeys) {
                sortedFormulaVarTitle[key] = FormulaVarTitle[key];
              }
              // 现在 sortedFormulaVarTitle 是排序后的对象
              FormulaVarTitle = sortedFormulaVarTitle;
              console.log("After FormulaVarTitle is ");
              console.log(FormulaVarTitle);

              let RngFormulas = TitleFormulasObj.RngFormulas;
              console.log("Before RngFormulas is " + RngFormulas);

              let TitleCell = TitleFormulasObj.TitleCell;

                      //为在Process 中处理替换成Process对应的的变量地址
                      for (let title in FormulaVarTitle) {      
                      
                          let ProcessTitleCell = ProcessRange.find(title, {
                                                  completeMatch: true, 
                                                  matchCase: true, 
                                                  searchDirection: "Forward"
                          });
                          ProcessTitleCell.load("address");
                          await context.sync();
                          console.log("title is " + title);
                          console.log("ProcessTitleCell is " + ProcessTitleCell.address);
                        let ProcessCell = ProcessTitleCell.getOffsetRange(1, 0); //往下一行才是数据的地址
                        ProcessCell.load("address");
                        await context.sync();
                        let ProcessCellAddress = ProcessCell.address.split("!")[1]; 
                          //let cellAddressRegex = new RegExp(cellAddress, "g");

                          const escapedTitle = escapeRegExp(title); // 转义后的 title

                          RngFormulas = RngFormulas.replace(new RegExp(`(?<![\\w\\d_])${escapedTitle}(?![\\w\\d_])`, 'g'), ProcessCellAddress).replace("=",""); // 这里必须用正则表达式，不然变量出现两次只会替换第一次。新公式为标题代替变量, 把 = 号去掉，下面替换成=IFERROR（////替换的时候可能有相同字符在一个变量标题里，需要处理 */
                          //这里必须进一步考虑，可能会有Revenue 和 RevenueAAA等 变量有重复的会被错误替换的问题 */
                          //(?<![\\w])：负向前瞻断言，确保 title 前面不是字母、数字或下划线（即不在单词的中间）。
                          // title：目标替换字符串。
                          // (?![\\w])：正向后瞻断言，确保 title 后面不是字母、数字或下划线（即不在单词的中间）。
                          // 这种方法适用于更复杂的场景，如包含空格的字符串。
                      }
                      
                      //若公式是（1+0），则上面这个内部的for循环不会执行，因为FormulaVarTitle中没有数据。因此没有下面的replace，则包含等号，=IFERROR(${RngFormulas},0) 这一步就会出错
                      RngFormulas = RngFormulas.replace("=",""); 
                      console.log("after RngFormulas is " + RngFormulas);

                      console.log("TitleCell.values[0][0] is " + TitleCell);
                      //在process中找到对应的公式应该输入的单元格
                      let ProcessFormulaCell = ProcessRange.find(TitleCell, {
                            completeMatch: true,
                            matchCase: true,
                            searchDirection: "Forward"
                          });
                      let NextRowFormulaCell = ProcessFormulaCell.getOffsetRange(1,0);
                      NextRowFormulaCell.formulas = [[`=IFERROR(${RngFormulas},0)`]];//往下一行填入公式

                      NextRowFormulaCell.load("address");
                      await context.sync();

                      //---------------开始把这个公式复制到一整列---------------------------------                 
                                                
                      let ProcessRngDetail = getRangeDetails(ProcessDataRng.address); // 返回的是一个对象
                      let FirstRow = ProcessRngDetail.topRow;
                      let EndRow = ProcessRngDetail.bottomRow -1 ; // 最后一行是Total，因此不能用一行的公式计算，需要计算列的和

                      let Column = getRangeDetails(NextRowFormulaCell.address).leftColumn;
                      console.log("FirstRow is " + FirstRow);
                      console.log("EndRow is " + EndRow);
                      console.log("Column is " + Column);

                      // 结合行列得出要复制的范围
                      let CopyFormulasAddress= `${Column}${FirstRow}:${Column}${EndRow}`;
                  
                      console.log("CopyFormulasAddress is AAAAA " + CopyFormulasAddress);
                      let CopyFormulasRange = ProcessSheet.getRange(CopyFormulasAddress);

                      CopyFormulasRange.copyFrom(NextRowFormulaCell,Excel.RangeCopyType.formulas,false,false); // 将求解公式拷贝到整一列，除了最后一行
                  
                      await context.sync();
                  
        }
    // }
    // await context.sync();

  });

}

// --------------------获取单元格的公式，并形成对象------------------
async function getFormulaObj(sheetName, formulaAddress) {
  
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    console.log("formulaAddress is:  " + formulaAddress)
    const formulaCell = sheet.getRange(formulaAddress);
    
    formulaCell.load("formulas, values, address");
    await context.sync();
    console.log("formulacell is " + formulaCell.address)
    
    //console.log("formulaCell.values is " + formulaCell.values[0][0])
    // const cellValue = formulaCell.formulas[0][0];
    // console.log("cellValue is: " + cellValue)
    // // if (typeof cellValue !== "string") {
    // //   console.error("The cell value is not a string or is empty????.");

    // //   return {};
    // // }
    
    const formula = formulaCell.formulas[0][0].replace(/\$/g, ""); // 去除公式里的$固定符号

    //console.log(formula);
    // const cellReferenceRegex = /([A-Z]+[0-9]+)/g;
    const cellReferenceRegex = /([A-Za-z]+)(\d+)/g;
    const cellReferences = formula.match(cellReferenceRegex);
    
    console.log("cellReferences is: "+ cellReferences)
    if (!cellReferences) {
      console.log("No cell references found in the formula.");
      return {};
    }

    const cellTitles = {}; // 创建一个对象
    
    for (const cellReference of cellReferences) {
    //   const match = cellReference.match(/([A-Z]+)([0-9]+)/);
        const match = cellReference.match(/([A-Za-z]+)(\d+)/);
      if (match) {
        
        const column = match[1];
        const row = parseInt(match[2]);
        const titleCellAddress = `${column}${row - 1}`;
        const titleCell = sheet.getRange(titleCellAddress);
        titleCell.load("values");
        await context.sync();
        const title = titleCell.values[0][0];
        cellTitles[title] = cellReference;
      }
    }
    //console.log("getFormulaCellTitles end")
    //console.log(cellTitles);
    return cellTitles;
  });
}


// 获得Range 四周的行数和列数的信息
//如果范围字符串是单个单元格（例如 AD9），则结束列和结束行与起始列和起始行相同。
//返回一个对象，包含 topRow、bottomRow、leftColumn 和 rightColumn 四个属性。
function getRangeDetails(rangeStr) {
  // 使用正则表达式提取列和行信息
  //   const regex = /([A-Z]+)(\d+):?([A-Z]+)?(\d+)?/;
    const regex = /([A-Za-z]+)(\d+):?([A-Za-z]+)?(\d+)?/;
  const match = rangeStr.match(regex);

  if (match) {
    const startColumn = match[1];
    const startRow = parseInt(match[2], 10);
    const endColumn = match[3] ? match[3] : startColumn;
    const endRow = match[4] ? parseInt(match[4], 10) : startRow;
    return {
      topRow: startRow,
      bottomRow: endRow,
      leftColumn: startColumn,
      rightColumn: endColumn
    };
  } else {
    throw new Error("Invalid range format");
  }
}


//开始在Step 中从第一个变量循环遍历并替代, 并填充Step中所有相应不同的格子
async function VarFromBaseTarget() {
  console.log("VarFromBaseTarget Start");
  await Excel.run(async (context) => {
      let Sheet = context.workbook.worksheets.getItem("Process");
      let ProcessRange = Sheet.getRange(StrGlobalProcessRange); // 获得全局变量中当前的Process中的Range,已经右移动
      let StartRange = ProcessRange.getCell(0,0);
      StartRange.load("address");
      ProcessRange.load("address,rowCount,columnCount");
      console.log("VarFromBaseTarget 1111111")
      await context.sync();

      console.log("StartRange is " +  StartRange.address);

      let BaseRange = Sheet.getRange(StrGblBaseProcessRng); // 从全局变量中获得BaseRange
      let BaseLastRow = BaseRange.getLastRow();
      let TargetRange = Sheet.getRange(StrGblTargetProcessRng); //从全局变量中获得TargetRange
      let DataRangeExclSum = StartRange.getOffsetRange(1,1).getAbsoluteResizedRange(ProcessRange.rowCount-2,ProcessRange.columnCount-1); // 获得不包括最下面一行加总的数据Range
      BaseRange.load("address,values");
      BaseLastRow.load("address,values");
      TargetRange.load("address,values");
      DataRangeExclSum.load("address");

      console.log("VarFromBaseTarget 222222222222")
      // await context.sync();

      let TitleLength = ProcessRange.columnCount -1
      let TitleRange = StartRange.getOffsetRange(0,1).getAbsoluteResizedRange(1,TitleLength ); // 变量标题行
      let TitleAndTypeRange = StartRange.getOffsetRange(-2,1).getAbsoluteResizedRange(4,TitleLength ); // 变量标题行和数据类型行,再加上第一行数据地址
      let CurrentVarRange = TitleRange.getOffsetRange(-1,0); // 标题往上一行，一整行输入目前正在替换的变量
      let TypeRange = TitleRange.getOffsetRange(-2,0); // 标题往上两行，获得对应的TypeRange
      let SumRange = TitleRange.getOffsetRange(ProcessRange.rowCount - 1,0); // 获得加总行

      let PreProcessRange = Sheet.getRange(StrGlobalPreviousProcessRange); // 获得上一步的PreProcessRange
      PreProcessRange.load("address,values");

      console.log("VarFromBaseTarget 333333333")

      TitleRange.load("address,values");
      TitleAndTypeRange.load("address,values");
      TypeRange.load("address,values")
      SumRange.load("address")
      await context.sync();
      console.log("VarFromBaseTarget 4444444")

      console.log("TargetRange address is " + TargetRange.address);
      console.log("BaseLastRow address is " + BaseLastRow.address);
      console.log("TitleAndTypeRange address is " + TitleAndTypeRange.address);

      let TitleAndTypeRangeAddress = await GetRangeAddress("Process",TitleAndTypeRange.address);
      console.log("TitleAndTypeRangeAddress is");
      console.log(TitleAndTypeRangeAddress);

      let SumAddress = getRangeDetails(DataRangeExclSum.address);
      let SumTopRow = SumAddress.topRow; //加总数据的其实行
      let SumBottomROw = SumAddress.bottomRow; //加总数据的结束行
      let SumRow = getRangeDetails(SumRange.address).bottomRow; //最后一行汇总行
      let VarRow = getRangeDetails(TitleRange.address).bottomRow + 1; //变量为标题的下一行行数
      
      // 开始循环变量并替换
      console.log("VarFromBaseTarget 55555555")
      for(let i =0 ;i<TitleLength; i++){
          let TitleCell = TitleAndTypeRangeAddress[2][i];
          let TypeCell= TitleAndTypeRangeAddress[0][i]; // 获取每个变量的Type
          let TitleCellValues = TitleAndTypeRange.values[2][i];
          let TypeCellValues = TitleAndTypeRange.values[0][i];
          // TitleCell.load("address,values");
          // TypeCell.load("address,values");
          // await context.sync();

          console.log("TitleCell is " + TitleCell);

          //NumVarianceReplace 也是从0开始 // 若小于等于则从Target中替换变量, 不相等从Base中获取原来的变量，但是标题不能等于Result(需要用公式计算)
          //在找到替换变量同时，计算相应的impact
          console.log("I is "+ i +"; NumVarianceReplace is " + NumVarianceReplace);
          if(i <= NumVarianceReplace && TypeCellValues != "Result"){
              let TargetCell = TargetRange.find(TitleCellValues, { //要在targetRange 中找到对应替换变量的单元格
                completeMatch: true,
                matchCase: true,
                searchDirection: "Forward"
              });
              TargetCell.load("address,values");
              await context.sync();
              console.log("TargetCell values is " + TargetCell.values);
              console.log("TargetCell address is " + TargetCell.address);

              let VarColumn = getRangeDetails(TargetCell.address).leftColumn;

              let VarInputCell= Sheet.getRange(TitleAndTypeRangeAddress[3][i])     //TitleCell.getOffsetRange(1,0); //变量输入单元格
              console.log("VarColumn is " + VarColumn);
              VarInputCell.values = [[`=${VarColumn}${VarRow}`]]; // 将变量等于target的值
              console.log("after VarInputCell");

              if(i ==  NumVarianceReplace && TypeCellValues != "Result"){
                  console.log("VarFromBaseTarget 555666");
                  let TitleCellRange = Sheet.getRange(TitleCell); //为了下面的copyfrom, 需要得到单元格，后面再同步
                  CurrentVarRange.copyFrom(TitleCellRange,Excel.RangeCopyType.values); // 给标题的上一行输入目前正在替换的变量
              
                      //-----------下面开始在右边新建一列作为对应变量变化产生的Impact------------------------
                      let ResultTypeRange = TypeRange.find("Result", {
                        completeMatch: true,
                        matchCase: true,
                        searchDirection: "Forward"
                      });

                      ResultTypeRange.load("address");
                      await context.sync();
                      console.log("ResultTypeRange is " + ResultTypeRange.address);
                      console.log("TypeRange is "  + TypeRange.address);

                      console.log("Impact 11111111")
                      let ResultTitleRange = ResultTypeRange.getOffsetRange(2,0); //往下移动两行，获得result对应的变量标题
                      ResultTitleRange.load("address,values");
                      //await context.sync();
                      console.log("Impact 2222222")
                      let ImpactTitleRange = StartRange.getOffsetRange(0,ProcessRange.columnCount); //在ProcessRange的最右边的格子
                      ImpactTitleRange.load("address");
                      console.log("Impact 33333")
                      await context.sync();
                      //console.log("StartRange is " +  StartRange.address);
                      //console.log("ProcessRange is " + ProcessRange.address);
                      //console.log("ImpactTitleRange is " + ImpactTitleRange.address); // 获得Impact标题的单元格
                      let ProcessRangeAddress = getRangeDetails(ProcessRange.address);
                      let ResultTopRow = ProcessRangeAddress.topRow;
                      let ResultBottomRow = ProcessRangeAddress.bottomRow;
                      let ResultColumn = getRangeDetails(ResultTitleRange.address).leftColumn;
                      let ResultRange = Sheet.getRange(`${ResultColumn}${ResultTopRow}:${ResultColumn}${ResultBottomRow}`);
                      console.log("Impact 44444")
                      ResultRange.load("address,format");
                      await context.sync();
                      console.log("ResultRange is " + ResultRange.address);

                      ImpactTitleRange.copyFrom(ResultRange,Excel.RangeCopyType.formats);// 复制前面Result 一列的格式，这里Format需要加s
                      await context.sync(); // 

                      //ImpactTitleRange.values = ResultTitleRange.values // 可以这样直接赋值~！
                      ImpactTitleRange.values =[[ResultTitleRange.values[0][0] + " Impact"  ]];  // 加上impact的标题
                      let ImpactVarRange = ImpactTitleRange.getOffsetRange(-1,0); //往上移动一格获得Impact对应的变量
                      let ImpactTypeRange = ImpactTitleRange.getOffsetRange(-2,0); //往上移动两格输入Impact这个类型 
                      console.log("Impact 55555")
                      ImpactVarRange.values = TitleCellValues; // 直接等于当前变量
                      console.log("Impact 66666")
                      ImpactTypeRange.values = [["Impact"]]; // 差异新的type，不能这样赋值 ImpactTypeRange.values[0][0] = [["Impact"]]
                      console.log("Impact 77777")
                      await context.sync();

                      // --------------------在impact 单元格中放入对应的计算公式--------------------
                      console.log("TitleCell.values is " + TitleCellValues);
                      console.log("PreProcessRange is " + PreProcessRange.address);
                      PreProcessRangeFirstRow = PreProcessRange.getRow(0);
                      PreProcessRangeFirstRow.load("address,values");
                      await context.sync();

                      console.log("PreProcessRangeFirstRow is " + PreProcessRangeFirstRow.address);

                      let PreResultTitleCell = PreProcessRangeFirstRow.find(ResultTitleRange.values[0][0], { // 这里在PreProcessRange中找对应的单元格，而不是在Target中找, 必须是TitleCell.values[0][0]，而不是TitleCell.values
                        completeMatch: true,
                        matchCase: true,
                        searchDirection: "Forward"
                      });

                      PreResultTitleCell.load("address,values");
                      console.log("After find PreResultTitleCell")
                      await context.sync();

                      console.log("PreResultTitleCell is " + PreResultTitleCell.address);

                      PreProcessResultColumn = getRangeDetails(PreResultTitleCell.address).leftColumn; // 获得preProcess对应的column
                      ImpactColumn = getRangeDetails(ImpactTitleRange.address).leftColumn; //获得Impact的column
                      console.log("PreProcessResultColumn is " + PreProcessResultColumn);
                      console.log("ImpactColumn is " + ImpactColumn);
                      
                      let ImpactDataFirstRow= ImpactTitleRange.getOffsetRange(1,0); // Impact标题往下移动一格
                      let ImpactDataRange = Sheet.getRange(`${ImpactColumn}${ResultTopRow+1}:${ImpactColumn}${ResultBottomRow}`); // 拼凑出ImpactData对应的Range
                      ImpactDataFirstRow.formulas = [[`=${ResultColumn}${ResultTopRow+1}-${PreProcessResultColumn}${ResultTopRow+1}`]]
                      ImpactDataFirstRow.load("formulas");
                      await context.sync();

                      ImpactDataRange.copyFrom(ImpactDataFirstRow,Excel.RangeCopyType.formulas) // 在Impact 列拷贝公式

                      ProcessRange = StartRange.getAbsoluteResizedRange(ProcessRange.rowCount,ProcessRange.columnCount); // 每次有一个Result 对应的Impact产生ProcessRange就往右加一列
                      ProcessRange.load("address,rowCount,columnCount"); // 重新加载，以防万一引用更新的Range出错
                      await context.sync();
            }
              StrGlobalProcessRange = ProcessRange.address // 更新全局变量

              await context.sync();
              console.log("VarFromBaseTarget 5555777");

          }else if(TypeCellValues != "Result"){   //若不是当前需要改变的变量，则等于Base的值

              let BaseCell = BaseRange.getRow(0).find(TitleCellValues, {
                    completeMatch: true,
                    matchCase: true,
                    searchDirection: "Forward"
                  });
              BaseCell.load("address");
              await context.sync();

              let VarColumn = getRangeDetails(BaseCell.address).leftColumn;
              let VarInputCell= Sheet.getRange(TitleAndTypeRangeAddress[3][i]); //变量输入单元格
              VarInputCell.values = [[`=${VarColumn}${VarRow}`]];
          }

          //给最后一行加如汇总公式, 不是SumN 且也不是result 从数据行加总，SumN 或 Result 则从base 同一行获得公式
          let SumCell = SumRange.getCell(0,i);
          let SumColumn = getRangeDetails(TitleCell).leftColumn;
          if(TypeCellValues != "SumN" && TypeCellValues != "Result"){
              SumCell.values =[[`=SUM(${SumColumn}${SumTopRow}:${SumColumn}${SumBottomROw})`]]

          }else{
            
            let BaseCell = BaseRange.getRow(0).find(TitleCellValues, {
              completeMatch: true,
              matchCase: true,
              searchDirection: "Forward"
            });
            BaseCell.load("address,formulas");
            await context.sync();

            let VarColumn = getRangeDetails(BaseCell.address).leftColumn;
            let VarRange = Sheet.getRange(`${VarColumn}${SumRow}`);
            VarRange.load("address,formulas");
            await context.sync();
            console.log("VarRange address is " + VarRange.address);

            //let VarColumn = getRangeDetails(BaseCell.address).leftColumn;
            //let BaseFormulas = Sheet.getRange(`${VarColumn}${SumRow}`)
            //SumCell.formulas =[[VarRange.formulas[0][0]]];
            SumCell.copyFrom(VarRange);
          }
      }
      //console.log("ReadyToCopy");

      console.log("DataRangeExclSum is " + DataRangeExclSum.address)

      // 给中间dataRange复制data第一行同样的数据
      DataRangeExclSum.copyFrom(TitleRange.getOffsetRange(1,0));

      NumVarianceReplace = NumVarianceReplace +1; // 一个Step完成后，全局变量+1，为下一个Step的处理计数

      await context.sync(); // 复制完以后这一行一定要加

  });
}


// --------------------获取Base ProcessRange中变量的个数------------------
async function GetNumVariance() {
  
  return await Excel.run(async (context) => {
    let ProcessSheet = context.workbook.worksheets.getItem("Process");
    let BaseRange = ProcessSheet.getRange(StrGblBaseProcessRng);
    let StartRange = BaseRange.getCell(0,0);
    BaseRange.load("address,rowCount,columnCount");

    await context.sync();

    let BaseTypeRange = StartRange.getOffsetRange(-2,1).getAbsoluteResizedRange(1,BaseRange.columnCount-1);
    BaseTypeRange.load("address,rowCount,columnCount,values");

    await context.sync();

    // let VarCount = 0;
    // //若base title不是Result,则作为变量需要计数
    // for(let i =0;i<BaseTitleRange.columnCount;i++){
    //     let BaseCell = BaseTitleRange.getCell(0,i);
    //     BaseCell.load("address,values");
    //     await context.sync();
    //     if(BaseCell.values != "Result"){
    //           VarCount++;

    //     }
    // }
    // 假设 BaseTypeRange.values 是一个二维数组
    let baseTypeValues = BaseTypeRange.values;

    // 遍历二维数组并移除值为 "ProcessSum" 的元素
    for (let i = 0; i < baseTypeValues.length; i++) {
        // 使用 filter 去掉每一行中值为 "ProcessSum" 的元素
        baseTypeValues[i] = baseTypeValues[i].filter(value => value !== "ProcessSum" && value !== "Null"); // 这个不应该成为变量替换的一部分
    }

    return baseTypeValues; // 虽然只有一行，但是是一个二维数组

  });
}

//------------根据变量的类型循环执行变量替换的步骤-----------------
async function VarStepLoop(VarFormulasObjArr) {
  
  await Excel.run(async (context) => {
      let Variance = await GetNumVariance(); // 返回一个二维数组
      console.log("Variance[0].length is " + Variance[0].length );

      for(let i = 0; i < Variance[0].length; i++){
        console.log("Variance[0][i] is " + Variance[0][i]);
        if(Variance[0][i]!="Result"){
            

            await copyProcessRange(); // 生成Step1
            await CopyFliedType(); // 获得字段的type

            await VarFromBaseTarget();
            // await GetBridgeDataFieldFormulas(); // 将Bridge Data中带有公式的拷贝到StepRange中
            await putFormulasToProcess(VarFormulasObjArr);
        }else{
            NumVarianceReplace++; // 这里跳过了Result,但是整体替换变量的个数还是往前走了一步，Result在ProcessRangeTitle中也算循环中的一个变量个数
            
          }
      };
      NumVarianceReplace = 0 ; //执行完全部循环后必须清零，不然程序会持续往下加
  });
}


//------------在Process中查找Base 和 Target中的Result 作为Bridge两端，找Impact和对应的变量作为中间的变化-----------------
//-----这里优化的时候，下边框range没有固定在base range 的最下边，而是usedRange的最下边，如果用户如果输入数据，那么下边的行数会变，代码会出错
//-----需要提示用户不能修改，或者干脆禁止修改，或者修改代码，控制在BaseRange最下边一行
async function BridgeFactors() {
  return await Excel.run(async (context) => {
    let Sheet = context.workbook.worksheets.getItem("Process");
    let OldUsedRange = Sheet.getUsedRange(); //Process 中使用的Range
    OldUsedRange.load("address,values,rowCount,columnCount");
    let TempSheet = context.workbook.worksheets.getItem("TempVar");
    let VarRange = TempSheet.getRange("B2"); //从临时表中获取BasePT Range的全局变量
    VarRange.load("values");
    await context.sync();
    //let BaseRange = Sheet.getRange(StrGblBaseProcessRng); //BasePT的Range
    let BaseRange = Sheet.getRange(VarRange.values[0][0]); //BasePT的Range


    BaseRange.load("address");
    await context.sync();
    console.log("BridgeFactors 1");

    let UsedRangeAddress = getRangeDetails(OldUsedRange.address);
    let UsedRngLeftColumn = UsedRangeAddress.leftColumn;
    let UsedRngRightColumn = UsedRangeAddress.rightColumn;
    let UsedRngTopRow = UsedRangeAddress.topRow;
    let BaseRangeAddress = getRangeDetails(BaseRange.address);
    let BaseRngTopRow= BaseRangeAddress.topRow;
    let BaseRngBottomRow= BaseRangeAddress.bottomRow;
    // //形成UsedRange
    let UsedRange = Sheet.getRange(`${UsedRngLeftColumn}${UsedRngTopRow}:${UsedRngRightColumn}${BaseRngBottomRow}`);
    UsedRange.load("address,values,rowCount,columnCount");
    await context.sync();
    console.log("UsedRange is " + UsedRange.address);
    // TypeRange.load("address,values,rowCount,columnCount");

    // let CurrentVarRange = Sheet.getRange(`${UsedRngLeftColumn}${BaseRngTopRow-1}:${UsedRngRightColumn}${BaseRngTopRow-1}`);
    // CurrentVarRange.load("address,values,rowCount,columnCount");

    // let ImpactRange = Sheet.getRange(`${UsedRngLeftColumn}${BaseRngBottomRow}:${UsedRngRightColumn}${BaseRngBottomRow}`);
    // ImpactRange.load("address,values,rowCount,columnCount");

    // await context.sync();
    
    // console.log("TypeRange is " + TypeRange.address);
    // let TypeRangeAddress = await GetRangeAddress("Process",TypeRange.address);
    // let UsedRangeDetail = await GetRangeAddress("Process",UsedRange.address);
    // console.log("CurrentVarRange is " + CurrentVarRange.address);
    // console.log("ImpactRange.address is " + ImpactRange.address);

    let BridgeFactors ={}; // 包含Bridge 中每个Factor的信息
    let RowCount = UsedRange.rowCount;
    let ColumnCount = UsedRange.columnCount;
    console.log("BridgeFactors 2");
    //循环查找TypeCell中的变量，并相应的信息放入对象中，包括（Result/Impact,TargetPT/当前替换变量，受影响的变量名，Impact的值）
    for(let Col = 0;Col < ColumnCount; Col ++){ //在第一行Type上循环
            // let TypeCell = TypeRange.getCell(0,TypeCount);
            // TypeCell.load("address,values");
            // await context.sync();
            // TypeCellColumn = getRangeDetails(TypeCell.address).leftColumn;
            // let CurrentVarCell = TypeCell.getOffsetRange(1,0);
            // CurrentVarCell.load("address,values");
            // let TitleCell = TypeCell.getOffsetRange(2,0);
            // TitleCell.load("address,values");
            // let SumCell = Sheet.getRange((`${TypeCellColumn}${BaseRngBottomRow}`));//获得Sum行对应单元格，Impact的总影响
            // SumCell.load("address,values");
            
            // await context.sync();

            let SumCellValues = UsedRange.values[RowCount-1][Col];
            let CurrentVarCellValues = UsedRange.values[1][Col];
            let TypeCellValues = UsedRange.values[0][Col];
            let TitleCellValues = UsedRange.values[2][Col];
            if(TypeCellValues == "Result" && (CurrentVarCellValues == "BasePT" || CurrentVarCellValues == "TargetPT")){
                //将Bridge头尾两端找到放入对象
                BridgeFactors[CurrentVarCellValues] ={
                    Type: TypeCellValues,
                    Title:TitleCellValues,
                    Sum: SumCellValues};
            }else if(TypeCellValues == "Impact"){
                BridgeFactors[CurrentVarCellValues] ={
                    Type: TypeCellValues,
                    Title:TitleCellValues,
                    Sum: SumCellValues};
            }

    }

    //对Bridge进行排序，将BasePT放在对象的第一位，Factors放在中间，TargetPT放在最后
    let sortedBridgeFactors = {};

    // 将 BasePT 放在第一位
    if (BridgeFactors.hasOwnProperty('BasePT')) {
        sortedBridgeFactors['BasePT'] = BridgeFactors['BasePT'];
    }

    // 将除 BasePT 和 TargetPT 之外的其他键按原本顺序添加
    for (let key in BridgeFactors) {
        if (key !== 'BasePT' && key !== 'TargetPT') {
            sortedBridgeFactors[key] = BridgeFactors[key];
        }
    }

    // 将 TargetPT 放在最后一位
    if (BridgeFactors.hasOwnProperty('TargetPT')) {
        sortedBridgeFactors['TargetPT'] = BridgeFactors['TargetPT'];
    }





    // 打印对象中的元素确认信息
    // for (let key in BridgeFactors) {    //第一层的Key
    //   if (BridgeFactors.hasOwnProperty(key)) {  //判断是否有Key
    //       console.log(`Key: ${key}`);
    //       let nestedObject = BridgeFactors[key]; //获取第一层的Key对应的对象
    //       for (let nestedKey in nestedObject) {  //第二层的对象的Key
    //           if (nestedObject.hasOwnProperty(nestedKey)) { 
    //               console.log(`${nestedKey}: ${nestedObject[nestedKey]}`); //获取第二场对应的Key的值
    //           }
    //       }
    //   }
    // }
    return sortedBridgeFactors;
  });
}

// 打印对象中的元素确认信息
async function PrintBridgeFactors() {
  await Excel.run(async (context) => {
      let Factors = await BridgeFactors();
      for (let key in Factors) {    //第一层的Key
        if (Factors.hasOwnProperty(key)) {  //判断是否有Key
            console.log(`Key: ${key}`);
            let nestedObject = Factors[key]; //获取第一层的Key对应的对象
            for (let nestedKey in nestedObject) {  //第二层的对象的Key
                if (nestedObject.hasOwnProperty(nestedKey)) { 
                    console.log(`${nestedKey}: ${nestedObject[nestedKey]}`); //获取第二场对应的Key的值
                }
            }
        }
      }
  });
}

// 创建waterfall工作表，生成Bridge数据，并返回相对应的单元格
async function BridgeCreate() {
  return await Excel.run(async (context) => {
      console.log("BridgeCreate1111")
      const workbook = context.workbook;
          // 检查是否存在同名的工作表
      let BridgeSheet = workbook.worksheets.getItemOrNullObject("Waterfall");
      await context.sync();

      if (BridgeSheet.isNullObject) {
        // 工作表不存在，创建新工作表
        BridgeSheet = context.workbook.worksheets.add("Waterfall");
        BridgeSheet.showGridlines = false; //隐藏工作表 'Waterfall' 的网格线
        await context.sync();
        console.log("创建了新工作表：" + "Waterfall");
      } else {

        BridgeSheet.delete();
        //await context.sync();
        
        BridgeSheet = context.workbook.worksheets.add("Waterfall");
        BridgeSheet.showGridlines = false; //隐藏工作表 'Waterfall' 的网格线
        await context.sync();
        console.log("已删除存在的工作表 Waterfall");
        console.log("创建了新工作表：" + "Waterfall");
        
      }

      let ColumnA = BridgeSheet.getRange("A:A");
      ColumnA.format.columnWidth = 10; // 设置 A 列宽度为 10
      //let BridgeSheet = context.workbook.worksheets.add("Waterfall");
      await context.sync();

      console.log("Waterfall onChanged event handler has been added.");
      console.log("BridgeCreate22222")
      let StartRange = BridgeSheet.getRange("B3");
      console.log("BridgeCreate22222233333")
      let Factors = await BridgeFactors(); //回传Bridge需要使用的factors对象
      console.log("BridgeCreate33333")
      let currentRange = StartRange;
      for (let key in Factors) {
        if (Factors.hasOwnProperty(key)) {
            // 将键值放入当前单元格
            currentRange.values = [[key]];
            
            // 将 sum 值放入右侧偏移一个单元格的位置
            currentRange.getOffsetRange(0, 1).values = [[Factors[key].Sum]];
            
            // 移动到下一行
            currentRange = currentRange.getOffsetRange(1, 0);
        }
      }
      currentRange = currentRange.getOffsetRange(-1, 1); // 循环结束后，回到两列的最右下角

      StartRange.load("address");
      currentRange.load("address");
      await context.sync();
      //获得BridgeRange
      let StartRangeAddress = getRangeDetails(StartRange.address);
      let CurrentRangeAddress = getRangeDetails(currentRange.address);
      let BridgeTopRow = StartRangeAddress.topRow;
      let BridgeBottomRow = CurrentRangeAddress.bottomRow;
      let BridgeLeftColumn = StartRangeAddress.leftColumn;
      let BridgeRightColumn = CurrentRangeAddress.rightColumn;
      let BridgeRange = BridgeSheet.getRange(`${BridgeLeftColumn}${BridgeTopRow}:${BridgeRightColumn}${BridgeBottomRow}`)

      BridgeRange.load("address");
      BridgeRange.format.autofitColumns(); // 自动调整宽度
      await context.sync();

      // BridgeRangeAddress = BridgeRange.address;


      //传递给TempVar 工作表，随时调用变量
      let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
      let BridgeRangeTitle = TempVarSheet.getRange("B5");
      let BridgeRangeVar = TempVarSheet.getRange("B6");
      BridgeRangeTitle.values = [["BridgeRange"]];
      BridgeRangeVar.values = [[`${BridgeRange.address}`]];

      await DoNotChangeCellWarning("Waterfall");
      return BridgeRange.address;

  });
}

// let BridgeDataFormatAddress = null; //Range地址全局变量，用来作为监控Bridge数据的变化，进而实时更新图形的标签等

// let BridgeRangeAddress = null;
//画出Bridge图形
async function DrawBridge() {
  await Excel.run(async (context) => {

    // isInitializing = false;
    // let BridgeRangeAddress = await BridgeCreate();  // 创建waterfall工作表，生成Bridge数据，并返回相对应的单元格，仅包含字段名和impact两列
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    let BridgeRangeVar = TempVarSheet.getRange("B6");
    BridgeRangeVar.load("values");
    await context.sync();

    let BridgeRangeAddress = BridgeRangeVar.values[0][0];

    console.log("BridgeRangeAddress is " + BridgeRangeAddress);
    // BridgeDataFormatAddress = BridgeRangeAddress; // 传递给全局函数
    
    // 获取名为 "Waterfall" 的工作表
    let sheet = context.workbook.worksheets.getItem("Waterfall");
    // 获取 Bridge 数据的范围
    let BridgeRange = sheet.getRange(BridgeRangeAddress);
    //let BridgeRange = sheet.getRange(BridgeRangeAddress);
    
    BridgeRange.load("address,values,rowCount,columnCount");
    await context.sync();

    let StartRange = BridgeRange.getCell(0, 0);
    let dataRange = StartRange.getOffsetRange(0, 2).getAbsoluteResizedRange(BridgeRange.rowCount, 4);


    //图形的数据范围
    let xAxisRange = StartRange.getAbsoluteResizedRange(BridgeRange.rowCount, 1); // 横轴标签范围
    let BlankRange = StartRange.getOffsetRange(0, 2).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let GreenRange = StartRange.getOffsetRange(0, 3).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let RedRange = StartRange.getOffsetRange(0, 4).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let AccRange = StartRange.getOffsetRange(0, 5).getAbsoluteResizedRange(BridgeRange.rowCount, 1); //辅助列
    let BridgeDataRange = StartRange.getOffsetRange(0, 1).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let BridgeFormats = StartRange.getOffsetRange(0,1).getAbsoluteResizedRange(BridgeRange.rowCount,5); //全部数据的范围，需要调整格式

    // 加载数据范围和横轴标签
    dataRange.load("address,values,rowCount,columnCount");
    xAxisRange.load("address,values,rowCount,columnCount");
    BlankRange.load("address,values,rowCount,columnCount");
    GreenRange.load("address,values,rowCount,columnCount");
    RedRange.load("address,values,rowCount,columnCount");
    AccRange.load("address,values,rowCount,columnCount");
    
    console.log("DrawBridge 0")

    //寻找BridgeDate sheet第一行带有Result的单元格
    let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
    let BridgeDataSheetRange = BridgeDataSheet.getUsedRange();
    let BridgeDataSheetFirstRow = BridgeDataSheetRange.getRow(0);

    BridgeDataSheetFirstRow.load("address");
    await context.sync();

    console.log("BridgeDataSheetFirstRow is " + BridgeDataSheetFirstRow.address);
    //await context.sync();
    console.log("DrawBridge 1");


    // // 找到result单元格
    // let ResultType = BridgeDataSheetFirstRow.find("Result", {
    //   completeMatch: true,
    //   matchCase: true,
    //   searchDirection: "Forward"
    // });
    // ResultType.load("address");
    // await context.sync();
    // console.log("DrawBridge 2")

    // //往下两行，获得Result数据单元格
    // let ResultCell = ResultType.getOffsetRange(2, 0);
    // ResultCell.load("numberFormat"); // 获得单元格的数据格式


    let VarTempResultFormat = TempVarSheet.getRange("B25");
    // await context.sync();
    // console.log("VarTempResultFormat here 1");
    VarTempResultFormat.load("numberFormat"); // 获得单元格的数据格式
    // await context.sync();
    // console.log("VarTempResultFormat here");

    // 将数据格式应用到 Bridge 数据范围
    BridgeFormats.copyFrom(
      VarTempResultFormat,
      Excel.RangeCopyType.formats // 只复制格式
    );
    // await context.sync();
    // console.log("BridgeFormats here");
    BridgeDataRange.load("address,values,rowCount,columnCount,text"); // 这个需要放在load了格式之后
    await context.sync();

    //console.log("ResultCell Formats is " + ResultCell.numberFormat[0][0]);
    console.log("dataRange is " + dataRange.address);
    console.log("xAxisRange is " + xAxisRange.address);
    console.log("BaseRange is " + BlankRange.address);
    console.log("GreenRange is " + GreenRange.address);
    console.log("RedRange is " + RedRange.address);
    console.log("AccRange is " + AccRange.address);

    //设置每个单元格的公式
    BlankRange.getCell(0, 0).formulas = [["=C3"]];
    console.log("DrawBridge 2.1");
    BlankRange.getCell(0, 0)
      .getOffsetRange(BridgeRange.rowCount - 1, 0)
      .copyFrom(BlankRange.getCell(0, 0));
    BlankRange.getCell(1, 0).formulas = [
      ["=IF(AND(G4<0,G3>0),G4,IF(AND(G4<=0,G3<=0,C4<=0),G4-C4,IF(AND(G4<0,G3<0,C4>0),G3+C4,SUM(C$3:C3)-F4)))"]
    ];
    BlankRange.getCell(0, 0)
      .getOffsetRange(1, 0)
      .getAbsoluteResizedRange(BridgeRange.rowCount - 2, 1)
      .copyFrom(BlankRange.getCell(1, 0));

    console.log("DrawBridge 3");
    AccRange.getCell(0, 0).formulas = [["=SUM($C$3:C3)"]];
    AccRange.getCell(0, 0)
      .getAbsoluteResizedRange(BridgeRange.rowCount - 1, 1)
      .copyFrom(AccRange.getCell(0, 0));
    AccRange.getCell(BridgeRange.rowCount - 1, 0).copyFrom(BlankRange.getCell(BridgeRange.rowCount - 1, 0), Excel.RangeCopyType.values);
    console.log("DrawBridge 4");
    GreenRange.getCell(0, 0).getOffsetRange(1, 0).formulas = [
      ["=IF(AND(G3<0,G4<0,C4>0),-C4,IF(AND(G3<0,G4>0,C4>0),C4+D4,IF(C4>0,C4,0)))"]
    ];
    GreenRange.getCell(0, 0)
      .getOffsetRange(1, 0)
      .getAbsoluteResizedRange(BridgeRange.rowCount - 2, 1)
      .copyFrom(GreenRange.getCell(0, 0).getOffsetRange(1, 0));
    RedRange.getCell(0, 0).getOffsetRange(1, 0).formulas = [
      ["=IF(AND(G3>0,G4<0,C4<0),D3,IF(AND(G3<=0,G4<=0,C4<=0),C4,IF(C4>0,0,-C4)))"]
    ];
    RedRange.getCell(0, 0)
      .getOffsetRange(1, 0)
      .getAbsoluteResizedRange(BridgeRange.rowCount - 2, 1)
      .copyFrom(RedRange.getCell(0, 0).getOffsetRange(1, 0));
    console.log("DrawBridge 5");

    //最后给辅助列设置灰色
    dataRange.format.fill.color = "#D3D3D3"; //将辅助列全部设置成灰色背景
    console.log("Setting Gray 1");
    let dataRangeTitle = dataRange.getRow(0).getOffsetRange(-1, 0);
    dataRangeTitle.merge();
    dataRangeTitle.format.fill.color = "#D3D3D3"; //将辅助列标题设置成灰色背景
    console.log("Setting Gray 2");
    dataRangeTitle.getCell(0, 0).values = [["辅助列"]];

    // 设置居中对齐
    dataRangeTitle.format.horizontalAlignment = "Center";
    dataRangeTitle.format.verticalAlignment = "Center";


    // 删除已有的图表，避免重复创建
    let charts = sheet.charts;
    charts.load("items/name");
    await context.sync();
    console.log("DrawBridge 6");

    // 检查并删除名为 "BridgeChart" 的图表（如果存在）
    for (let i = 0; i < charts.items.length; i++) {
      if (charts.items[i].name === "BridgeChart") {
        charts.items[i].delete();
        break;
      }
    }

    console.log("DrawBridge 7");
    // 插入组合图表（柱状图和折线图）
    let chart = sheet.charts.add(Excel.ChartType.columnStacked, dataRange, Excel.ChartSeriesBy.columns);
    chart.name = "BridgeChart"; // 设置图表名称，便于后续查找和删除
    
        // 隐藏图表图例
    chart.legend.visible = false;

    // 定义目标单元格位置（例如 D5）

    // 设置图表位置，左上角对应单元格
    chart.setPosition("I3");

    console.log("DrawBridge 8");

    // 设置图表的位置和大小
    // chart.top = 50;
    // chart.left = 50;
    // chart.width = 400;
    let labelCount = xAxisRange.rowCount; // 横坐标标签数量
    let labelWidth = 50; // 每个标签所需的宽度（像素），可以根据实际需求调整
    let minWidth = 400; // 图表最小宽度
    let maxWidth = 1000; // 图表最大宽度
    chart.width = Math.min(Math.max(labelCount * labelWidth, minWidth), maxWidth); // 根据标签数量调整宽度
    chart.height = 250;

    await context.sync();
    console.log("DrawBridge 9");
    // 设置横轴标签
    chart.axes.categoryAxis.setCategoryNames(xAxisRange);

    // 将轴标签位置设置为底部
    //chart.axes.valueAxis.position = "Automatic"; // 这里设置为Minimun 也只能在0轴的位置，不能是最低的负值下方
    let valueAxis = chart.axes.valueAxis;
    valueAxis.load("minimum");
    await context.sync();
    chart.axes.valueAxis.setPositionAt(valueAxis.minimum);

    // 获取图表的数据系列
    console.log("DrawBridge 10");
    const seriesD = chart.series.getItemAt(0); // Base列
    console.log("DrawBridge 10.1");
    const seriesE = chart.series.getItemAt(1); // 获取Green列的数据系列
    console.log("DrawBridge 10.2");
    const seriesF = chart.series.getItemAt(2); // 获取Red列的数据系列
    console.log("DrawBridge 10.3");
    const seriesLine = chart.series.getItemAt(3); // Bridge列

    seriesD.format.fill.clear();
    seriesD.format.line.clear();
    seriesE.format.fill.clear();
    seriesE.format.line.clear();
    seriesF.format.fill.clear();
    seriesF.format.line.clear();
    seriesLine.format.fill.clear();
    seriesLine.format.line.clear();

    seriesLine.chartType = Excel.ChartType.line; //插入Line
    //seriesLine.dataLabels.showValue = true;
    // 设置线条颜色为透明
    //seriesLine.format.line.color = "blue" ;
    seriesLine.format.line.lineStyle  = "None";
    console.log("DrawBridge 10.4");
    seriesLine.points.load("count"); //这一步必须
    console.log("DrawBridge 10.5");
    await context.sync();

    //设置线条的各种数据标签的颜色和位置等
    for (let i = 0; i < seriesLine.points.count; i++) {
      // let CurrentBridgeRange = BridgeDataRange.getCell(i, 0);
      // CurrentBridgeRange.load("values,text");
      // await context.sync();
      //seriesLine.points.getItemAt(i).dataLabel.text = String(CurrentBridgeRange.values[0][0]);

      console.log("BridgeDataRange.text[i][0] is " + BridgeDataRange.text[i][0]);
      if (i == 0 || i == seriesLine.points.count -1){
        // seriesLine.points.getItemAt(i).dataLabel.text = CurrentBridgeRange.text[0][0];
        seriesLine.points.getItemAt(i).dataLabel.text = BridgeDataRange.text[i][0];
        seriesLine.points.getItemAt(i).dataLabel.numberFormat = VarTempResultFormat.numberFormat[0][0]; //设置数据格式
        seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#0070C0"  // 蓝色
        // if(CurrentBridgeRange.values[0][0] >= 0){

          if(BridgeDataRange.values[i][0] >= 0){
          seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.top;
        }else{
          seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.bottom;
        }
        
      // }else if (CurrentBridgeRange.values[0][0] > 0) {
      }else if (BridgeDataRange.values[i][0] > 0) {
        // seriesLine.points.getItemAt(i).dataLabel.text = CurrentBridgeRange.text[0][0];
        seriesLine.points.getItemAt(i).dataLabel.text = BridgeDataRange.text[i][0];
        seriesLine.points.getItemAt(i).dataLabel.numberFormat = VarTempResultFormat.numberFormat[0][0]; //设置数据格式
        seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#00B050"  //绿色
        seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.top;

      // } else if (CurrentBridgeRange.values[0][0] < 0) {
      } else if (BridgeDataRange.values[i][0] < 0) {
        // seriesLine.points.getItemAt(i).dataLabel.text = CurrentBridgeRange.text[0][0];
        seriesLine.points.getItemAt(i).dataLabel.text = BridgeDataRange.text[i][0];
        seriesLine.points.getItemAt(i).dataLabel.numberFormat = VarTempResultFormat.numberFormat[0][0]; //设置数据格式
        seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#FF0000" //红色
        seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.bottom;

      } else {
        // seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#000000"  //黑色
        // seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.top;
      }
    }
    console.log("DrawBridge 10.6");

    seriesD.points.load("count");
    seriesE.points.load("count");
    seriesF.points.load("count");
    AccRange.load("address,values,rowCount,columnCount"); //这里需要再次加载，原来的加载还没有values
    await context.sync();
    console.log("DrawBridge 10.6.1");

    // 为 D 列的数据点设置填充颜色
    // seriesD.format.fill.clear();
    // seriesD.format.line.clear();
    for (let i = 0; i < seriesD.points.count; i++) {
      console.log("DrawBridge 10.6.2");
      // let BeforeAccRange = AccRange.values[i-1][0];      //getCell(i - 1, 0);
      let BeforeAccRange = (i > 0) ? AccRange.values[i-1][0] : null;   // 这里修改成为数组以后，有可能会越界，和原来getCell不会考虑越界不同
      let CurrentAccRange = AccRange.values[i][0];        //getCell(i, 0);
      // BeforeAccRange.load("values");
      // CurrentAccRange.load("values");

      // await context.sync();
      let point = seriesD.points.getItemAt(i);
      if(i>0){
        console.log("BeforeAccRange 0 is " + AccRange.values[i-1][0]);
        console.log("CurrentAccRange 0 is " + AccRange.values[i][0]);
      }
      if (i == 0 || i == seriesD.points.count - 1) {
        point.format.fill.setSolidColor("#0070C0"); // 设置为起始和终点颜色
        //seriesD.points.items[i].dataLabel.showValue = true;
        //seriesD.points.items[i].dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
      } else if (i > 0 && BeforeAccRange > 0 && CurrentAccRange < 0) {
        console.log("BeforeAccRange 1 is " + BeforeAccRange);
        console.log("CurrentAccRange 1 is " + CurrentAccRange);
        point.format.fill.setSolidColor("#FF0000"); // 设置为红色
      } else if (i > 0 && BeforeAccRange < 0 && CurrentAccRange > 0) {
        console.log("BeforeAccRange 2 is " + BeforeAccRange);
        console.log("CurrentAccRange 2 is " + CurrentAccRange);
        point.format.fill.setSolidColor("#00B050"); // 设置为绿色
      } else {
        point.format.fill.clear(); // 设置为无填充
      }
    }
    console.log("DrawBridge 10.7");
    //seriesE.dataLabels.showValue = true;
    //seriesE.dataLabels.position = Excel.ChartDataLabelPosition.insideBase ;

    // await context.sync();
    // 为E列数据点设置绿色
    // seriesE.format.fill.clear();
    // seriesE.format.line.clear();
    for (let i = 0; i < seriesE.points.items.length; i++) {
      let CurrentGreenRange = GreenRange.values[i][0];      //getCell(i, 0);
      // CurrentGreenRange.load("values");
      // await context.sync();
      let point = seriesE.points.getItemAt(i);

      point.format.fill.setSolidColor("#00B050");
      if (CurrentGreenRange !== 0) {
        //seriesE.points.items[i].dataLabel.showValue = true;
        //seriesE.points.items[i].dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
      }
    }
    console.log("DrawBridge 10.8");
    // 为F列数据点设置红色
    // seriesF.format.fill.clear();
    // seriesF.format.line.clear();
    for (let i = 0; i < seriesF.points.items.length; i++) {
      let CurrentRedRange = RedRange.values[i][0];       //getCell(i, 0);
      // CurrentRedRange.load("values");
      // await context.sync();F
      let point = seriesF.points.getItemAt(i);
      point.format.fill.setSolidColor("#FF0000");
      if (CurrentRedRange !== 0) {
        //seriesF.points.items[i].dataLabel.showValue = true;
        //seriesF.points.items[i].dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
      }
    }
    activateWaterfallSheet(); // 最后需要active waterfall 这个工作表
    console.log("DrawBridge 10.9");
    await context.sync();
  });
}


// 获取ProcessSum在Bridge Data Temp 中的地址//》》》》》》这里假设SumProcess是必须连续的，需要修改 */
// async function GetProcessSumRange() {
//   return await Excel.run(async (context) => {
//     console.log("GetProcessSumRange 1");
//     let sheet = context.workbook.worksheets.getItem("Bridge Data Temp");
//     let FirstRow = sheet.getUsedRange().getRow(0);
//     // let FirstCell = FirstRow.getCell(0, 0);
//     FirstRow.load("address,values,columnCount");
//     // FirstCell.load("address,values");
//     await context.sync();
//     console.log("GetProcessSumRange 2");
//     // console.log(FirstRow.address);
//     // console.log(FirstCell.address);

//     let StartIndex = null; //记录ProcessSum的起始位置
//     let NumIndex = 0; //记录ProcessSum的数量
//     for (let i = 0; i < FirstRow.columnCount; i++){
//       let CurrentCell = FirstRow.values[0][i]; //getOffsetRange(0, i);
//       // CurrentCell.load("address,values");
//       // await context.sync();

//       console.log("CurrentCell is " + CurrentCell);
//       if (CurrentCell == "ProcessSum") {
//         if (NumIndex == 0) {
//           StartIndex = i;
//         }
//         NumIndex++;
//       }
//     }
//     console.log("StartIndex is " + StartIndex);
//     console.log("NumIndex is " + NumIndex);

//     let ProcessSumRange = FirstRow.getOffsetRange(0, StartIndex).getAbsoluteResizedRange(1, NumIndex);
//     ProcessSumRange.load("address");
//     await context.sync();

//     console.log(ProcessSumRange.address);
//     return ProcessSumRange.address;
//   });
// }


// 获取ProcessSum在Bridge Data Temp 中的地址//
async function GetProcessSumRange() {
  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Bridge Data Temp");
    let FirstRow = sheet.getUsedRange().getRow(0);
    FirstRow.load("values, columnCount");
    await context.sync();

    let processSumIndices = [];
    for (let i = 0; i < FirstRow.columnCount; i++) {
      if (FirstRow.values[0][i] === "ProcessSum") {
        processSumIndices.push(i);
      }
    }

    let addressArray = [];
    processSumIndices.forEach(colIndex => {
      let cell = FirstRow.getCell(0, colIndex);
      cell.load("address");
      addressArray.push(cell);
    });

    await context.sync();
    return addressArray.map(cell => cell.address);
  });
}




// 循环运行RunProcess 获得所有需要解出的变量
async function ResolveLoop() {
  await Excel.run(async (context) => {
    console.log("ResolveLoop 1");
    let sheet = context.workbook.worksheets.getItem("Bridge Data Temp");
    let ProcessSumRangeCellAddress = await GetProcessSumRange();
    // console.log("ProcessSumRangeAddress is " + ProcessSumRangeAddress);
    // let ProcessSumRange = sheet.getRange(ProcessSumRangeAddress);
    // let ProcessSumStart = ProcessSumRange.getCell(0,0);
    // ProcessSumRange.load("address,values,rowCount,columnCount");
    // ProcessSumStart.load("address");
    // await context.sync();
    // let ProcessSumRangeCellAddress = await GetRangeAddress("Bridge Data Temp", ProcessSumRangeAddress);
    // let ProcessSumRangeCellAddress = ProcessSumRangeAddress;
    console.log("ProcessSumRangeCellAddress is ");
    console.log(ProcessSumRangeCellAddress);
    console.log("ResolveLoop 2");
    // console.log("ProcessSumRange is " + ProcessSumRange.address)

    for (let i = 0; i < ProcessSumRangeCellAddress.length;i++){
      // let ProcessSumCell = ProcessSumRange[0][i]; //ProcessSumStart.getOffsetRange(0,i);
      // ProcessSumCell.load("address,values");
      // await context.sync();

      // StrGblProcessSumCell = ProcessSumCell.address;
      let ProcessSumCell = ProcessSumRangeCellAddress[i];
      console.log("ProcessSumCell is " + ProcessSumCell)
      await runProcess(ProcessSumCell);

      // 如果processSumRange 的第一个单元格没有SumN 则会在runProcess 找不到，strGlobalFormulasCell 则会没有被赋值，如果运行下面的函数会出错。需要先判断
      // 》》》》》这里如果ProcessSumRange 单元格没有SumN， 这种情况的话，下面两个函数会重复运行两次，浪费时间，需要修改》》》》》
      if (strGlobalFormulasCell !== null) {  
        await GetFormulasAddress("Bridge Data Temp", strGlobalFormulasCell ,"Process", strGlbBaseLabelRange);
        await CopyFormulas();
      }

    }
    console.log("ResolveLoop End");
  });
}


// // 分解Bridge data 中result的公式，创建FormulasBreakdown，并在其中分解，并复制到BridgeData
// async function FormulaBreakDown() {
//   await Excel.run(async (context) => {
//     await copyAndModifySheet("Bridge Data", "FormulasBreakdown"); // 创建FormulasBreakdown工作表
//     let FormulaSheet = context.workbook.worksheets.getItem("FormulasBreakdown");
//     let FormulaRange = FormulaSheet.getUsedRange();
//     let FirstRow = FormulaRange.getRow(0); // 获取Type行，找到Result
//     FirstRow.load("address,values");
//     await context.sync();

//     console.log("FirstRow.address is " + FirstRow.address);

//     // 找到result单元格
//     let ResultType = FirstRow.find("Result", {
//       completeMatch: true,
//       matchCase: true,
//       searchDirection: "Forward"
//     });
//     ResultType.load("address");
//     //往下两行，获得Result对应的公式
//     let ResultCell = ResultType.getOffsetRange(2, 0);
//     ResultCell.load("address,formulas");
//     let ResultTitle =ResultType.getOffsetRange(1, 0);
//     ResultTitle.load("values");
//     await context.sync();
//     console.log("ResultTitle is " + ResultTitle.values[0][0]);

//     const newValue = ResultTitle.values[0][0];
//     if (!ArrVarPartsForPivotTable.includes(newValue)) { // 保证是唯一的变量
//       ArrVarPartsForPivotTable.push(newValue); // 给数据透视表筛选变量
//     }

//     // ArrVarPartsForPivotTable.push(ResultTitle.values[0][0]); //给数据透视表筛选变量

//     await FindNextFormulas(ResultCell.address); // 1>>>>>查找公式单元格中是否还有进一步引用的公式, 并最终反应在第一个单元格中, 这里已经去掉$固定符号

//     //下面需要重新load 一次，不然后面的代码不知道上一步已经改变了单元格内容。
//     ResultCell.load("address,formulas");
//     await context.sync();

//     let formula = await removeUnnecessaryParentheses(ResultCell.formulas[0][0].replace("=","")); ////取出掉公式里没有必要的括号
//     console.log("Remove ParenTheses is :" + formula);
//     console.log(formula);
//     ResultCell.formulas = [[formula]];
//     await context.sync();

//     //------------------------修改连续除法 Start----------------------
//     // let formula = range.formulas[0][0];
//     console.log("Replace formulas is " + formula);

//     // 阶段 1: 替换所有括号为 __replace__*
//     let innerMostParenthesesRegex = /\(([^()]*)\)/g; // match[0] 包含括号，match[1] 不包含括号，仅括号内内容
//     let tempFormulas = {};
//     let i = 0;
//     let match;

//     //--------------当有括号的时候需要循环下面的部分，并且不断的替换括号中的内容，然后不断的循环
//     while ((match = innerMostParenthesesRegex.exec(formula)) !== null) {
//       let innerExpr = match[0];
//       let key = `__replace__${i}`;
//       tempFormulas[key] = innerExpr;
//       formula = formula.replace(innerExpr, key);
//       innerMostParenthesesRegex.lastIndex = 0; // 重置正则索引
//       i++;
//       console.log("Dividend formulas is " + formula);
//       if(match[0].includes("/")){
//     //----------修正连续除号变乘法部分 start-----------------
//                 let formulaArray = await processFormulaObjforSplitDividend(formula); // 生成公式的分解对象数组
//                 console.log("formulaArray here is ");
//                 console.log(JSON.stringify(formulaArray, null, 2));
//                 let isConsecutiveDivisions = await checkConsecutiveDivisions(formulaArray); ////找到公式中连续除号的位置
//                 console.log("isConsecutiveDivisions here is ");
//                 console.log(JSON.stringify(isConsecutiveDivisions, null, 2));
//                 // 修改公式, 这里返回的是对象数组
//                 let modifiedFormula = modifyFormula(formulaArray, isConsecutiveDivisions.positions); //// 修改公式，插入括号和运算符替换
//                 console.log("modifiedFormula here is");
//                 console.log(JSON.stringify(modifiedFormula, null, 2));
//                 // 输出修改后的公式
//                 // let strModifiedFormula = formatFormula(modifiedFormula); // 将数组合并输出公式
//                 // console.log("modifiedFormula is " + strModifiedFormula);

//                 formula = formatFormula(modifiedFormula); // 将数组合并输出公式
//                 console.log("After change formula is " + formula);
//                 // ResultCell.formulas = [["=" + strModifiedFormula]];
//                 // await context.sync();
                
//     //----------修正连续除号变乘法部分 end-------------------
//       }
//     }
//     //-----------------当没有括号的时候，一样需要循环下面的部分
//     //----------修正连续除号变乘法部分 start-----------------
//     let formulaArray = await processFormulaObjforSplitDividend(formula); // 生成公式的分解对象数组
//     console.log("formulaArray here is ");
//     console.log(JSON.stringify(formulaArray, null, 2));
//     let isConsecutiveDivisions = await checkConsecutiveDivisions(formulaArray); ////找到公式中连续除号的位置
//     console.log("isConsecutiveDivisions here is ");
//     console.log(JSON.stringify(isConsecutiveDivisions, null, 2));
//     // 修改公式, 这里返回的是对象数组
//     let modifiedFormula = modifyFormula(formulaArray, isConsecutiveDivisions.positions); //// 修改公式，插入括号和运算符替换
//     console.log("modifiedFormula here is");
//     console.log(JSON.stringify(modifiedFormula, null, 2));
//     // 输出修改后的公式
//     // let strModifiedFormula = formatFormula(modifiedFormula); // 将数组合并输出公式
//     // console.log("modifiedFormula is " + strModifiedFormula);

//     formula = formatFormula(modifiedFormula); // 将数组合并输出公式
//     console.log("After change formula is " + formula);
//     // ResultCell.formulas = [["=" + strModifiedFormula]];
//     // await context.sync();
    
//     //----------修正连续除号变乘法部分 end-------------------


//     // 阶段 2: 还原 __replace__* 为原始表达式
//     // 获取所有键并按索引从高到低排序（确保外层优先替换）
//     const keys = Object.keys(tempFormulas).sort((a, b) => {
//       const numA = parseInt(a.split("__replace__")[1]);
//       const numB = parseInt(b.split("__replace__")[1]);
//       return numB - numA; // 反向排序
//     });

//     // 遍历每个键并替换回原始表达式
//     keys.forEach((key) => {
//       const regex = new RegExp(key.replace(/\$/g, "\\$"), "g"); // 转义特殊字符
//       formula = formula.replace(regex, tempFormulas[key]);
//     });

//     console.log("Restored formula: " + formula);
//     console.log("Temp formulas:", tempFormulas);

//     // 可选：将还原后的公式写回单元格
//     ResultCell.formulas = [[formula]];
//     await context.sync();

//     //------------------------修改连续除法 End----------------------
//     console.log("before processFormulaObj formula is" + formula);

//     await processFormulaObj(ResultCell.address); // 不返回任何的值，函数修改成完全为了数据透视表筛选数据用


//     ///////////////////-----考虑有括号的情况下，判断四则运算的结果是否是SumN 还是Additive--------

//     // while ((matchForAdditive[1] = innerMostParenthesesRegex.exec(formula)) !== null) {
//     //     console.log("matchForAdditive[1] is " + matchForAdditive[1]); // matchForAdditive[1] 不包括括号

//     //     // 匹配公式中的所有单元格引用和括号内的表达式
//     //       // let cellReferences = formula.match(/([A-Za-z]+\d+|\b\d+\b)/g); //考虑单纯的数字情况
//     //       // let parts = []; // 用来保存公式的每个部分

//     //     // // 分割公式，保留运算符和括号
//     //     //   let formulaParts = formula.split(/([+\-*/()])/g).filter(part => part.trim() !== "");

//     //     // 现在 formula == "A3+B3*C3" 这样的

//     //     // 1. 分离出运算符和操作数。这里做一个非常粗糙的 split，然后再拼接运算符
//     //     //    也可以用更复杂的正则或解析器。这里只是演示思路。
//     //     //    以 + - * / 作为分隔符，并保留分隔符（用于后续组装）。
//     //     //    例如 "A3+B3*C3" -> ["A3", "+", "B3", "*", "C3"]。
//     //     const tokenPattern = /([+\-*/()])/; 
//     //     // split 之后会把分隔符也拆出来
//     //     let tokens = matchForAdditive[1].split(tokenPattern).map(t => t.trim()).filter(Boolean);
//     //     // tokens => ["A3", "+", "B3", "*", "C3"]

//     //     // 2. 对每个单元格引用（例如 "A3", "B3"）进行类型替换
//     //     //    规则：如果是单元格引用 Xn，则到 "X1" 去读取 Additive / SumN。
//     //     //    这里假设只有一个工作表，不考虑绝对/相对/跨表等复杂情况。
//     //     for (let i = 0; i < tokens.length; i++) {
//     //       const t = tokens[i];
//     //       // 如果是运算符 (+ - * /)，就直接跳过
//     //       if (t === "+" || t === "-" || t === "*" || t === "/") {
//     //         continue;
//     //       }

//     //       // 否则，尝试判断是不是一个像 "A3" 这样的引用
//     //       // 简单用正则：([A-Za-z]+)(\d+)
//     //       const CellAddress = t.match(/^([A-Za-z]+)(\d+)$/);
//     //       if (CellAddress) {
//     //         const colLetters = CellAddress[1]; // "A" / "B" / "C"
//     //         // const rowNumber = match[2]; // "3"

//     //         // 构造我们要去读取的 “类型定义单元格”: 比如 A3 -> A1
//     //         const typeCellAddress = colLetters + "1";

//     //         // 从这个单元格中读出 "Additive" 或 "SumN"
//     //         const typeCell = sheet.getRange(typeCellAddress);
//     //         typeCell.load("values");
//     //         await context.sync();
//     //         const cellValue = typeCell.values[0][0];
//     //         // 假设用户确保在 A1, B1, C1 等位置写的就是 "Additive" 或 "SumN"

//     //         // 替换当前 token 为这个类型字符串
//     //         tokens[i] = cellValue;
//     //       } else { //***这里加入判断识别单纯的数字如10，直接作为可以相加的数？ */
//     //         // 如果不是运算符，也不是单元格引用，可根据需求做处理

            
//     //         // 这里直接不处理
//     //       }
//     //     }

//     //     // 3. 组装成一个最终字符串（用空格隔开，便于后续 determineExpressionType() 解析）
//     //     //    例如 ["Additive", "+", "SumN", "*", "Additive"] -> "Additive + SumN * Additive"
//     //     const expression = tokens.join(" ");

//     // }

//      ////////////////////-----考虑有括号的情况下，判断四则运算的结果是否是SumN 还是Additive--End----

//     await reorderFormula(ResultCell.address); // 

//     console.log("ResultType is " + ResultType.address);
//     console.log("ResultCell is " + ResultCell.address);


//     await processFormula(ResultCell.address); //2>>>>>>>>>>> 对公式里的运算符和优先级，从左到右加上括号
//     await SplitFormula(ResultCell.address); //3
//   });
// }


//////////////////////对四则运算进行是否是SumN进行判断/////////////////

// /**
// * @param {string} expression - 例如 "Additive + SumN * Additive"
// * @returns {string} - 返回 "Additive" 或 "SumN"
// **/
// function determineExpressionType(expression) {
//  try {
//    // 1. 将表达式分割成 Token
//    // 假设用户输入的表达式用空格隔开：Additive + SumN * Additive
//    // 得到 ["Additive", "+", "SumN", "*", "Additive"]
//    const tokens = expression.split(/\s+/).filter(Boolean);

//    // 2. 先处理 * / （乘除）—— 因为它们优先级较高
//    handleMultiplyDivide(tokens);

//    // 3. 再处理 + - （加减）
//    handleAddSubtract(tokens);

//    // 4. 处理完成后，tokens 应该只剩一个元素——最终的类型
//    if (tokens.length === 1) {
//      return tokens[0];
//    } else {
//      // 如果还剩下不止一个，说明表达式格式不符合预期
//      return "表达式有误，无法判定";
//    }
//  } catch (err) {
//    console.error(err);
//    return "Error";
//  }
// }

// /**
// * 处理 tokens 中的乘除运算
// * @param {string[]} tokens
// */
// function handleMultiplyDivide(tokens) {
//  let i = 0;
//  while (i < tokens.length) {
//    const token = tokens[i];
//    if (token === "*" || token === "/") {
//      // 取出左右操作数
//      const leftType = tokens[i - 1];
//      const rightType = tokens[i + 1];
//      // 合并为新的结果
//      const newType = combineTypes(leftType, rightType, token);
//      // 将 i-1, i, i+1 三个元素替换成一个新的类型
//      tokens.splice(i - 1, 3, newType);
//      // 回退索引，继续往前处理
//      i = i - 1;
//    } else {
//      i++;
//    }
//  }
// }

// /**
// * 处理 tokens 中的加减运算
// * @param {string[]} tokens
// */
// function handleAddSubtract(tokens) {
//  let i = 0;
//  while (i < tokens.length) {
//    const token = tokens[i];
//    if (token === "+" || token === "-") {
//      // 取出左右操作数
//      const leftType = tokens[i - 1];
//      const rightType = tokens[i + 1];
//      // 合并为新的结果
//      const newType = combineTypes(leftType, rightType, token);
//      // 将 i-1, i, i+1 三个元素替换成一个新的类型
//      tokens.splice(i - 1, 3, newType);
//      i = i - 1;
//    } else {
//      i++;
//    }
//  }
// }

// /**
// * 根据自定义规则，将两种类型通过指定运算符，得到结果类型
// * @param {string} type1 - "Additive" 或 "SumN"
// * @param {string} type2 - "Additive" 或 "SumN"
// * @param {string} operator - "+", "-", "*", "/"
// * @returns {string} - "Additive" 或 "SumN"
// */
// function combineTypes(type1, type2, operator) {
//  // 加减规则（不变）
//  if (operator === "+" || operator === "-") {
//    // 只有两个都是 Additive 才返回 Additive，否则返回 SumN
//    if (type1 === "Additive" && type2 === "Additive") {
//      return "Additive";
//    } else {
//      return "SumN";
//    }
//  }

//  // 乘法新规则
//  if (operator === "*") {
//    // - Additive * Additive → SumN
//    // - Additive * SumN → Additive
//    // - SumN * Additive → Additive
//    // - SumN * SumN → SumN
//    if (type1 === "Additive" && type2 === "Additive") {
//      return "SumN";
//    } else if (type1 === "Additive" && type2 === "SumN") {
//      return "Additive";
//    } else if (type1 === "SumN" && type2 === "Additive") {
//      return "Additive";
//    } else {
//      return "SumN";
//    }
//  }

//  // 除法新规则
//  if (operator === "/") {
//    // - Additive / Additive → SumN
//    // - Additive / SumN → Additive
//    // - SumN / Additive → SumN
//    // - SumN / SumN → SumN
//    if (type1 === "Additive" && type2 === "Additive") {
//      return "SumN";
//    } else if (type1 === "Additive" && type2 === "SumN") {
//      return "Additive";
//    } else if (type1 === "SumN" && type2 === "Additive") {
//      return "SumN";
//    } else {
//      return "SumN";
//    }
//  }

//  // 出现未知运算符时，默认返回 SumN
//  return "SumN";
// }




//////////////////////对四则运算进行是否是SumN进行判断///End//////////////

// 1>>>>>查找公式单元格中是否还有进一步引用的公式, 并最终反应在第一个单元格中
async function FindNextFormulas(FormulaRangeAddress) {  
  return await Excel.run(async (context) => {
    let BridgeDataSheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    let FormulaRange = BridgeDataSheet.getRange(FormulaRangeAddress);
    FormulaRange.load("address,values,formulas");

    await context.sync();
    console.log(FormulaRange.address, FormulaRange.values[0][0], FormulaRange.formulas[0][0]);

    let CellFormula = FormulaRange.formulas[0][0].replace(/\$/g, ""); //替换$等在公式里的固定符号
    FormulaRange.formulas = [[CellFormula]]; // 这里要赋值回去，否则影响后面的取数
    await context.sync();
    console.log("Formulas is " + CellFormula);

    // let CellReferences = CellFormula.match(/([A-Z]+[0-9]+)/g);
      let CellReferences = CellFormula.match(/([A-Za-z]+\d+)/g);
    console.log(CellReferences);

    //循环查找公式中是否存在进一步的公式
    for (let i = 0; i < CellReferences.length; i++) {
      let CellAddress = CellReferences[i];
      let Cell = BridgeDataSheet.getRange(CellAddress);
      Cell.load("address,values,formulas");
      await context.sync();
      
      // let CellFormulas = Cell.formulas[0][0];
      // CellFormulas = CellFormulas.replace("=","");

      // console.log("Cell.formulas[0][0]" + CellFormulas);

      if (Cell.values[0][0] != Cell.formulas[0][0]) {
        // let CellFormulas = Cell.formulas[0][0];
        // CellFormulas = CellFormulas.replace("=","");

        // console.log("Cell.formulas[0][0]" + CellFormulas);

        // let tokenPattern = /([+\-*/()])/; 
        // let tokens = CellFormulas.split(tokenPattern).map(t => t.trim()).filter(Boolean); //分解除公式中的所有元素
        // console.log("FindNextFormulas tokens is");
        // console.log(tokens);

        // //将公式中的（J3-K3) 转换成（ARR-ARR2)这样的变量名，去和checkType2Var全局变量中的数组做对比

        // for (let i = 0; i < tokens.length; i++) {
        //   let part = tokens[i];
        //   let isOperator = /[+\-*/()]/.test(part);
        //   if (!isOperator) {
        //     // 在 FormulaTokens 中查找匹配的 Token
        //     let NextVar = BridgeDataSheet.getRange(part);
        //     let NextTitle = NextVar.getOffsetRange(-1,0); //王上一层获得J3的标题ARR
        //     NextTitle.load("values");
        //     await context.sync();

        //     // 将J3 转换乘ARR 
        //     tokens[i] = NextTitle.values[0][0];

        //   }
        // }

        // CellFormulas = tokens.join(""); //现在应该已经从（J3-K3) 转换成（ARR-ARR2)这样的变量名
        // console.log("after change, CellFormulas is " + CellFormulas);

        

        // if(checkType2Var.includes(CellFormulas)){
        //   console.log("checkType2Var for nextformula")
        //   continue;   // 如果这个变量是因为SumN+sumN作为分母新加入的变量，则不需要进一步迭代，不需要替换下一层公式，不然会重复生成同样的变量
        // }

        await FindNextFormulas(CellAddress); // 嵌套循环 不断查找, 这里必须加入await, 不然不等这一步完成，顺序不对
        Cell.load("formulas"); // 这里需要重新load一遍，因为上一步循环嵌套已经更新了公式，不然没法反应都最终的公式中
        await context.sync();

        //将
        let modifiedFormula = `(${Cell.formulas[0][0].substring(1)})`; // 最外层加上括号 / .substring(1)：从该公式字符串的第二个字符开始提取子字符串（即去掉第一个字符）。
        console.log("modifiedFormula is " + modifiedFormula);

        let Newformula = FormulaRange.formulas[0][0].replace(CellReferences[i], modifiedFormula);
        FormulaRange.formulas = [[Newformula]];
        await context.sync();

        console.log("FormulaRange.formulas[0][0] is " + FormulaRange.formulas[0][0]);
        //CellReferences[i] = modifiedFormula; // 找到下一层公式后，修改替换原来公式
        console.log(`${CellReferences[i]} is ` + CellReferences[i]);
      }
    }

    console.log("CellReferences is " + CellReferences);
    console.log("Loop");

    await context.sync();
    return FormulaRange.formulas;
  });
}

//2>>>>>>>>>>> 对公式里的运算符和优先级，从左到右加上括号
async function processFormula(FormulaAddress) {
  await Excel.run(async (context) => {

      let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
      let formulaCell = sheet.getRange(FormulaAddress); // 假设公式在E1
      formulaCell.load("formulas");
      await context.sync();

      let formula = formulaCell.formulas[0][0].replace("=","");
      console.log("processFormula is " + formula)

      let keyFormula = {}; // 存储替换的公式和键值对
      let keyCounter = 1; // 用于生成键的计数器

      // 辅助函数：生成唯一的键
      function generateKey() {
        return `_M${keyCounter++}_`;
      }

      // 1. 处理公式中的括号
      while (/\([^()]*\)/.test(formula)) {
        formula = formula.replace(/\([^()]*\)/, (match) => {
          let innerExpr = match.slice(1, -1); // 去掉括号
          let key = handleInnerExpression(innerExpr); // 处理括号内的表达式并返回键
          return key;
          
        });
      }

      // 2. 处理没有括号的公式
      formula = handleInnerExpression(formula);

      // 3. 逐步恢复公式，从最后一个键开始替换
      let keys = Object.keys(keyFormula).reverse(); // 获取键的数组，并反转顺序

      for (let key of keys) {
        formula = formula.replace(key, keyFormula[key]);
      }

      formulaCell.formulas = [["=" + formula]]
      console.log("processFormula end is " + formula)
      return formula;

      // 辅助函数：处理表达式，添加括号
      function handleInnerExpression(innerExpr) {
        // 找到表达式中的所有运算符（+ - * /）
        let operators = innerExpr.match(/[+\-*/]/g);

        // 如果表达式中没有运算符，直接返回原始表达式
        if (!operators) {
          return innerExpr;
        }

        // 如果表达式中只有一个运算符
        if (operators.length === 1) {
          // 为表达式添加括号，并存储到 keyFormula 对象中，返回键
          let key = generateKey();
          keyFormula[key] = `(${innerExpr})`;
          return key;
        } else {
          // 如果表达式中有多个运算符，优先处理乘法和除法
          while (/[*\/]/.test(innerExpr)) {
            innerExpr = innerExpr.replace(/[\w\d.]+[*\/][\w\d.]+/, (match) => {
              // 为找到的第一个乘法或除法表达式添加括号
              let key = generateKey();
              keyFormula[key] = `(${match})`;
              return key; // 用键替换表达式中相应的部分
            });
          }

          // 处理剩下的加法和减法
          if (/[+\-]/.test(innerExpr)) {
            // 如果剩余部分中只有加法和减法，则将其用括号括起来，并存储为键值对
            let key = generateKey();
            keyFormula[key] = `(${innerExpr})`;
            innerExpr = key; // 用键替换表达式中相应的部分
          }

          return innerExpr; // 返回最终的表达式或键
        }
      }
  });    
}


//3>>>>>>>>分解公式里带括号的，不断扩大，并在右方单元格不断扩展放置结果， 并在Bridge Data中复制同样的公式列
async function SplitFormula(FormulaAddress) {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
    let formulaCell = sheet.getRange(FormulaAddress);
    let UsedRange = sheet.getUsedRange();

    formulaCell.load("formulas,address");
    UsedRange.load("address");
    await context.sync();

    let UsedRightRange = sheet.getRange(
      `${getRangeDetails(UsedRange.address).rightColumn}${getRangeDetails(formulaCell.address).bottomRow}`
    );

    let formula = `${formulaCell.formulas[0][0].replace("=", "")}`; //不用再最外层加上括号，因为已经在addSplit里加入了最外层括号
    console.log("formula is " + formula);

    var regex = /\([^\(\)]*\)/g; // 匹配最内层的括号
    var match;

    let BracketNo = 1; //用于计数有多少括号的先后排序
    // let Bracket = {};
    while ((match = regex.exec(formula)) !== null) {
      // 当前匹配的括号内容
      let matchedPart = match[0];
      console.log("matchedPart is " + matchedPart);
      //Bracket[`Bracket${BracketNo}`] = matchedPart; //

      let BracketCell = UsedRightRange.getOffsetRange(0, BracketNo); //每次循环往右移动一格
      BracketCell.load("address,formulas");
      await context.sync();

      // 使用最新地址替换当前匹配部分
      // 先判断之前是否已经有了相同的公式被分解在之前的单元格里，例如(Revenue - Cost)/ Revenue, revenue 部分已经在之前分解，分母不能再重复
      let PreMatch = 0; //用来判断是否需要跳出while的剩余代码
      for(let i = 1;i <BracketNo;i++){
        let CurrentCell = UsedRightRange.getOffsetRange(0,i); // 循环到目前位置所有的分解单元格
        CurrentCell.load("address, values, formulas");
        await context.sync();

        if(CurrentCell.formulas[0][0].replace("=","") == matchedPart ){
          formula = formula.replace(matchedPart, CurrentCell.address.split("!")[1]); //替换使用之前已经被分解的单元格
          regex.lastIndex = 0; // 循环的过程中，搜索的位置会不断往后移动，需要重置
          //BracketNo++;
          PreMatch = 1;
          break;// 找到后跳出for循环
        };

      }
      
      if(PreMatch ==1){
        PreMatch =0;
        continue; // 不执行while循环剩下的代码
      }

      formula = formula.replace(matchedPart, BracketCell.address.split("!")[1]);
      console.log("formula is " + formula);
      BracketCell.formulas = [[`=${matchedPart}`]]; //在最新的地址写入目前匹配的公式

      regex.lastIndex = 0; // 循环的过程中，搜索的位置会不断往后移动，需要重置
      BracketNo++;
    }

    let CurrentRange = UsedRightRange.getOffsetRange(0, 1).getAbsoluteResizedRange(1, BracketNo - 1);
    CurrentRange.load("address");
    await context.sync();

    BracketNo = await DeleteNoUseProcessSumRange(CurrentRange.address,BracketNo); //3.1>>>>>删除掉对求解SumN没有作用的单元格，返回减少后的BracketNo

 
    //拷贝到Bridge Data对应的单元格中
    let SplitFormulaRange = UsedRightRange.getOffsetRange(0, 1).getAbsoluteResizedRange(1, BracketNo - 1);
    SplitFormulaRange.load("address");
    await context.sync();

    let TypRange = SplitFormulaRange.getOffsetRange(-2,0); //FormulasBreakdown 中的Type

    let BridgeDataSplitRange = BridgeDataSheet.getRange(SplitFormulaRange.address.split("!")[1]);
    BridgeDataSplitRange.copyFrom(SplitFormulaRange);
    let BridgeUsedRange = BridgeDataSheet.getUsedRange();
    BridgeDataSplitRange.load("address, values, formulas");
    BridgeUsedRange.load("address");

    let SplitTypeRange = BridgeDataSplitRange.getOffsetRange(-2, 0); //获得最上一行，放入ProcessSum
    SplitTypeRange.copyFrom(TypRange); // 拷贝Type


    // SplitTypeRange.load("rowCount, columnCount");
    // await context.sync();

    // // 遍历范围中的每个单元格，并设置值为 "ProcessSum"
    // for (let i = 0; i < SplitTypeRange.rowCount; i++) {
    //   for (let j = 0; j < SplitTypeRange.columnCount; j++) {
    //     let cell = SplitTypeRange.getCell(i, j);
    //     cell.values = [["ProcessSum"]];
    //   }
    // }

    await context.sync();

    console.log("BridgeDataSplitRange is " + BridgeDataSplitRange.address);
    console.log("BridgeUsedRange is " + BridgeUsedRange.address);

    let BridgeSplitBottomRow = getRangeDetails(BridgeUsedRange.address).bottomRow;
    let BridgeDataSplitRangeAddress = getRangeDetails(BridgeDataSplitRange.address);
    let BridgeSplitTopRow = BridgeDataSplitRangeAddress.topRow;
    let BridgeSplitLeftColumn = BridgeDataSplitRangeAddress.leftColumn;
    let BridgeSplitRightColumn = BridgeDataSplitRangeAddress.rightColumn;
    let BridgeFullSplitRange = BridgeDataSheet.getRange(
      `${BridgeSplitLeftColumn}${BridgeSplitTopRow}:${BridgeSplitRightColumn}${BridgeSplitBottomRow}`
    );
    BridgeFullSplitRange.load("address");
    await context.sync();

    console.log(BridgeFullSplitRange.address);

    console.log(BridgeSplitTopRow);
    console.log(BridgeSplitBottomRow);
    console.log(BridgeSplitLeftColumn);
    console.log(BridgeSplitRightColumn);
    BridgeFullSplitRange.copyFrom(BridgeDataSplitRange);
    await context.sync();

    //复制到标题
    let SplitTitleRange = BridgeDataSplitRange.getOffsetRange(-1, 0);
    SplitTitleRange.copyFrom(BridgeDataSplitRange);
    SplitTitleRange.load("address,formulas,values");

    let BreakDownTitle = SplitFormulaRange.getOffsetRange(-1,0); // 在breakdown sheet 中也还原变量的标题
    BreakDownTitle.copyFrom(SplitFormulaRange);
    BreakDownTitle.load("address,formulas,values");

    await context.sync();

    console.log("SplitTitleRange is " + SplitTitleRange.values[0][0]);
    console.log("BreakDownTitle is " + BreakDownTitle.values[0][0]);

    await replaceReferencesInRange("Bridge Data", SplitTitleRange.address);
    await replaceReferencesInRange("FormulasBreakdown", BreakDownTitle.address); 

    ////////////-----------------添加给数据透视表的数据-----------------

    let VarRange = UsedRightRange.getOffsetRange(-1, 1).getAbsoluteResizedRange(1, BracketNo - 1); // 网上一格，获得变量的名称输入给数据透视表
    VarRange.load("address,values")

    await context.sync();
    console.log("VarRange Address is " + VarRange.address);
    console.log("VarRange Values is");
    console.log(JSON.stringify(VarRange.values, null, 2));

    ArrVarPartsForPivotTable = ArrVarPartsForPivotTable.concat(VarRange.values[0]); // 追加ProcessSum的数据给数据透视表
    console.log("ArrVarPartsForPivotTable is ");
    console.log(JSON.stringify(ArrVarPartsForPivotTable, null, 2));
    ////////////-----------------添加给数据透视表的数据----END-------------


  });
}

//3.1>>>>>删除掉对求解SumN没有作用的单元格，返回减少后的BracketNo
async function DeleteNoUseProcessSumRange(rangeAddress, BracketNo) {
  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    let range = sheet.getRange(rangeAddress); // 根据地址获取Range对象
    let SolveVar = []; // 定义一个数组SolveVar

    console.log("delete 1");
    console.log("rangeAddress is " + rangeAddress);

    // 加载Range中的公式和地址
    range.load(["formulas", "address", "columnCount"]);
    await context.sync();  // 确保属性已经加载

    // 3. 从左到右循环这个Range的每一个单元格
    for (let i = 0; i < range.columnCount; i++) {
      let cell = range.getCell(0, i);
      cell.load("address,formulas,values");
      await context.sync();

      let formula = cell.formulas[0][0];
      let address = cell.address.split("!")[1];

      // 3.1 解析当前单元格X里的公式，匹配出其中变量对应的单元格
      let matches = formula.match(/[A-Za-z]+\d+/g) || [];
      let cellObj = { Address: address, NonAdditive: false, reference: false };
      SolveVar.push(cellObj);  //这一步无论分解出来的变量是否是SumN，都会存到SolveVar中

      // 3.1.1 循环判断进一步分解出来的，每一个匹配出来的变量
      for (let match of matches) {
        let refCell = sheet.getRange(match);
        let cellAbove = refCell.getOffsetRange(-2, 0); // 向上移动两行的单元格
        let cellTitle = refCell.getOffsetRange(-1, 0); // 向上移动一行的单元格

        cellAbove.load("values");
        cellTitle.load("values");
        await context.sync();  // 确保属性已经加载

        let titleValue = cellTitle.values[0][0];
        let isNonAdditive = cellAbove.values[0][0] === "SumN";

        //--标签0121---
        if (isNonAdditive) {
          // 3.1.1.1 如果SolveVar数组中没有这个Title
          let existingTitle = SolveVar.find(item => item.Title === titleValue);
          if (!existingTitle) {
            cellObj.NonAdditive = true; // 这一步把目标单元格cellObj标记为包含Non-Additive的类型
            console.log("cellObj.NonAdditive is " + cellObj.Address)
            SolveVar.push({ Address: match, Title: titleValue, NonAdditive: false, reference: false });

          } else {
            // 3.1.1.2 如果SolveVar数组中已经存在同样的Title
            //如果已经存在同样的字段存在SolveVar中，则目标单元格cellObj就不标记为Non-Additive类型
            cellObj.NonAdditive = false;
          }
        }
      }
      console.log("delete 2");
      console.log(JSON.stringify(SolveVar, null, 2));

      // 3.2 判断当前单元格X的SumN的键值，如果是true，则公式里所有的单元格的对象的reference都为true
      //只有cellObj标记为Non-Additive的类型，才能往下进行
      if (cellObj.NonAdditive) {
        console.log("cellObj.NonAdditive is");
        console.log(cellObj.NonAdditive);
        for (let match of matches) {
          console.log("match is " + match)
          let refObj = SolveVar.find(item => item.Address === match);
          //console.log("refObj address is " + refObj.Address)
          if (refObj) {
            refObj.reference = true;
            console.log("refObj with reference address is " + refObj.Address)

            //在被引用的单元格里继续迭代深入看是否有进一步引用的公式，找到单元格并将reference 改成true***这里会不会有引用单元格还没有生成对象的情况？
            //经过分析，在0121标签的地方，已经把所有需要的引用单元格放进了SolveVar中。
            // 但是如果在初始的SplitRange中，如果是一个A Nov-additive引用B Nov-additive，A和B都在0121中被生成了对象。
            // 可是如果B Nov-additive继续在SplitRange中继续引用单元格C，则C应该还没有被生成，需要在下一步生成

            async function ReferenceLoop(RangeAddress) {
              return await Excel.run(async (context) => {
                let Sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
                let Range = Sheet.getRange(RangeAddress);
                Range.load("address,formulas,values");
                await context.sync();

                let formula = Range.formulas[0][0];
                let value = range.values[0][0];
                // let matches = formula.match(/[A-Z]+\d+/g) || [];
                let matches = formula.match(/[A-Za-z]+\d+/g) || [];

                for (let match of matches) {
                  let refCell = Sheet.getRange(match);
                  let cellAbove = refCell.getOffsetRange(-2, 0); // 向上移动两行的单元格
                  let cellTitle = refCell.getOffsetRange(-1, 0); // 向上移动一行的单元格

                  cellAbove.load("values");
                  cellTitle.load("values");
                  await context.sync();  // 确保属性已经加载

                  let titleValue = cellTitle.values[0][0];
                  //let isNonAdditive = cellAbove.values[0][0] === "SumN";

                  for (let match of matches) {
                    let refObj = SolveVar.find(item => item.Address === match);
                    console.log("refObj address2 is " + refObj.Address)
                    if (refObj) {
                      refObj.reference = true;
                      console.log("refObj with reference2 address is " + refObj.Address)
                      //一直迭代到没有公式的单元格
                      if (!(formula === value)) {
                        ReferenceLoop(refObj.Address); //自身进一步迭代 ***** 是否会迭代到SolveVar 数组中还不存在的情况？下面用else部分解决
                      }
                    } else {
                      //如果SolveVar中还没有生成对象，则现在生成，因为这里已经是因为SumN不断迭代需要形成的，reference赋值成true
                      //else这一步还没有测试，需要用
                      SolveVar.push({ Address: match, Title: titleValue, NonAdditive: false, reference: true });
                      //一直迭代到没有公式的单元格
                      if (!(formula === value)) {
                        ReferenceLoop(refObj.Address);
                      }

                    }
                  }
                }
              });
            }

            ReferenceLoop(refObj.Address); // 调用


          }
        }
      }
    }
    console.log("delete 3");
    // 4. 循环 Range A中的所有单元格，执行删除操作// 改成修第一行的标题为null
    for (let i = range.columnCount - 1; i >= 0; i--) {
      let cell = range.getCell(0, i);
      cell.load("address,formulas,values");
      await context.sync();

      let address = cell.address.split("!")[1];

      let cellObj = SolveVar.find(item => item.Address === address);
      console.log("CellObj is " + cell.address)
      if (cellObj && !cellObj.NonAdditive && !cellObj.reference) {

        console.log("Delete Address is " + cellObj.Address);
        //let DeleteCOlumn = getRangeDetails(cell.address).leftColumn
        // cell.delete(Excel.DeleteShiftDirection.left);
        //BracketNo--;
        cell.getOffsetRange(-2, 0).values = [["Null"]];
        //sheet.getRangeByIndexes(0, i, sheet.getUsedRange().rowCount, 1).delete(Excel.DeleteShiftDirection.left);
      } else {

        cell.getOffsetRange(-2, 0).values = [["ProcessSum"]];

      }
    }

    await context.sync();
    return BracketNo;
  }).catch(function (error) {
    console.log(error);
  });
}


// 将公式公的单元格替换为单元格对应的字符串
async function replaceReferencesInRange(SheetName,rangeAddress) {
  try {
    await Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getItem(SheetName);
      var range = sheet.getRange(rangeAddress);
      range.load("formulas");
      await context.sync();

      var formulas = range.formulas;
      var rowCount = formulas.length;
      var colCount = formulas[0].length;

      for (let i = 0; i < rowCount; i++) {
        for (let j = 0; j < colCount; j++) {
          let formula = formulas[i][j];
          let updatedFormula = formula;

          // 提取公式中的所有单元格引用
        //   var cellReferences = formula.match(/([A-Z]+[0-9]+)/g);
            var cellReferences = formula.match(/([A-Za-z]+\d+)/g);

          if (cellReferences) {
            for (let ref of cellReferences) {
              let cell = sheet.getRange(ref);
              cell.load("values");
              await context.sync();

              // 获取单元格的值，并将其替换到公式中
              let cellValue = cell.values[0][0].toString();
              updatedFormula = updatedFormula.replace(ref, cellValue);
            }
          }

          // 更新单元格中的公式
          range.getCell(i, j).values = [[`${updatedFormula.replace("=", "")}`]];
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}



//如果Result是除法结尾，则执行操作用公式代替sumif》》》》》如果Result的结果不是除法，乘法是不是也不能相加？也需要公式？
async function ResultDivided() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Bridge Data Temp");
    let range = sheet.getUsedRange().getRow(0);
    range.load("address, values");
    await context.sync();
    console.log("range.address is " + range.address);

    // 在Bridge Data Temp 第一行找到result的单元格
    let ResultCell = range.find("Result", {
      completeMatch: true,
      matchCase: true,
      searchDirection: "Forward"
    });
    ResultCell.load("address");
    await context.sync();
    console.log("ResultCell.address is " + ResultCell.address);

    //往下两行找到带有公式的单元格
    let ResultFormulaRange = ResultCell.getOffsetRange(2,0);
    ResultFormulaRange.load("address,formulas");
    await context.sync();
    console.log("ResultFormulaRange is " + ResultFormulaRange.formulas);

    //判断Result单元格的结果是SumN还是SumY
    ResultSumType = checkFormulaResultType(ResultFormulaRange.formulas[0][0].replace("=",""),FormulaTokens); //给全局变量赋值，判断是SumY还是SumN
    console.log("Result type is " + ResultSumType);

    let LastDivided = isLastOperatorDivision(ResultFormulaRange.formulas[0][0]);
    console.log("LastDivided is " + LastDivided.isDivision);
    console.log("Denominator is " + LastDivided.denominator);
    StrGlbIsDivided = LastDivided.isDivision; // 赋值给全局变量，在Process中计算Contribution的时候判断
    console.log("StrGlbIsDivided is " + StrGlbIsDivided);

    if(LastDivided){

      //往上一行找变量的标题
      let SecondRow = ResultFormulaRange.getOffsetRange(-1,0);
      SecondRow.load("values");
      await context.sync();

      //Formula 形成完整的 Room GOP=(Room Revenue+Room Labor+Room Exp.)/Room Revenue
      let Formula = ResultFormulaRange.address.split("!")[1] +  ResultFormulaRange.formulas[0][0]; //
      let ThirdRow = ResultFormulaRange.getOffsetRange(1, 0); // 放在公式单元格的下一行
      ThirdRow.values = [[Formula]];
      ThirdRow.load("address");
      await context.sync();
      console.log("Result formulas is " + Formula)

      // 获得公式中变量和变量名的对象
      let cellTitles = await getFormulaCellTitles("Bridge Data Temp", ThirdRow.address);
      objGlobalFormulasAddress = cellTitles;
      console.log(cellTitles);
      // 将变量名替代变量
      await replaceCellAddressesWithTitles("Bridge Data Temp", ThirdRow.address, ThirdRow.address, cellTitles);
      ThirdRow.load("values");
      await context.sync();
      let Denominator = isLastOperatorDivision(ThirdRow.values[0][0]).denominator; // 获取用Title而不是变量组成的分母
      console.log("Denominator in Title is " + Denominator);
      StrGlbDenominator = Denominator; //赋值给全局变量，后面计算contribution调用

      strGlobalFormulasCell = ThirdRow.address;
      console.log("ThirdRow.address is " + ThirdRow.address)
      console.log("Result strGlobalFormulasCell is " + strGlobalFormulasCell);
      console.log("Result strGlbBaseLabelRange is " + strGlbBaseLabelRange);

      await GetFormulasAddress("Bridge Data Temp", strGlobalFormulasCell ,"Process", strGlbBaseLabelRange);
      await CopyFormulas();

    }
    await context.sync();
  });
}


// 代码逻辑
// 去除外层多余括号：

// 在处理之前，首先去除公式最外层的括号（如果存在），以便更容易分析公式结构。
// 遍历公式字符：

// 使用 for 循环遍历公式中的每个字符，并且根据括号层次 (level) 来判断当前字符是否在最外层。
// 只有在最外层时，才记录操作符。
// 判断最后的操作符：

// 在 operators 数组中记录了所有最外层的操作符。最后判断数组中的最后一个操作符是否为除号 /。
// 示例输出
// 公式：((A+B)*C+D)/E
// 最后的操作符是否为 /：true
// 适用情况
// 此代码适用于包含任意数量括号和运算符的公式，并且可以正确判断公式中最外层的最后一个操作符是否为除号 /。
function isLastOperatorDivision(formula) {
  // 去掉公式外层的括号和等号
  //formula = formula.trim().replace("=", "");
  formula = formula.split("=")[1]; // 为了适应 A= B+C 这样的情况
  
  if (formula.startsWith("(") && formula.endsWith(")")) {
    formula = formula.substring(1, formula.length - 1).trim();
  }

  console.log("Formula in isDivision is " + formula);
  
  let operators = [];
  let level = 0;
  let lastDivisionIndex = -1; // 记录最后一个 '/' 的位置

  // 遍历公式中的每个字符
  for (let i = 0; i < formula.length; i++) {
    let char = formula[i];

    if (char === '(') {
      level++; // 进入新的括号层次
    } else if (char === ')') {
      level--; // 退出当前的括号层次
    } else if (level === 0 && ['+', '-', '*', '/'].includes(char)) {
      operators.push(char); // 只记录括号外的操作符
      if (char === '/') {
        lastDivisionIndex = i; // 记录最后一个除号的位置
      }
    }
  }

  // 判断最后一个操作符是否为 '/'
  if (operators.length > 0 && operators[operators.length - 1] === '/') {
    // 如果最后一个操作符是 '/', 返回分母
    let denominator = formula.substring(lastDivisionIndex + 1).trim();
    return {
      isDivision: true,
      denominator: denominator
    };
  }

  // 否则返回 false
  return {
    isDivision: false,
    denominator: null
  };
}



// //将公式里可加的数据尽量往左边移动
// async function reorderFormula(FormulaAddress) {
//   await Excel.run(async (context) => {
//     var sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
//     var formulaCell = sheet.getRange(FormulaAddress);
//     formulaCell.load("formulas");
//     await context.sync();

//     var formula = formulaCell.formulas[0][0];

//     // 移除公式中的等号
//     if (formula.startsWith("=")) {
//       formula = formula.substring(1);
//     }

//     // 匹配公式中的所有单元格引用和括号内的表达式
//     // var cellReferences = formula.match(/([A-Z]+\d+)/g);
//       // var cellReferences = formula.match(/([A-Za-z]+\d+)/g); //》》》》这里没有考虑其他的单纯的系数，例如（1+0），需要测试-1，-A1，等这种情况
//       var cellReferences = formula.match(/([A-Za-z]+\d+|\b\d+\b)/g); //考虑单纯的数字情况****需要加入判断小数，负数等
//       var parts = []; // 用来保存公式的每个部分

//     // 分割公式，保留运算符和括号
//     let formulaParts = formula.split(/([+\-*/()])/g).filter(part => part.trim() !== "");
//     for (let part of formulaParts) {
//       let isOperator = /[+\-*/()]/.test(part);
//       // let isNum = /^\d+$/.test(part); //》》》》》》需要加入判断小数，负数等
//       let isNum = /^-?\d+(\.\d+)?$/.test(part);//》》》已修改，不过在用户须知中应该告知可以填什么数值，科学计数法应该不行？
//       //》》》》》将运算符和变量等存入对象，需要考虑系数, 加入判断不能是数字
//       console.log("part from Cell is " + part);

//       if (!isOperator && cellReferences.includes(part) && !isNum) {    //如果不是运算符号，也不是数字，则在单元格里进行判断
//         let cellAbove = sheet.getRange(part).getOffsetRange(-2, 0); // 向上两行
//         cellAbove.load("values");
//         await context.sync();

//         // 判断是否为 SumN
//         parts.push({
//           value: part,
//           isOperator: isOperator,
//           isNonAdditive: cellAbove.values[0][0] === "SumN",
//         });
//         console.log("put 1 in SumN is ");
//         console.log(cellAbove.values[0][0] === "SumN");
//       } else if(!isOperator && cellReferences.includes(part) && isNum){ //如果不是运算符号，而是数字，直接判断为不是SumN

//         parts.push({
//           value: part,
//           isOperator: isOperator,
//           isNonAdditive: false,
//         });
//         console.log("put 2 in SumN is ");
//         console.log(false);
//       }
//       else {   // 其他情况，可能是某个变量名字？
//         parts.push({
//           value: part,
//           isOperator: isOperator,
//           isNonAdditive: false,
//           //precedence: getPrecedence(part) // 为运算符添加优先级
//         });
//         console.log("put 3 in SumN is ");
//         console.log(false);
//       }
//       //console.log(JSON.stringify(parts, null, 2));
//     }
//     console.log(JSON.stringify(parts, null, 2));
//     //console.log("parts is", parts)

//     // 重新构造公式》》》》》》这里没有考虑内嵌括号整体移动的情况，只是考虑了单个变量或者系数
//     let newFormula = [];
//     //let LoopCondition = true;
//     let LoopCondition = true;
//     while (LoopCondition) { // 有可能有多个SumN/SumN/SumY 这样的情况，每次都从头循环才能保证最后一个SumY到最前
//       let MoveNum = false; //计算循环一次有没有移动过变量
//       for (let i = 0; i < parts.length; i++) {
//         // if (parts[i].isOperator || parts[i].isNonAdditive || i == 0) {
//         if (parts[i].isOperator || parts[i].isNonAdditive || i == 0) {
//           console.log("part A i is " + i);
//           console.log(JSON.stringify(parts[i], null, 2));
//           // 如果变量是Non-Additive，第一个对象，或者是一个符号，则进入下一个迭代。
//           continue;

//         } else if (!parts[i].isOperator) {
//           console.log("part B i is " + i);
//           console.log(JSON.stringify(parts[i], null, 2));
//           // 如果parts[i]是非NonAdditive的变量, 则往前搜索

//           for (let j = i - 1; j >= 0; j--) {
//             // 如果变量前一个是 /, +, -, ( , )  或者循环到第一个对象，则进入下一个迭代。/ 应该也不需要处理，除非用户出错，因为不存在不可以加总的数除以可以加总的数，并且有意义的情况。
//             //》》》》》这里可能需要修改，因为发现了SumN / SumN 有意义的情况 /
//             //****下面的意思是如果有括号，则不考虑移动，后续需要将带括号的整体堪称一个对象进行移动 */
//             if (parts[j].value == "/" || parts[j].value == "+" || parts[j].value == "-" || parts[j].value == "(" || parts[j].value == ")" || j == 0) {
//               console.log("part C j is " + j);
//               console.log(JSON.stringify(parts[j], null, 2));
//               break;
//               //如果找到变量前是*，并且*再之前是一个不能相加的变量。A*B / A*++B /A*+-B 等情况
//             } else if (parts[j].value == "*" && !parts[j - 1].isOperator && parts[j - 1].isNonAdditive) {
//               console.log("part D j is " + j); 
//               console.log(JSON.stringify(parts[j], null, 2));
//               //则两个变量交换位置。
//               moveObjectInArray(parts, i, i - j); //把后面的可相加的数移动到前面
//               moveObjectInArray(parts, j-1, -(i - j +1)); //把前面不可相加的数移动到符号后面，因为被插入可相加的数，因此移动要+1, 往后移动前面要加负号
//               let formulaString = parts.map(part => part.value).join('');
//               console.log("formulaString is " + formulaString)
//               MoveNum = true;
//               break;
//             } 
//           }

//           continue;
//         }

//       }
//       //在循环到最后的时候判断有没有变量移动过
//       if (MoveNum) {
//         LoopCondition = true; //继续while循环

//       } else {
//         LoopCondition = false; //退出while循环
//       }

//       // LoopCondition--

//     }
//     // 更新公式
//     let formulaString = parts.map(part => part.value).join('');
//     console.log("formulaString is " + formulaString);
//     sheet.getRange(FormulaAddress).formulas = [[`=${formulaString}`]];
//     await context.sync();
//   });
// }

// //移动公式里的变量位置
// function moveObjectInArray(arr, index, num) {
//   console.log("before move arr is ");
//   console.log(JSON.stringify(arr, null, 2));
//   console.log("index is " + index);
//   console.log("num is " + num);
//   // 确保参数合法性
//   // if (index < 0 || index >= arr.length || num <= 0) {
//   //   console.error("Invalid index or num value");
//   //   return arr;
//   // }

//   // 计算目标位置
//   let newIndex = index - num;

//   // 确保目标位置不小于0
//   if (newIndex < 0) {
//     newIndex = 0;
//   }

//   // 获取要移动的对象
//   const objectToMove = arr.splice(index, 1)[0];

//   // 将对象插入到新位置
//   arr.splice(newIndex, 0, objectToMove);
//   console.log("after move arr is ");
//   console.log(JSON.stringify(arr, null, 2));
//   return arr;
// }


// -----生成公式的分解对象数组----这里修改为完全只为了给数据透视表筛选变量
async function processFormulaObj(RangeAddress) {
  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    let cell = sheet.getRange(RangeAddress);  // 获取单元格 Q3 的公式
    cell.load("formulas");
    await context.sync();  // 同步，确保公式已加载

    let formula = cell.formulas[0][0].replace("=", "");  // 获取 Q3 中的公式字符串
    console.log(formula)

    // 分割公式，保留运算符和括号
    let formulaParts = formula.split(/([+\-*/()])/g).filter(part => part.trim() !== "");
    console.log("formulaParts is");
    console.log(JSON.stringify(formulaParts, null, 2));

    //-----------------为了数据透视表的筛选变量，只显示和公式相关的变量----------------------------
    let formulaPartsNoOperator = formulaParts.filter(part => part.trim() !== "" && !/[+\-*/()]/.test(part)); // 去掉空白部分和运算符号供获取变量给数据透视表
    console.log("formulaPartsNoOperator is");
    console.log(JSON.stringify(formulaPartsNoOperator, null, 2));

    let TempArrVar = []; //为数据透视表筛选数据
    // 遍历每个部分，创建对应的对象并加入数组
    for (let part of formulaPartsNoOperator) {
      // 检查 part 是否为纯数字，如果不是则获取单元格范围
      if (!/^\d+$/.test(part)) {
        let refCell = sheet.getRange(part);  // 获取变量对应的单元格

        let cellAbove = refCell.getOffsetRange(-1, 0);  // 获取上面两行的单元格

        cellAbove.load("values");  // 加载上面两行单元格的值

        TempArrVar.push(cellAbove); // 暂存 cellAbove 对象
      }
    }
    // 同步加载所有数据
    await context.sync();  // 确保所有 cellAbove 的值已加载

    // 遍历已加载的数据并提取值
    TempArrVar = TempArrVar.map(cellAbove => cellAbove.values[0][0]); // 使用这个数据包含的变量筛选数据透视表
    console.log("TempArrVar is ");
    console.log(JSON.stringify(TempArrVar, null, 2));

    //ArrVarPartsForPivotTable = ArrVarPartsForPivotTable.concat(TempArrVar); // 追加透视表筛选数据
    ArrVarPartsForPivotTable = Array.from(new Set([...ArrVarPartsForPivotTable, ...TempArrVar]));// 追加透视表筛选数据，并保证没有重复变量
    console.log("ArrVarPartsForPivotTable is ");
    console.log(JSON.stringify(ArrVarPartsForPivotTable, null, 2));
    //-----------------为了数据透视表的筛选变量，只显示和公式相关的变量----End------------------------


    // let formulaArray = []; // 存储解析出来的公式部分
    // let cellsToLoad = []; // 存储所有需要加载的单元格

    // // 遍历每个部分，创建对应的对象并加入数组
    // for (let part of formulaParts) {
    //   let isOperator = /[+\-*/()]/.test(part); // 判断是否为运算符或括号
    //   let formulaObj = {
    //     formulaParts: part,
    //     NonAdditive: false, // 默认false
    //     isOperator: isOperator, // 根据正则判断
    //     refCell: null, // 记录引用单元格
    //     cellAbove: null // 记录上面两行的单元格
    //   };

    //   // 处理变量部分，如果不是运算符或括号
    //   if (!isOperator) {
    //     //// 检查 part 是否为纯数字，如果不是则获取单元格
    //     if (!/^\d+$/.test(part)) {
    //       let refCell = sheet.getRange(part); // 获取变量对应的单元格
    //       let cellAbove = refCell.getOffsetRange(-2, 0); // 获取上面两行的单元格

    //       cellAbove.load("values"); // 加载上面两行单元格的值
    //       cellsToLoad.push({ formulaObj, cellAbove }); // 存储需要加载的单元格和对应的对象
    //     }
    //   }

    //   // 将对象加入 formula 数组
    //   formulaArray.push(formulaObj);
    // }

    // // 一次性同步所有加载的数据
    // await context.sync();

    // // 遍历加载的单元格并更新 formulaObj
    // for (let { formulaObj, cellAbove } of cellsToLoad) {
    //   if (cellAbove.values[0][0] === "SumN") {
    //     formulaObj.NonAdditive = true; // 如果上方单元格是 "SumN"，更新对象属性, 对象是引用的，因此修改会直接影响formulaArray里的对象
    //   }
    // }

    // // 输出结果，您可以根据需要将其存储或进一步处理
    // console.log(JSON.stringify(formulaArray, null, 2));



    // return formulaArray;
  }).catch(function (error) {
    console.log(error);
  });
}



// 生成公式的分解对象数组
async function processFormulaObjforSplitDividend(Formula) {
  return await Excel.run(async (context) => {
    // let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    // let cell = sheet.getRange(RangeAddress);  // 获取单元格 Q3 的公式
    // cell.load("formulas");
    // await context.sync();  // 同步，确保公式已加载

    let formula = Formula.replace("=", "");  // 获取 Q3 中的公式字符串
    console.log(" processFormulaObj is ");
    console.log(formula);

    // 分割公式，保留运算符和括号
    let formulaParts = formula.split(/([+\-*/()])/g).filter(part => part.trim() !== "");
    console.log("formulaParts is");
    console.log(JSON.stringify(formulaParts, null, 2));

  
    let formulaArray = []; // 存储解析出来的公式部分
    let cellsToLoad = []; // 存储所有需要加载的单元格

    // 遍历每个部分，创建对应的对象并加入数组
    for (let part of formulaParts) {
      let isOperator = /[+\-*/()]/.test(part); // 判断是否为运算符或括号
      let formulaObj = {
        formulaParts: part,
        // NonAdditive: false, // 默认false
        isOperator: isOperator, // 根据正则判断
        // refCell: null, // 记录引用单元格
        // cellAbove: null // 记录上面两行的单元格
      };


      // 将对象加入 formula 数组
      formulaArray.push(formulaObj);
    }

    // 一次性同步所有加载的数据
    await context.sync();


    // 输出结果，您可以根据需要将其存储或进一步处理
    console.log("formulaArray is ");
    console.log(JSON.stringify(formulaArray, null, 2));

    return formulaArray;
  }).catch(function (error) {
    console.log(error);
  });
}






//找到公式中连续除号的位置
function checkConsecutiveDivisions(formulaArray) {
  let consecutiveDivisions = 0;
  let positions = [];  // 用于存储所有连续除号的位置
  let currentStart = -1;  // 记录当前连续除号的开始位置
  let currentEnds = [];  // 用于存储当前连续除号的结束位置

  for (let i = 0; i < formulaArray.length; i++) {
    let obj = formulaArray[i];

    if (obj && obj.formulaParts !== undefined && obj.isOperator) {
      if (obj.formulaParts === "/") {
        consecutiveDivisions++;

        // 如果这是第一个除号，记录其起始位置
        if (consecutiveDivisions === 1) {
          currentStart = i;
          currentEnds = [];  // 清空当前的结束位置
        }

        // 记录后续连续的除号位置
        if (consecutiveDivisions > 1) {
          currentEnds.push(i);
        }
      } else {
        // 当遇到非除号时，检查是否有连续除号要存储
        if (consecutiveDivisions >= 2) {
          let divisionPositions = [currentStart, ...currentEnds];
          positions.push(divisionPositions);  // 只存储一次连续除号
        }

        // 重置计数器和当前的开始、结束位置
        consecutiveDivisions = 0;
        currentStart = -1;
        currentEnds = [];
      }
    }
  }

  // 在循环结束后，检查最后是否还有未存储的连续除号
  if (consecutiveDivisions >= 2) {
    let divisionPositions = [currentStart, ...currentEnds];
    positions.push(divisionPositions);  // 存储最后一组连续除号
  }

  // 返回包含连续除号位置信息的对象
  return { result: positions.length > 0, positions: positions };
}

// 修改公式，插入括号和运算符替换
function modifyFormula(formulaArray, positions) {
  let modifiedFormula = [...formulaArray];  // 创建一个新的数组以避免直接修改原数组
  let offset = 0;  // 用于记录插入括号后导致的索引偏移

  positions.forEach(group => {
    let start = group[0] + offset; // 加上偏移量
    let end = group[group.length - 1] + offset; // 加上偏移量

    // 在第一个除号左边加上左括号
    modifiedFormula.splice(start+1, 0, { formulaParts: "(", isOperator: true });
    offset++;  // 插入左括号后，公式长度增加

    // 将除了第一个之外的除号替换为乘号, 需要先执行这一步再执行下一步，因为这一步不需要offset++，暂时长度不需要改变
    for (let i = 1; i < group.length; i++) {
      modifiedFormula[group[i] + offset].formulaParts = "*";
    }

    // 在最后一个除号右边的右操作数后加上右括号
    modifiedFormula.splice(end + 2 +1, 0, { formulaParts: ")", isOperator: true });
    offset++;  // 插入右括号后，公式长度增加

  });

  return modifiedFormula;
}

// 将数组合并输出公式
function formatFormula(formulaArray) {
  return formulaArray.map(part => part.formulaParts).join('');
}





//输入起始索引和相对索引获得在工作表中的地址
function getCellAddress(baseIndex, offsetIndex) {
      // 辅助函数：将列索引转换为列字母
      function indexToColumn(colIndex) {
        let col = "";
        colIndex++; // 转为 1-based
        while (colIndex > 0) {
          const remainder = (colIndex - 1) % 26;
          col = String.fromCharCode(65 + remainder) + col; // 根据 remainder 动态生成字符
          colIndex = Math.floor((colIndex - 1) / 26);
        }
        return col;
      }

      // 基础单元格的行和列索引
      const [baseRow, baseCol] = baseIndex;
      // 偏移量的行和列索引
      const [offsetRow, offsetCol] = offsetIndex;

      // 计算目标单元格的行和列索引
      const targetRow = baseRow + offsetRow;
      const targetCol = baseCol + offsetCol;

      // 转换为 A1 地址
      const columnLetter = indexToColumn(targetCol);
      return `${columnLetter}${targetRow + 1}`; // 行号为 1-based
}

// 示例调用
// const baseIndex = [9, 6]; // A 的索引，例如 G10 对应 [9, 6]
// const offsetIndex = [0, 1]; // B 相对于 A 的偏移量，例如 +1 列
// console.log(getCellAddress(baseIndex, offsetIndex)); // 输出: "H10"

//补齐数组中有空行的长度
function normalizeArray(arr) {
  // 找出最长行的长度
  const maxColCount = Math.max(...arr.map(row => row.length));

  // 补齐所有行到相同的列数
  return arr.map(row => {
    const newRow = [...row]; // 克隆当前行
    while (newRow.length < maxColCount) {
      newRow.push(null); // 用 null 补齐
    }
    return newRow;
  });
}


async function Contribution() {
  await Excel.run(async (context) => {
    console.log("Contribution Start");
    let parts = null;
    //如果StrGlbDenominator 不是null,则说明计算的最后一步是除法，可以运行下面的代码，如果是null，则说明最后一步是除法以外的，下面用另外的算法
    if(StrGlbDenominator !== null){
      console.log("Contribution 0");
      let FormulaTitle = StrGlbDenominator; //*****/需要改成参数传递

      // 正则解释：
      // 1. [^\+\-\*\/\(\)]+ 匹配连续的非运算符字符（即变量名部分）
      // 2. [\+\-\*\/\(\)] 匹配单个运算符：+、-、*、/、(、)
      // 使用全局标志 g 进行匹配
      if(checkType2Var.includes(StrGlbDenominator)){

        parts = [`${StrGlbDenominator}`]; //如果StrGlbDenominator是SumN+SumN的一部分，则不需要做拆分处理

      } else{
          const regex = /[^\+\-\*\/\(\)]+|[\+\-\*\/\(\)]/g;
          parts = FormulaTitle.match(regex);
          // 去除token前后可能存在的空白，并过滤掉空串
          parts = parts.map(token => token.trim()).filter(token => token !== "");
      }
      console.log("StrGlbDenominator is " + StrGlbDenominator);
      console.log("parts is ");
      console.log(parts);
    }
    console.log("Contribution 1");
    let ProcessSheet = context.workbook.worksheets.getItem("Process");
    let UsedRange = ProcessSheet.getUsedRange();
    UsedRange.load("address,values,rowCount,columnCount");
    let ProcessRange = ProcessSheet.getRange(StrGlobalProcessRange); //*** */ 需要改成参数传递
    console.log("StrGlobalProcessRange is " + StrGlobalProcessRange);

    let ProcessStartRange = ProcessRange.getCell(0, 0); //左上角第一个单元格
    ProcessRange.load("address,values,rowCount,columnCount");
    await context.sync();

    //-------------------------------------------
    let ProcessAddress = getRangeDetails(ProcessRange.address);
    let ProcessLastColumn = ProcessAddress.rightColumn //最右边的列
    let ProcessBottomRow = ProcessAddress.bottomRow //最下边的列

    let ProcessRangeRightTop = ProcessRange.getCell(0, ProcessRange.columnCount-1); //获得Range右上角的单元格，为之后拷贝格式
    let ProcessRangeRightBottom = ProcessRange.getCell(ProcessRange.rowCount-1, ProcessRange.columnCount-1); //获得Range右下角的单元格，为之后拷贝格式
    //------------------------------------------

    // let ProcessLastColumn = getRangeDetails(ProcessRange.address).rightColumn //最右边的列
    // let ProcessFirstRow = getRangeDetails(ProcessRange.address).topRow //最上面的行
    // let ProcessLastRow = getRangeDetails(ProcessRange.address).bottomRow //最下面的行

    console.log("ProcessLastColumn is " + ProcessLastColumn);
    let AllProcessFirstRow = ProcessSheet.getRange(`A1:${ProcessLastColumn}1`); // 整个ProcessSheet的第一行
    let AllProcessSecondRow = AllProcessFirstRow.getOffsetRange(1, 0); //// 整个ProcessSheet的第二行
    let AllProcessThirdRow = AllProcessFirstRow.getOffsetRange(2, 0); //// 整个ProcessSheet的第三行
    let AllProcessFourthdRow = AllProcessFirstRow.getOffsetRange(3, 0); //// 整个ProcessSheet的第三行
    let AllProcessLastCell = ProcessSheet.getRange(`${ProcessLastColumn}3`); // 整个Process第三行最后一个单元格
    let MixStartCell = AllProcessLastCell.getOffsetRange(0, 3); //往右移动3格获得Mix起始单元格
    let MixFirstRow = MixStartCell.getOffsetRange(1, 0); //往下第一个有公式的格子

    //计算出dominator和MixRange的最大的Range，为了获得最右边的列，进而建立一个起点为整张工作表的Range，为了后面获得全局地址
    let ProcessExtentRange = MixStartCell.getAbsoluteResizedRange(1,ProcessRange.columnCount*2); 
    ProcessExtentRange.load("address");

    console.log("Contribution 2.1");

    let ProcessTitle = ProcessStartRange.getOffsetRange(0, 1).getAbsoluteResizedRange(1, ProcessRange.columnCount - 1);//需要循环的标题
    let ProcessType = ProcessTitle.getOffsetRange(-2, 0); //Type, Result，SumN 等类型
    ProcessTitle.load("address,values,rowCount,columnCount");
    ProcessType.load("address,values");
    console.log("Contribution 2.2");
    AllProcessFirstRow.load("address,values,rowCount,columnCount");
    AllProcessSecondRow.load("address,values,rowCount,columnCount");
    AllProcessThirdRow.load("address,values,rowCount,columnCount");
    AllProcessFourthdRow.load("address,values,rowCount,columnCount");
    MixStartCell.load("address,rowIndex,columnIndex");
    MixFirstRow.load("address");
    console.log("Contribution 2")
    await context.sync();

    //---------------------------------------------------
    let ProcessExtentRangeRightColumn = getRangeDetails(ProcessExtentRange.address).rightColumn;
    //Process拓展后，包含denominator 和 Mix的工作表的全部单元格，没有直接使用UsedRange
    let ProcessAllRange = ProcessSheet.getRange(`A1:${ProcessExtentRangeRightColumn}${ProcessBottomRow}`); 
    ProcessAllRange.load("values,formulas,address,rowCount,columnCount");
    await context.sync();
    //---------------------------------------------------

    //--------------------------------------------------
    console.log("MixStartCell is " + MixStartCell.address);
    console.log("ProcessAllRange.address is " + ProcessAllRange.address);
    let ProcessAllRangeAddress = await GetRangeAddress("Process",ProcessAllRange.address); // 获得每个单元格的地址信息

    //--------------------------------------------------

    // let FirstRowAddress = await GetRangeAddress("Process",AllProcessFirstRow.address);
    // let SecondRowAddress = await GetRangeAddress("Process",AllProcessSecondRow.address);
    // let ThirdRowAddress = await GetRangeAddress("Process",AllProcessThirdRow.address);
    // let FourthRowAddress = await GetRangeAddress("Process",AllProcessFourthdRow.address);
    console.log("ProcessTitle is " + ProcessTitle.address);
    console.log("ProcessRange is " + ProcessRange.address);
    console.log("ProcessType is " + ProcessType.address);
    console.log("ProcessFirstRow is " + AllProcessFirstRow.address);
    console.log("ProcessSecondRow is " + AllProcessSecondRow.address);
    console.log("MixStartCell is " + MixStartCell.address);

    //找到Result的单元格
    // let ResultRange = ProcessType.find("Result", {
    //   completeMatch: true,
    //   matchCase: true,
    //   searchDirection: "Forward"
    // });
    // let ResultFormulaRange = ResultRange.getOffsetRange(3, 0);
    // ResultRange.load("address");
    // ResultFormulaRange.load("address,formulas,values");
    // await context.sync();
    // console.log("ResultRange is " + ResultRange.address);
    // console.log("ResultFormula is " + ResultFormulaRange.formulas[0][0])

    // 初始化数组并添加开头的元素
    let arr = [["BasePT", "SumY"]];

    // for (let z = 0; z < ProcessTitle.columnCount; z++) {
    //   let TitleCell = ProcessTitle.getCell(0, z);
    //   let TitleType = ProcessTitle.getCell(-2, z);
    //   TitleCell.load("address,values");
    //   TitleType.load("address,values");
    //   await context.sync();
    //   //下面这些数据类型不进入Contribution的计算，防止result在插在变量中出现的时候后面的查询出现问题
    //   if (!["Result", "ProcessSum", "Impact", "","NULL"].includes(TitleType.values[0][0])) {
    //       arr.push([TitleCell.values[0][0], TitleType.values[0][0]]);
    //   }
    // }

    for (let z = 0; z < ProcessTitle.columnCount; z++) {
      let TitleCell = ProcessTitle.values[0][z];
      // let TitleCell = ProcessAllRange.values[2][z];   //AllRange的第三行
      let TitleType = ProcessType.values[0][z];
      // let TitleType = ProcessAllRange.values[0][z];   //AllRange的第一行
      //下面这些数据类型不进入Contribution的计算，防止result在插在变量中出现的时候后面的查询出现问题
      if (!["Result", "ProcessSum", "Impact", "","NULL"].includes(TitleType)) {
          arr.push([TitleCell, TitleType]);
      }
    }

    // 在数组末尾添加指定的元素
    arr.push(["TargetPT", "SumY"]);

    console.log("arr is : " + JSON.stringify(arr));

    //-------------------------------------------------------------------------
    //创建一个二维数组，用于存放动态生成分母和Mix的formulas 或者是values
    let MixArrRow = ProcessAllRange.rowCount;
    let MixArrColumn = ProcessTitle.columnCount;
    console.log("MixArrRow is " + MixArrRow);
    console.log("MixArrColumn is " + MixArrColumn);

    //获得从工作表第1行，Index为0开始的Domination和MixRange的起始Index,作为后面用相对Index计算出工作表的绝对Index，进而计算Address
    let MixStartRowIndex = MixStartCell.rowIndex - 2; 
    let MixStartColumnIndex = MixStartCell.columnIndex;
    let MixStartIndex = [MixStartRowIndex, MixStartColumnIndex]; 
    console.log("MixStartRowIndex is " + MixStartRowIndex);
    console.log("MixStartColumnIndex is " + MixStartColumnIndex);
    console.log("MixStartIndex is");
    console.log(MixStartIndex);

    // 创建一个二维数组，所有元素初始为 null，大小为需要填入的ProcessRange单元格
    // const MixArr = Array.from({length: MixArrRow}, () => new Array(MixArrColumn).fill(null));
    let MixArr = Array.from({length: MixArrRow}, () => new Array(0).fill(null)); // 列设为0，动态添加，行需要固定好
    // let MixArr = [[]]; // 创建动态的数组
    //Array.from({ length: rows }, () => Array(initialCols).fill(null))
    //-------------------------------------------------------------------------

    //判断Result列的公式最后一步是否是除法
    //let StrGlbIsDivided = true; //》》》》测试用，整合需要删掉》》》

    //let VarStartRange = AllProcessThirdRow.getCell(0,0);
    //console.log("VarRange is " + VarRange);

    console.log("StrGlbIsDivided is " + StrGlbIsDivided);

    let ContributionStartCell = null; // Process表中Contribution的起始单元格，也为后面variance 表格做为基础地址使用

    //先判断最后一步是否是除法
    if (StrGlbIsDivided && ResultSumType === "SumN") {
      console.log("Enter Mix");
      // 循环每个变量，计算出每一步变量变化对应的被除数的Mix
      let iColumn = 0;
      for (let z = 0; z < arr.length; z++) {
        let Title = arr[z][0];
        let Type = arr[z][1];
        console.log(`arr[${z}][1] is` + arr[z][1] );
        //TitleCell.load("address,values");
        //TitleType.load("address,values");
        //await context.sync();
        console.log("Enter Mix 1");
        console.log("StrGlbDenominator is " + StrGlbDenominator);
        // let cellName = null;
        // 创建 parts 的副本，避免修改原数组

        // if (parts) {
        //   console.log("parts is " + JSON.stringify(parts));
        // } else {
        //   console.log("parts is undefined or null");
        // }

        let TempParts = [...parts]; // 使用扩展运算符创建一个新的副本
        console.log("TempParts is ");
        console.log(TempParts);

        if (!["Result", "ProcessSum", "Impact", "NULL",""].includes(Type)) {

          console.log("TitleCell is " + Title);


          // 遍历数组parts 中的所有变量, 在process第三行中找到相应的单元格
          for (let i = 0; i < parts.length; i++) {
            let variable = parts[i];
            console.log("variable is " + variable);

            // 只处理变量（忽略运算符和括号）
            if (/[^+\-*/^()]+/.test(variable)) {  // 检测非运算符、非括号的变量

              // 遍历 RangeA 查找所有匹配的单元格
              for (let j = 0; j < ProcessAllRange.columnCount; j++) {
                // let VarCell = ProcessAllRange.values[2][j];   // getCell(0, j); 
                // VarCell.load("address,values"); //获取第三行的每个单元格
                // await context.sync();

                // 如果单元格的值等于当前变量名
                if (ProcessAllRange.values[2][j] === variable && ProcessAllRange.values[1][j] === Title) {
                  console.log("variable2 is " + variable);
                  // console.log("VarCell is " + VarCell);
                  // 检查符合条件的单元格
                  // let upperCell = AllProcessFirstRow.values[0][j];

                  // 判断上方两行的单元格是否符合条件

                  // 查找上方一行是否等于变量名
                  // let oneRowUp = AllProcessSecondRow.values[0][j];
                  // let oneRowDown = AllProcessFourthdRow.values[0][j];

                  // console.log("oneRowUp is " + oneRowUp);
                  console.log("TitleCell is " + Title);
                  // if (oneRowUp === Title ) {
                    // 符合条件，使用该单元格
                    // console.log("oneRowUp is OK " + oneRowUp);
                    //cellName = VarCell;
                    // ProcessAllRangeAddress
                    // let cellAddress = FourthRowAddress[0][j].split('!')[1];
                    let cellAddress = ProcessAllRangeAddress[3][j].split('!')[1];
                    // 将变量替换为 Cell Var 的地址
                    TempParts[i] = cellAddress;
                    console.log("TempParts is " + TempParts);


                    break;  // 找到符合条件的单元格后退出循环
                  // }
                }
              }
            }

            // 如果找到了符合条件的 cellName，继续处理
            //console.log("cellName is " + cellName);
            // if (cellName) {
            //   // 查找符合条件的Cell Var并往下移动一行
            //   let cellVar = cellName.getOffsetRange(1, 0);
            //   cellVar.load("address");
            //   await context.sync();

            //   // 获取 Cell Var 的地址，例如 A1 样式
            //   let cellAddress = cellVar.address.split('!')[1];

            //   // 将变量替换为 Cell Var 的地址
            //   TempParts[i] = cellAddress;
            //   console.log("TempParts is " + TempParts);
            // }
          }

          let finalFormula = "=" + TempParts.join("");
          console.log("finalFormula is " + finalFormula);
          console.log(`${Title} End~!!!!`)
          // 重新创建 TempParts 的副本，避免影响下一次循环
          TempParts = [...parts];

          //将替换好的公式放入Process对应的单元格，作为每一个变量替换后的分母
          // MixStartCell.values = [[Title]];
          //从第三行开始时标题
          MixArr[2][iColumn] = `="${Title}"`;
          // let MixFirstRow = MixStartCell.getOffsetRange(1, 0);
          //第四行开始是第一行带有公式formulas的单元格
          MixArr[3][iColumn] = finalFormula;
          // MixFirstRow.numberFormat = '#,##0.00'; // 设置数据格式，因为时计算出来的，无法复制前面的单元格
          // MixFirstRow.load('address');
          // let MixTwoUpRow = MixStartCell.getOffsetRange(-2, 0);
          MixArr[0][iColumn] = `="Denominator"`;
          // await context.sync();

          // let MixRange = MixFirstRow.getAbsoluteResizedRange(ProcessRange.rowCount - 1, 1); // Mix的一整列
          // MixRange.load("address");
          // // MixRange.copyFrom(MixFirstRow, Excel.RangeCopyType.formulas, false, false); 
          // MixRange.copyFrom(MixFirstRow, Excel.RangeCopyType.formulasAndNumberFormats, false, false);//将公式拷贝到一整行
          // await context.sync();
          //从第一个变量单元格开始往右移动，从第4行开始，Z列是计算denominator，(z+1)*2-1是计算Mix
          let DenominatorCellAddress = ProcessAllRangeAddress[3][MixStartColumnIndex + z*2]; 
          console.log("DenominatorCellAddress is " + DenominatorCellAddress);
          let DenominatorAddressDetail = getRangeDetails(DenominatorCellAddress);
          let DenominatorTopRow = DenominatorAddressDetail.topRow;
          let DenominatorColumn = DenominatorAddressDetail.leftColumn;
          let DenominatorBottom = ProcessBottomRow;   //getRangeDetails(MixRange.address).bottomRow;

          //计算Mix
          
          // MixStartCell = MixStartCell.getOffsetRange(0, 1); // 自身往右移动一格
          MixArr[2][iColumn +1] = `="${Title}"`; // 往右移动一格
          // let MixTwoUpRow = MixStartCell.getOffsetRange(-2, 0);
          MixArr[0][iColumn +1] = `="Mix"`;
          let MixFormula = `=IFERROR(${DenominatorColumn}${DenominatorTopRow}/\$${DenominatorColumn}\$${DenominatorBottom},0)`
          console.log("MixFormula is " + MixFormula);
          // MixFirstRow = MixFirstRow.getOffsetRange(0, 1);
          MixArr[3][iColumn +1] = MixFormula;
          // 设置百分比格式并保留两位小数
          // MixFirstRow.numberFormat = '0.00%';
          // await context.sync();

          // MixRange = MixRange.getOffsetRange(0, 1);
          // MixRange.copyFrom(MixFirstRow, Excel.RangeCopyType.formulasAndNumberFormats, false, false);

          // MixStartCell = MixStartCell.getOffsetRange(0, 1); // 自身往右移动一格
          // await context.sync();
          
        }
        iColumn = iColumn + 2;
      }
      //console.log("test1")

      MixArr = normalizeArray(MixArr); // 补齐其中有的空行，使得列数一样，如果数组不对齐，则不能给单元格赋值fomulas      
      console.log("MixArr is");
      console.log(MixArr);

      // 获取行数
      let MixRowCount = MixArr.length;
      // 获取列数（假设所有行的列数相同）
      let MixColCount = MixArr[0].length;
      let InputMixStartCell = MixStartCell.getOffsetRange(-2,0); //从第一行开始的单元格
      let InputMixRange = InputMixStartCell.getAbsoluteResizedRange(MixRowCount, MixColCount);
      InputMixRange.formulas = MixArr;
      
      // 复制第 4 行到第 5 行及之后的所有行
      let rowToCopy = InputMixRange.getRow(3); // 第 4 行
      let rangeToFill = InputMixRange.getRow(4).getOffsetRange(0, 0).getAbsoluteResizedRange(MixRowCount - 4, MixColCount); // 第 5 行到最后一行
      rangeToFill.copyFrom(rowToCopy, Excel.RangeCopyType.formulas); // 复制公式

      MixStartCell = MixStartCell.getOffsetRange(0, MixColCount); //移动到Denomination 和 Mix单元格之后
      await context.sync();
      console.log("Contribution 7.1");


      //开始整理格式
      InputMixRange.getRow(2).copyFrom(ProcessRangeRightTop,Excel.RangeCopyType.formats); //复制标题格式
      InputMixRange.getRow(MixRowCount - 1).copyFrom(ProcessRangeRightBottom,Excel.RangeCopyType.formats); //复制汇总格式

      // 获取从第 4 行开始的范围
      const rangeFromFourthRow = InputMixRange.getRow(3).getAbsoluteResizedRange(MixRowCount - 3, MixColCount);

      // 遍历每一列 设置数据格式
      for (let colIndex = 0; colIndex < MixColCount; colIndex++) {
        const columnRange = rangeFromFourthRow.getColumn(colIndex);

        if ((colIndex + 1) % 2 === 1) {
          // 单数列（1, 3, 5, ...）
          columnRange.numberFormat = '#,##0.00';
        } else {
          // 双数列（2, 4, 6, ...）
          columnRange.numberFormat = '0.00%';
        }
      }

      await context.sync();
      console.log("Contribution 7.2");

      //------------------------------计算contribution---------------------------------------
      ContributionStartCell = MixStartCell.getOffsetRange(0,1) // 往右移动一格
      let NewUsedRange = ProcessSheet.getUsedRange(); // 这里的UsedRange是Process 工作表的更新后的适用范围
      let FirstRow = NewUsedRange.getRow(0); 
      let SecondRow = NewUsedRange.getRow(1);
      let ThirdRow = NewUsedRange.getRow(2);
      let FourthRow = NewUsedRange.getRow(3);
      let BottomRow = NewUsedRange.getRow(UsedRange.rowCount-1); //这里可以使用上一步的UsedRange，省去一次sync
      NewUsedRange.load("address,values")
      FirstRow.load("address,values");
      SecondRow.load("address,values");
      ThirdRow.load("address,values");
      FourthRow.load("address,values");
      BottomRow.load("address,values");
      await context.sync();

      // let ProcessFirstRowAddress = await GetRangeAddress("Process",FirstRow.address);
      // let ProcessSecondRowAddress = await GetRangeAddress("Process",SecondRow.address);
      // let ProcessThirdRowAddress = await GetRangeAddress("Process",ThirdRow.address);
      let ProcessFourthRowAddress = await GetRangeAddress("Process",FourthRow.address);
      let ProcessBottomRowAddress = await GetRangeAddress("Process",BottomRow.address);
      console.log("NewUsedRange is " + NewUsedRange.address);
      //console.log("test2")
      //不循环BasePT 和 Target PT，因此z =1, arr.length -1
      for (let z = 1 ; z < arr.length -1; z++) {
        console.log("Current Var is" + arr[z][0]);
        //在第一行找到Mix 以及对应的当前变量
        // let CurrentMixTitle = null;
        // let BeforeMixTitle = null;
        // let CurrentResultCell = null;
        // let BeforeResultCell = null;
        // let BeforeTotalResultCell = null;

        // let CurrentMixAddress = null;
        let CurrentResultAddress = null;
        let BeforeMixAddress = null;
        let BeforeResultAddress = null;
        let BeforeTotalResultAddress = null;

        //console.log("FirstRowValues length is " + FirstRow.values[0].length)
        //下面必须是FirstRowValues[0].length，而不能是FirstRowValues.length，这样length是1，因为只有一行
        for (let i = 0; i < FirstRow.values[0].length -1;i++){
          //console.log("test4")

            //找到当前变量对应的Result的相关信息
            if (FirstRow.values[0][i] === "Result" && arr[z][0] === SecondRow.values[0][i]) {
                // CurrentResultCell = SecondRow.getCell(0,i).getOffsetRange(2,0); //获取下面两格，其中的包含Result结果单元格
                // CurrentResultCell.load("address,values");
                // await context.sync();
                console.log("ProcessFourthRowAddress[0][i] is " + ProcessFourthRowAddress[0][i]);
                // CurrentResultAddress = CurrentResultCell.address.split("!")[1]; //获取Current地址
                CurrentResultAddress = ProcessFourthRowAddress[0][i].split("!")[1]; 
                console.log("CurrentResultAddress is " + CurrentResultAddress)
            }

            //找到前一个Mix的相关信息,这里需要是arr[z-1][0]
            //for (let j = 0; j < FirstRow.values[0].length - 1; j++) {
            // console.log("FirstRow.values[0][i] is " + FirstRow.values[0][i]);
            // console.log("ThirdRow.values[0][i] is " + ThirdRow.values[0][i]);
            if (FirstRow.values[0][i] === "Mix" && arr[z-1][0] === ThirdRow.values[0][i]) {
                  // let BeforeMixTitle = ThirdRow.getCell(0,i);
                  // let BeforeMixCell = BeforeMixTitle.getOffsetRange(1, 0);//往下移动一格，找到带有值的Mix
                  // BeforeMixTitle.load("address,values");
                  // BeforeMixCell.load("address,values");
                  // await context.sync();
                  console.log("ProcessFourthRowAddress[0][i] 2 is " + ProcessFourthRowAddress[0][i]);
                  BeforeMixAddress = ProcessFourthRowAddress[0][i].split("!")[1]; //获取单元格Mix地址A1等
                  console.log("BeforeMixAddress is " + BeforeMixAddress);

              //}

            }

            //找到前一个变量对应的Result的相关信息，这里需要arr[z-1][0]
            if (FirstRow.values[0][i] === "Result" && arr[z-1][0] === SecondRow.values[0][i]) {
              // BeforeResultCell = SecondRow.getCell(0, i).getOffsetRange(2, 0); //获取下面两格，其中的包含Result结果单元格
              // BeforeTotalResultCell = SecondRow.getCell(0, i).getOffsetRange(ProcessRange.rowCount, 0); //获取最下面一行的Total Result
              // BeforeResultCell.load("address,values");
              // BeforeTotalResultCell.load("address,values");
              // await context.sync();

              BeforeResultAddress = ProcessFourthRowAddress[0][i].split("!")[1]; //获取Current地址
              BeforeTotalResultAddress = ProcessBottomRowAddress[0][i].split("!")[1]; //获取Current地址
              console.log("BeforeResultAddress is " + BeforeResultAddress);
              console.log("BeforeTotalResultAddress is " + BeforeTotalResultAddress);
            }

            //如果第一行是Mix, 并且第三行等于数组中的变量，则i就是对应的列
            //执行到这一步，上面的if应该已经把contribution的公式变量都找到了
            if (FirstRow.values[0][i] === "Mix" && arr[z][0] === ThirdRow.values[0][i]) {
              //console.log("FirstRow.values[0][i] is " + FirstRow.values[0][i]);
              //console.log("ThirdRow.values[0][i] is " + ThirdRow.values[0][i]);
              //console.log("I is " + i);
              let CurrentMixTitle = ThirdRow.getCell(0, i); //找到对应的第三行的标题
              let CurrentMixCell = CurrentMixTitle.getOffsetRange(1, 0); //往下移动一格，找到带有值的Mix 
              let CurrentType = ContributionStartCell.getOffsetRange(-2, 0); //获contribution取标题单元格
              CurrentType.values = [["Contribution"]];
              ContributionStartCell.copyFrom(CurrentMixTitle); //复制标题
              CurrentMixTitle.load("address,values");
              CurrentMixCell.load("address,values");
              await context.sync();

              let CurrentMixAddress = CurrentMixCell.address.split("!")[1]; //获取单元格Mix 地址 A1等
              console.log("CurrentMixAddress is " + CurrentMixAddress);
              
              //找到了全部变量，开始生成公式
              let BeforeTotalResultAddressDetail = getRangeDetails(BeforeTotalResultAddress);
              let BeforeTotalResultRow = BeforeTotalResultAddressDetail.bottomRow;
              let BeforeTotalResultColumn = BeforeTotalResultAddressDetail.leftColumn;
              let ContributionFormula = `=(${CurrentMixAddress}-${BeforeMixAddress})*(${BeforeResultAddress}-\$${BeforeTotalResultColumn}\$${BeforeTotalResultRow})+${CurrentMixAddress}*(${CurrentResultAddress}-${BeforeResultAddress})`
              console.log("ContributionFormula is " + ContributionFormula);

              let ContributionFirstRow = ContributionStartCell.getOffsetRange(1,0); //往下一格放入公式
              ContributionFirstRow.formulas = [[ContributionFormula]];

              //---------给单元格设置格式，不然有可能是excel自动判断的格式-----------------------------
              let ResultType = FirstRow.find("Result", {
                completeMatch: true,
                matchCase: true,
                searchDirection: "Forward"
              });
              ResultType.load("address");
              await context.sync();
              //往下4行，获得Result数据单元格
              let ResultCell = ResultType.getOffsetRange(4, 0);
              // ResultCell.load("numberFormat"); // 获得单元格的数据格式
              // await context.sync();
              // 将数据格式应用到 Bridge 数据范围
              ContributionFirstRow.copyFrom(ResultCell, Excel.RangeCopyType.formats // 只复制格式
              );

              let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
              let TempRangeTitle = TempVarSheet.getRange("B24");
              let ResultFormat = TempVarSheet.getRange("B25");
              TempRangeTitle.values = [["Result Format"]];
              ResultFormat.copyFrom(ResultCell, Excel.RangeCopyType.formats);
              ResultFormat.copyFrom(ResultCell, Excel.RangeCopyType.values);
              await context.sync();
              //----------给单元格设置格式，不然有可能是excel自动判断的格式----------End-------------------

              let ContributionColumn = ContributionFirstRow.getAbsoluteResizedRange(ProcessRange.rowCount-1,1);//扩大到整个列
              ContributionColumn.copyFrom(ContributionFirstRow);

              ContributionStartCell = ContributionStartCell.getOffsetRange(0, 1); //往右移动一格
              console.log("Contribution X");
              await context.sync();
            }

          }


      }

    }else{   //如果不是除法而是可以加总的则直接把分母设置为1，Mix都一样都是平均的

      // 循环每个变量，计算出每一步变量变化对应的被除数的Mix
      let iColumn = 0;
      for (let z = 0; z < arr.length; z++) {
        let Title = arr[z][0];
        let Type = arr[z][1];
        console.log(`arr[${z}][1] is` + arr[z][1]);
        //TitleCell.load("address,values");
        //TitleType.load("address,values");
        //await context.sync();
        console.log("Enter Mix 1");
        // console.log("StrGlbDenominator is " + StrGlbDenominator);
        // let cellName = null;
        // 创建 parts 的副本，避免修改原数组

        // if (parts) {
        //   console.log("parts is " + JSON.stringify(parts));
        // } else {
        //   console.log("parts is undefined or null");
        // }

        // let TempParts = [...parts]; // 使用扩展运算符创建一个新的副本
        // console.log("TempParts is " + TempParts);

        if (!["Result", "ProcessSum", "Impact", "NULL", ""].includes(Type)) {

          console.log("TitleCell is " + Title);


          // // // 遍历数组parts 中的所有变量, 在process第三行中找到相应的单元格
          // // for (let i = 0; i < parts.length; i++) {
          // //   let variable = parts[i];
          // //   console.log("variable is " + variable);

          // //   // 只处理变量（忽略运算符和括号）
          // //   if (/[^+\-*/^()]+/.test(variable)) {  // 检测非运算符、非括号的变量

          // //     // 遍历 RangeA 查找所有匹配的单元格
          // //     for (let j = 0; j < ProcessAllRange.columnCount; j++) {
          // //       // let VarCell = ProcessAllRange.values[2][j];   // getCell(0, j); 
          // //       // VarCell.load("address,values"); //获取第三行的每个单元格
          // //       // await context.sync();

          // //       // 如果单元格的值等于当前变量名
          // //       if (ProcessAllRange.values[2][j] === variable && ProcessAllRange.values[1][j] === Title) {
          // //         console.log("variable2 is " + variable);
          // //         // console.log("VarCell is " + VarCell);
          // //         // 检查符合条件的单元格
          // //         // let upperCell = AllProcessFirstRow.values[0][j];

          // //         // 判断上方两行的单元格是否符合条件

          // //         // 查找上方一行是否等于变量名
          // //         // let oneRowUp = AllProcessSecondRow.values[0][j];
          // //         // let oneRowDown = AllProcessFourthdRow.values[0][j];

          // //         // console.log("oneRowUp is " + oneRowUp);
          // //         console.log("TitleCell is " + Title);
          // //         // if (oneRowUp === Title ) {
          // //         // 符合条件，使用该单元格
          // //         // console.log("oneRowUp is OK " + oneRowUp);
          // //         //cellName = VarCell;
          // //         // ProcessAllRangeAddress
          // //         // let cellAddress = FourthRowAddress[0][j].split('!')[1];
          // //         let cellAddress = ProcessAllRangeAddress[3][j].split('!')[1];
          // //         // 将变量替换为 Cell Var 的地址
          // //         TempParts[i] = cellAddress;
          // //         console.log("TempParts is " + TempParts);


          // //         break;  // 找到符合条件的单元格后退出循环
          // //         // }
          // //       }
          // //     }
          // //   }

          //   // 如果找到了符合条件的 cellName，继续处理
          //   //console.log("cellName is " + cellName);
          //   // if (cellName) {
          //   //   // 查找符合条件的Cell Var并往下移动一行
          //   //   let cellVar = cellName.getOffsetRange(1, 0);
          //   //   cellVar.load("address");
          //   //   await context.sync();

          //   //   // 获取 Cell Var 的地址，例如 A1 样式
          //   //   let cellAddress = cellVar.address.split('!')[1];

          //   //   // 将变量替换为 Cell Var 的地址
          //   //   TempParts[i] = cellAddress;
          //   //   console.log("TempParts is " + TempParts);
          //   // }
          // }

          let finalFormula = "=1";
          console.log("finalFormula is " + finalFormula);
          console.log(`${Title} End~!!!!`)
          // 重新创建 TempParts 的副本，避免影响下一次循环
          // TempParts = [...parts];

          //将替换好的公式放入Process对应的单元格，作为每一个变量替换后的分母
          // MixStartCell.values = [[Title]];
          //从第三行开始时标题
          MixArr[2][iColumn] = `="${Title}"`;
          // let MixFirstRow = MixStartCell.getOffsetRange(1, 0);
          //第四行开始是第一行带有公式formulas的单元格
          MixArr[3][iColumn] = finalFormula;
          // MixFirstRow.numberFormat = '#,##0.00'; // 设置数据格式，因为时计算出来的，无法复制前面的单元格
          // MixFirstRow.load('address');
          // let MixTwoUpRow = MixStartCell.getOffsetRange(-2, 0);
          MixArr[0][iColumn] = `="Denominator"`;
          // await context.sync();

          // let MixRange = MixFirstRow.getAbsoluteResizedRange(ProcessRange.rowCount - 1, 1); // Mix的一整列
          // MixRange.load("address");
          // // MixRange.copyFrom(MixFirstRow, Excel.RangeCopyType.formulas, false, false); 
          // MixRange.copyFrom(MixFirstRow, Excel.RangeCopyType.formulasAndNumberFormats, false, false);//将公式拷贝到一整行
          // await context.sync();
          //从第一个变量单元格开始往右移动，从第4行开始，Z列是计算denominator，z+1是计算Mix
          let DenominatorCellAddress = ProcessAllRangeAddress[3][MixStartColumnIndex + z*2];
          console.log("DenominatorCellAddress is " + DenominatorCellAddress);
          let DenominatorAddressDetail = getRangeDetails(DenominatorCellAddress);
          let DenominatorTopRow = DenominatorAddressDetail.topRow;
          let DenominatorColumn = DenominatorAddressDetail.leftColumn;
          let DenominatorBottom = ProcessBottomRow;   //getRangeDetails(MixRange.address).bottomRow;

          //计算Mix

          // MixStartCell = MixStartCell.getOffsetRange(0, 1); // 自身往右移动一格
          MixArr[2][iColumn + 1] = `="${Title}"`; // 往右移动一格
          // let MixTwoUpRow = MixStartCell.getOffsetRange(-2, 0);
          MixArr[0][iColumn + 1] = `="Mix"`;
          let MixFormula = `=${DenominatorColumn}${DenominatorTopRow}/\$${DenominatorColumn}\$${DenominatorBottom}`
          console.log("MixFormula is " + MixFormula);
          // MixFirstRow = MixFirstRow.getOffsetRange(0, 1);
          MixArr[3][iColumn + 1] = MixFormula;
          // 设置百分比格式并保留两位小数
          // MixFirstRow.numberFormat = '0.00%';
          // await context.sync();

          // MixRange = MixRange.getOffsetRange(0, 1);
          // MixRange.copyFrom(MixFirstRow, Excel.RangeCopyType.formulasAndNumberFormats, false, false);

          // MixStartCell = MixStartCell.getOffsetRange(0, 1); // 自身往右移动一格
          // await context.sync();

        }
        iColumn = iColumn + 2;
      }
      //console.log("test1")

      MixArr = normalizeArray(MixArr); // 补齐其中有的空行，使得列数一样，如果数组不对齐，则不能给单元格赋值fomulas      
      console.log("MixArr is");
      console.log(MixArr);

      // 获取行数
      let MixRowCount = MixArr.length;
      // 获取列数（假设所有行的列数相同）
      let MixColCount = MixArr[0].length;
      let InputMixStartCell = MixStartCell.getOffsetRange(-2, 0); //从第一行开始的单元格
      let InputMixRange = InputMixStartCell.getAbsoluteResizedRange(MixRowCount, MixColCount);
      InputMixRange.formulas = MixArr;

      // 复制第 4 行到第 5 行及之后的所有行
      let rowToCopy = InputMixRange.getRow(3); // 第 4 行
      let rangeToFill = InputMixRange.getRow(4).getOffsetRange(0, 0).getAbsoluteResizedRange(MixRowCount - 4, MixColCount); // 第 5 行到最后一行
      rangeToFill.copyFrom(rowToCopy, Excel.RangeCopyType.formulas); // 复制公式

      MixStartCell = MixStartCell.getOffsetRange(0, MixColCount); //移动到Denomination 和 Mix单元格之后
      await context.sync();
      console.log("Contribution 7.1");


      //开始整理格式
      InputMixRange.getRow(2).copyFrom(ProcessRangeRightTop, Excel.RangeCopyType.formats); //复制标题格式
      InputMixRange.getRow(MixRowCount - 1).copyFrom(ProcessRangeRightBottom, Excel.RangeCopyType.formats); //复制汇总格式

      // 获取从第 4 行开始的范围
      const rangeFromFourthRow = InputMixRange.getRow(3).getAbsoluteResizedRange(MixRowCount - 3, MixColCount);

      // 遍历每一列 设置数据格式
      for (let colIndex = 0; colIndex < MixColCount; colIndex++) {
        const columnRange = rangeFromFourthRow.getColumn(colIndex);

        if ((colIndex + 1) % 2 === 1) {
          // 单数列（1, 3, 5, ...）
          columnRange.numberFormat = '#,##0.00';
        } else {
          // 双数列（2, 4, 6, ...）
          columnRange.numberFormat = '0.00%';
        }
      }

      await context.sync();
      console.log("Contribution 7.2");

      //--------------------------------------------计算contribution-------------------------------------------
      ContributionStartCell = MixStartCell.getOffsetRange(0, 1) // 往右移动一格
      let NewUsedRange = ProcessSheet.getUsedRange(); // 这里的UsedRange是Process 工作表的更新后的适用范围
      let FirstRow = NewUsedRange.getRow(0);
      let SecondRow = NewUsedRange.getRow(1);
      let ThirdRow = NewUsedRange.getRow(2);
      let FourthRow = NewUsedRange.getRow(3);
      let BottomRow = NewUsedRange.getRow(UsedRange.rowCount - 1); //这里可以使用上一步的UsedRange，省去一次sync
      NewUsedRange.load("address,values")
      FirstRow.load("address,values");
      SecondRow.load("address,values");
      ThirdRow.load("address,values");
      FourthRow.load("address,values");
      BottomRow.load("address,values");
      await context.sync();

      // let ProcessFirstRowAddress = await GetRangeAddress("Process",FirstRow.address);
      // let ProcessSecondRowAddress = await GetRangeAddress("Process",SecondRow.address);
      // let ProcessThirdRowAddress = await GetRangeAddress("Process",ThirdRow.address);
      let ProcessFourthRowAddress = await GetRangeAddress("Process", FourthRow.address);
      let ProcessBottomRowAddress = await GetRangeAddress("Process", BottomRow.address);
      console.log("NewUsedRange is " + NewUsedRange.address);
      //console.log("test2")
      //不循环BasePT 和 Target PT，因此z =1, arr.length -1
      for (let z = 1; z < arr.length - 1; z++) {
        console.log("Current Var is" + arr[z][0]);
        //在第一行找到Mix 以及对应的当前变量
        // let CurrentMixTitle = null;
        // let BeforeMixTitle = null;
        // let CurrentResultCell = null;
        // let BeforeResultCell = null;
        // let BeforeTotalResultCell = null;

        // let CurrentMixAddress = null;
        let CurrentResultAddress = null;
        let BeforeMixAddress = null;
        let BeforeResultAddress = null;
        let BeforeTotalResultAddress = null;

        //console.log("FirstRowValues length is " + FirstRow.values[0].length)
        //下面必须是FirstRowValues[0].length，而不能是FirstRowValues.length，这样length是1，因为只有一行
        for (let i = 0; i < FirstRow.values[0].length - 1; i++) {
          //console.log("test4")

          //找到当前变量对应的Result的相关信息
          if (FirstRow.values[0][i] === "Result" && arr[z][0] === SecondRow.values[0][i]) {
            // CurrentResultCell = SecondRow.getCell(0,i).getOffsetRange(2,0); //获取下面两格，其中的包含Result结果单元格
            // CurrentResultCell.load("address,values");
            // await context.sync();
            console.log("ProcessFourthRowAddress[0][i] is " + ProcessFourthRowAddress[0][i]);
            // CurrentResultAddress = CurrentResultCell.address.split("!")[1]; //获取Current地址
            CurrentResultAddress = ProcessFourthRowAddress[0][i].split("!")[1];
            console.log("CurrentResultAddress is " + CurrentResultAddress)
          }

          //找到前一个Mix的相关信息,这里需要是arr[z-1][0]
          //for (let j = 0; j < FirstRow.values[0].length - 1; j++) {
          // console.log("FirstRow.values[0][i] is " + FirstRow.values[0][i]);
          // console.log("ThirdRow.values[0][i] is " + ThirdRow.values[0][i]);
          if (FirstRow.values[0][i] === "Mix" && arr[z - 1][0] === ThirdRow.values[0][i]) {
            // let BeforeMixTitle = ThirdRow.getCell(0,i);
            // let BeforeMixCell = BeforeMixTitle.getOffsetRange(1, 0);//往下移动一格，找到带有值的Mix
            // BeforeMixTitle.load("address,values");
            // BeforeMixCell.load("address,values");
            // await context.sync();
            console.log("ProcessFourthRowAddress[0][i] 2 is " + ProcessFourthRowAddress[0][i]);
            BeforeMixAddress = ProcessFourthRowAddress[0][i].split("!")[1]; //获取单元格Mix地址A1等
            console.log("BeforeMixAddress is " + BeforeMixAddress);

            //}

          }

          //找到前一个变量对应的Result的相关信息，这里需要arr[z-1][0]
          if (FirstRow.values[0][i] === "Result" && arr[z - 1][0] === SecondRow.values[0][i]) {
            // BeforeResultCell = SecondRow.getCell(0, i).getOffsetRange(2, 0); //获取下面两格，其中的包含Result结果单元格
            // BeforeTotalResultCell = SecondRow.getCell(0, i).getOffsetRange(ProcessRange.rowCount, 0); //获取最下面一行的Total Result
            // BeforeResultCell.load("address,values");
            // BeforeTotalResultCell.load("address,values");
            // await context.sync();

            BeforeResultAddress = ProcessFourthRowAddress[0][i].split("!")[1]; //获取Current地址
            BeforeTotalResultAddress = ProcessBottomRowAddress[0][i].split("!")[1]; //获取Current地址
            console.log("BeforeResultAddress is " + BeforeResultAddress);
            console.log("BeforeTotalResultAddress is " + BeforeTotalResultAddress);
          }

          //如果第一行是Mix, 并且第三行等于数组中的变量，则i就是对应的列
          //执行到这一步，上面的if应该已经把contribution的公式变量都找到了
          if (FirstRow.values[0][i] === "Mix" && arr[z][0] === ThirdRow.values[0][i]) {
            //console.log("FirstRow.values[0][i] is " + FirstRow.values[0][i]);
            //console.log("ThirdRow.values[0][i] is " + ThirdRow.values[0][i]);
            //console.log("I is " + i);
            let CurrentMixTitle = ThirdRow.getCell(0, i); //找到对应的第三行的标题
            let CurrentMixCell = CurrentMixTitle.getOffsetRange(1, 0); //往下移动一格，找到带有值的Mix 
            let CurrentType = ContributionStartCell.getOffsetRange(-2, 0); //获contribution取标题单元格
            CurrentType.values = [["Contribution"]];
            ContributionStartCell.copyFrom(CurrentMixTitle); //复制标题
            CurrentMixTitle.load("address,values");
            CurrentMixCell.load("address,values");
            await context.sync();

            let CurrentMixAddress = CurrentMixCell.address.split("!")[1]; //获取单元格Mix 地址 A1等
            console.log("CurrentMixAddress is " + CurrentMixAddress);

            //找到了全部变量，开始生成公式
            let BeforeTotalResultAddressDetail = getRangeDetails(BeforeTotalResultAddress);
            let BeforeTotalResultRow = BeforeTotalResultAddressDetail.bottomRow;
            let BeforeTotalResultColumn = BeforeTotalResultAddressDetail.leftColumn;
            let ContributionFormula = `=(${CurrentMixAddress}-${BeforeMixAddress})*(${BeforeResultAddress}-\$${BeforeTotalResultColumn}\$${BeforeTotalResultRow})+${CurrentMixAddress}*(${CurrentResultAddress}-${BeforeResultAddress})`
            console.log("ContributionFormula is " + ContributionFormula);

            let ContributionFirstRow = ContributionStartCell.getOffsetRange(1, 0); //往下一格放入公式
            ContributionFirstRow.formulas = [[ContributionFormula]];

            //---------给单元格设置格式，不然有可能是excel自动判断的格式-----------------------------
            let ResultType = FirstRow.find("Result", {
              completeMatch: true,
              matchCase: true,
              searchDirection: "Forward"
            });
            ResultType.load("address");
            await context.sync();
            //往下4行，获得Result数据单元格
            let ResultCell = ResultType.getOffsetRange(4, 0);
            ResultCell.load("numberFormat"); // 获得单元格的数据格式

            // 将数据格式应用到 Bridge 数据范围
            ContributionFirstRow.copyFrom(ResultCell, Excel.RangeCopyType.formats // 只复制格式
            );

            let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
            let TempRangeTitle = TempVarSheet.getRange("B24");
            let ResultFormat = TempVarSheet.getRange("B25");
            TempRangeTitle.values = [["Result Format"]];
            ResultFormat.copyFrom(ResultCell,Excel.RangeCopyType.formats);
            ResultFormat.copyFrom(ResultCell,Excel.RangeCopyType.values);


            //----------给单元格设置格式，不然有可能是excel自动判断的格式----------End-------------------

            let ContributionColumn = ContributionFirstRow.getAbsoluteResizedRange(ProcessRange.rowCount - 1, 1);//扩大到整个列
            ContributionColumn.copyFrom(ContributionFirstRow);

            ContributionStartCell = ContributionStartCell.getOffsetRange(0, 1); //往右移动一格
            console.log("Contribution X");
            await context.sync();
          }

        }

      }

    }

    // 以上计算Contribution 是否是除法的两种清空结束后，把Contribution 结束最右列再移动了一列的地址保存在全局变量中
    ContributionStartCell.load("address");
    await context.sync();

    ContributionEndCellAddress = ContributionStartCell.address;
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    let ContributionEndName = TempVarSheet.getRange("B18");
    ContributionEndName.values =[["ContributionEnd"]];
    let ContributionEnd = TempVarSheet.getRange("B19");
    ContributionEnd.values = [[ContributionEndCellAddress]];
    await context.sync();
  });
}

//创建临时储存变量的工作表
async function CreateTempVar() {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    // 检查是否存在同名的工作表
    let BridgeSheet = workbook.worksheets.getItemOrNullObject("TempVar");
    await context.sync();

    if (BridgeSheet.isNullObject) {
      // 工作表不存在，创建新工作表
      BridgeSheet = context.workbook.worksheets.add("TempVar");
      await context.sync();
      console.log("创建了新工作表：" + "TempVar");
      let range = BridgeSheet.getRange("A1");
      range.values = [["TempVar"]];
      await context.sync();
    }

      await DoNotChangeCellWarning("TempVar");
  });
}

//获取Bridge Data表格的数据格式 *********右边扩展的列需要设定格式，不然太乱******
async function getBridgeDataFormats() {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const sheet = workbook.worksheets.getItem("Bridge Data");

    // 获取第二行的标题 (第二行假设为 2 行)
    let range = sheet.getUsedRange();
    const secondRowRange = range.getRow(1);
    secondRowRange.load("values");

    // 获取第三行的数据 (第三行假设为 3 行)
    const thirdRowRange = range.getRow(2);
    thirdRowRange.load("numberFormat");

    await context.sync(); // 确保已加载行数据
    console.log("secondRowRange is " + secondRowRange.values);
    // 创建一个对象来保存标题和数据格式
    let titleFormatMapping = {};

    // 获取第二行的标题和第三行的数据格式
    const titles = secondRowRange.values[0];
    const formats = thirdRowRange.numberFormat[0];

    // 将标题和相应的数据格式放入对象中
    for (let i = 0; i < titles.length; i++) {
      const title = titles[i];
      const format = formats[i];

      if (title) { // 确保标题存在
        titleFormatMapping[title] = format;
      }
    }

    console.log(titleFormatMapping);
    return titleFormatMapping;
  }).catch(function (error) {
    console.error("Error: ", error);
  });
}


//-----------------控制警告提示出现在最开始的地方------------------
async function showWarning() {
  const warningPrompt = document.getElementById('warningPrompt');
  const modalOverlay = document.getElementById("modalOverlay");
  const container = document.querySelector(".container");

  // 显示模态遮罩和提示框
  modalOverlay.style.display = "block";
  warningPrompt.style.display = "flex";
  container.classList.add("disabled");

}

async function hideWarning() {
  const warningPrompt = document.getElementById('warningPrompt');
  const modalOverlay = document.getElementById('modalOverlay');
  const container = document.querySelector('.container');

  // 隐藏提示框和模态遮罩
  warningPrompt.style.display = 'none';
  modalOverlay.style.display = 'none';

  // 恢复容器的交互
  container.classList.remove('disabled');
}

document.getElementById('confirmWarningPrompt').addEventListener('click', () => {
  hideWarning();
});
//-----------------控制警告提示出现在最开始的地方 END------------------


//---------------------------隐藏并保护多个工作表---------------------------------
async function disableScreenUpdating(context) {
  context.application.suspendApiCalculationUntilNextSync();
  context.application.suspendScreenUpdatingUntilNextSync();
  await context.sync(); // 确保挂起操作同步完成
}

async function enableScreenUpdating(context) {
  // 使用 Excel 的替代方法手动恢复计算和屏幕更新
  context.application.calculate(Excel.CalculationType.full); // 重新计算以确保一致性
  await context.sync(); // 确保恢复操作同步完成
}

async function protectSheets(context, sheetNames) {
  sheetNames.forEach(sheetName => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.protection.protect(); // 保护工作表以防止修改
  });
  await context.sync();
}

async function unprotectSheets(context, sheetNames) {
  sheetNames.forEach(sheetName => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.protection.unprotect(); // 取消保护工作表
  });
  await context.sync();
}

async function hideSheets(context, sheetNames) {
  sheetNames.forEach(sheetName => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.visibility = Excel.SheetVisibility.hidden; // 隐藏工作表以防止用户操作
  });
  await context.sync();
}

async function unhideSheets(context, sheetNames) {
  sheetNames.forEach(sheetName => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.visibility = Excel.SheetVisibility.visible; // 取消隐藏工作表
  });
  await context.sync();
}

//---------------------------隐藏并保护多个工作表 END---------------------------------


//建立用户使用的Contribution Table
async function CreateContributionTable() {
  await Excel.run(async (context) => {

    // 获取Process 中 contribution的单元格地址
    const ProcessSheet = context.workbook.worksheets.getItem("Process");
    let ProcessUsedRange = ProcessSheet.getUsedRange();
    ProcessUsedRange.load("address");
    await context.sync();

    console.log("ProcessUsedRange is" + ProcessUsedRange.address);

    let BottomRow = getRangeDetails(ProcessUsedRange.address).bottomRow;
    let TitleRange = ProcessSheet.getRange(`B3:B${BottomRow}`);
    TitleRange.load("address");
    await context.sync();
    console.log("TitleRange is " + TitleRange.address);
    console.log("BottomRow is " + BottomRow);

    let ContriAddress = await findContributionCells();
    console.log("ContriAddress is ");
    console.log(ContriAddress);

    let ContriLeftColumn = getRangeDetails(ContriAddress.leftCell).leftColumn;
    let ContriRightColumn = getRangeDetails(ContriAddress.rightCell).rightColumn;
    let ContributionRange = ProcessSheet.getRange(`${ContriLeftColumn}3:${ContriRightColumn}${BottomRow}`);
    ContributionRange.load("address,rowCount,columnCount");
    await context.sync();

    console.log("ContributionRange is " + ContributionRange.address);
    console.log("Row is" + ContributionRange.rowCount);
    console.log("Column is " + ContributionRange.columnCount);

    // 在Waterfall 表格中找到UsedRange的左下角单元格
    let WaterfallSheet = context.workbook.worksheets.getItem("Waterfall");
    // let WaterfallUsedRange = WaterfallSheet.getUsedRange();
    // WaterfallUsedRange.load("address");
    // await context.sync();
    // console.log("Waterfall used range is " + WaterfallUsedRange.address);

    // let WaterfallLeftColumn = getRangeDetails(WaterfallUsedRange.address).leftColumn;
    // let WaterfallBottomRow = getRangeDetails(WaterfallUsedRange.address).bottomRow;
    // let WaterfallLeftBottomCell = WaterfallSheet.getRange(`${WaterfallLeftColumn}${WaterfallBottomRow}`);
    // WaterfallLeftBottomCell.load("address");
    // await context.sync();

    // console.log("WaterfallleftBottom is " + WaterfallLeftBottomCell.address);

    //将Process Contribution 的Title拷贝到Waterfall 工作表
    // let ContributionTitleStart = WaterfallLeftBottomCell.getCell(3, 0); //往下移动3格，作为起始格子，可以根据需要变动
    let ContributionTitle = WaterfallSheet.getRange("I24"); //固定到Waterfall图表的下方
    ContributionTitle.values =[["Contribution Analysis"]];
    let ContributionTitleStart = WaterfallSheet.getRange("I25"); //固定到Waterfall图表的下方
    ContributionTitleStart.copyFrom(TitleRange, Excel.RangeCopyType.formats);
    ContributionTitleStart.copyFrom(TitleRange, Excel.RangeCopyType.values);
    ContributionTitleStart.load("address");
    let ContributionTitleRange = ContributionTitleStart.getAbsoluteResizedRange(ContributionRange.rowCount,1); // Title的列对应的Range
    ContributionTitleRange.load("address");
    await context.sync();



    //将Contribution表格的起始地址放入TempVar表格中,供Link使用
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    let ContributeionVarName = TempVarSheet.getRange("B9");
    ContributeionVarName.values = [["ContriAddress"]];
    let ContributionTitleStartVar = TempVarSheet.getRange("B10");
    ContributionTitleStartVar.values = [[ContributionTitleStart.address]];

    await context.sync();

    //将Process Contribution 的数据拷贝到Waterfall 工作表
    let ContributionTableStart = ContributionTitleStart.getCell(0,1); //往右移动一列
    ContributionTableStart.load("address");
    ContributionTableStart.copyFrom(ContributionRange, Excel.RangeCopyType.formats);
    ContributionTableStart.copyFrom(ContributionRange, Excel.RangeCopyType.values);
    //Contribution 数据范围
    let ContributionTableRange = ContributionTableStart.getAbsoluteResizedRange(ContributionRange.rowCount, ContributionRange.columnCount);
    let ContributionTableFirstRow = ContributionTableRange.getRow(0); //获取表格的第一行
    ContributionTableRange.load("address");
    await context.sync();

    console.log("ContributionTableStart is " + ContributionTableStart.address);
    console.log("ContributionTableRange is " + ContributionTableRange.address);

    //Title 和 Contribution 数据的范围合计，设置表格格式
    let ContributionTitleStartAddress = getRangeDetails(ContributionTitleStart.address);
    let ContriLeft = ContributionTitleStartAddress.leftColumn;
    let ContriTop = ContributionTitleStartAddress.topRow;
    let ContributionTableRangeAddress = getRangeDetails(ContributionTableRange.address);
    let ContriRight = ContributionTableRangeAddress.rightColumn;
    let ContriBottom = ContributionTableRangeAddress.bottomRow;
    // console.log(ContriLeft);
    // console.log(ContriTop);
    // console.log(ContriRight);
    // console.log(ContriBottom);
    let ContriTableAllRange = WaterfallSheet.getRange(`${ContriLeft}${ContriTop}:${ContriRight}${ContriBottom}`);
    ContriTableAllRange.load("address");
    await context.sync();
    //将Contribution Key的Range放入TempVar表格中，供Variance表格使用
    let ContributionName = TempVarSheet.getRange("B15");
    ContributionName.values = [["ContributionName"]];
    let ContributionForVariance = TempVarSheet.getRange("B16");
    ContributionForVariance .values = [[ContriTableAllRange.address]];

    // await context.sync();
    let ContriTableFirstRow = ContriTableAllRange.getRow(0); // 第一行
    let ContriTableLastRow = ContriTableAllRange.getLastRow(); //最后一行
    await context.sync();

    // 清除第一行的所有边框
    ContriTableFirstRow.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.none;
    ContriTableFirstRow.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.none;
    ContriTableFirstRow.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.none;
    ContriTableFirstRow.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.none;

    // 设置第一行的背景颜色为淡蓝色
    ContriTableFirstRow.format.fill.color = "#DDEBF7";; // 淡蓝色

    // 设置第一行的字体为粗体
    ContriTableFirstRow.format.font.bold = true;

    // 清除最后一行的所有边框
    ContriTableLastRow.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.none;
    ContriTableLastRow.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.none;
    ContriTableLastRow.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.none;
    ContriTableLastRow.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.none;

    // 设置最后一行的背景颜色为淡蓝色
    ContriTableLastRow.format.fill.color = "#DDEBF7"; // 淡蓝色

    // 设置最后一行的字体为粗体
    ContriTableLastRow.format.font.bold = true;
    
    //表格加上外边框
    ContriTableAllRange.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.continuous;
    ContriTableAllRange.format.borders.getItem('EdgeTop').weight = Excel.BorderWeight.thin;
    ContriTableAllRange.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.continuous;
    ContriTableAllRange.format.borders.getItem('EdgeBottom').weight = Excel.BorderWeight.thin;
    ContriTableAllRange.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.continuous;
    ContriTableAllRange.format.borders.getItem('EdgeLeft').weight = Excel.BorderWeight.thin;
    ContriTableAllRange.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.continuous;
    ContriTableAllRange.format.borders.getItem('EdgeRight').weight = Excel.BorderWeight.thin;

    // 自动调整整个表格的列宽
    ContriTableAllRange.format.autofitColumns();

    // 设置第一行的文本对齐格式为自动换行，并且上下左右居中
    ContriTableFirstRow.format.wrapText = true;
    // ContriTableFirstRow.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    // ContriTableFirstRow.format.verticalAlignment = Excel.VerticalAlignment.center;
    //数据部分全部居中对齐
    ContributionTableRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    ContributionTableRange.format.verticalAlignment = Excel.VerticalAlignment.center;
    // 设置最后一行的文本对齐格式为自动换行，并且上下左右居中
    ContriTableLastRow.format.wrapText = true;
    // ContriTableLastRow.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    // ContriTableLastRow.format.verticalAlignment = Excel.VerticalAlignment.center;
    // Title 靠左对齐
    ContributionTitleRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
    ContributionTitleRange.format.verticalAlignment = Excel.VerticalAlignment.center;

    await context.sync();

    await insertHyperlink("Contribution", "Waterfall", "C1"); //设置Contributioin Link

    await WaterfallVarianceTable(); //创建用户使用的VarianceTable
  });
}


//找到Process中第一行的Contribution的Range
async function findContributionCells() {
  try {
    return await Excel.run(async (context) => {
      // 获取工作表“Process”
      const sheet = context.workbook.worksheets.getItem("Process");
      // 获取第一行的范围
      let UsedRange = sheet.getUsedRange();
      await context.sync();
      let range = UsedRange.getRow(0);

      // 加载单元格的值和地址
      range.load("values, address");

      // 同步上下文
      await context.sync();

      let rangeAddress = await GetRangeAddress("Process",range.address);

      let leftCell = null; // 最左边的“Contribution”单元格地址
      let rightCell = null; // 最右边的“Contribution”单元格地址

      // 遍历第一行的所有单元格
      for (let i = 0; i < range.values[0].length; i++) {
        //console.log("range.values[0][i] is " + range.values[0][i]);
        // 如果单元格的值为“Contribution”
        if (range.values[0][i] === "Contribution") {
          // 如果leftCell为空，说明这是第一个找到的“Contribution”单元格
          //console.log("step1");
          // let ContriCell = range.getCell(0, i);
          // ContriCell.load("address");
          // //console.log("step2");
          // await context.sync();

          if (leftCell === null) {
            // leftCell = ContriCell.address;
            leftCell = rangeAddress[0][i];
          }
          // 更新rightCell为当前单元格地址
          // rightCell = ContriCell.address;
          rightCell = rangeAddress[0][i];
        }
      }

      // 如果找到了“Contribution”单元格
      if (leftCell && rightCell) {
        console.log(`Leftmost Contribution cell: ${leftCell}`);
        console.log(`Rightmost Contribution cell: ${rightCell}`);
        return { leftCell, rightCell };
      } else {
        // 如果没有找到“Contribution”单元格
        console.log("No Contribution cells found.");
        return null;
      }
    });
  } catch (error) {
    // 捕获并输出错误
    console.error(error);
  }
}

//画出Bridge图形
async function DrawBridge_onlyChart() {
  await Excel.run(async (context) => {

    // let BridgeRangeAddress = await BridgeCreate();  // 创建waterfall工作表，生成Bridge数据，并返回相对应的单元格，仅包含字段名和impact两列
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    let BridgeRangeVar = TempVarSheet.getRange("B6");
    BridgeRangeVar.load("values");
    await context.sync();

    let BridgeRangeAddress = BridgeRangeVar.values[0][0];

    console.log("BridgeRangeAddress is " + BridgeRangeAddress);
    // BridgeDataFormatAddress = BridgeRangeAddress; // 传递给全局函数
    
    // 获取名为 "Waterfall" 的工作表
    let sheet = context.workbook.worksheets.getItem("Waterfall");
    // 获取 Bridge 数据的范围
    let BridgeRange = sheet.getRange(BridgeRangeAddress);
    //let BridgeRange = sheet.getRange(BridgeRangeAddress);
    
    BridgeRange.load("address,values,rowCount,columnCount");
    await context.sync();

    let StartRange = BridgeRange.getCell(0, 0);
    let dataRange = StartRange.getOffsetRange(0, 2).getAbsoluteResizedRange(BridgeRange.rowCount, 4);
    //图形的数据范围
    let xAxisRange = StartRange.getAbsoluteResizedRange(BridgeRange.rowCount, 1); // 横轴标签范围
    let BlankRange = StartRange.getOffsetRange(0, 2).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let GreenRange = StartRange.getOffsetRange(0, 3).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let RedRange = StartRange.getOffsetRange(0, 4).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    let AccRange = StartRange.getOffsetRange(0, 5).getAbsoluteResizedRange(BridgeRange.rowCount, 1); //辅助列
    let BridgeDataRange = StartRange.getOffsetRange(0, 1).getAbsoluteResizedRange(BridgeRange.rowCount, 1);
    // let BridgeFormats = StartRange.getOffsetRange(0,1).getAbsoluteResizedRange(BridgeRange.rowCount,5); //全部数据的范围，需要调整格式

    // 加载数据范围和横轴标签
    dataRange.load("address,values,rowCount,columnCount");
    xAxisRange.load("address,values,rowCount,columnCount");
    BlankRange.load("address,values,rowCount,columnCount");
    GreenRange.load("address,values,rowCount,columnCount");
    RedRange.load("address,values,rowCount,columnCount");
    AccRange.load("address,values,rowCount,columnCount");
    BridgeDataRange.load("address,values,rowCount,columnCount");
    console.log("DrawBridge 0");

    //寻找BridgeDate sheet第一行带有Result的单元格
    let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
    let BridgeDataSheetRange = BridgeDataSheet.getUsedRange();
    let BridgeDataSheetFirstRow = BridgeDataSheetRange.getRow(0);
    //await context.sync();
    console.log("DrawBridge_onlyChart 1");

    // 找到result单元格
    let ResultType = BridgeDataSheetFirstRow.find("Result", {
      completeMatch: true,
      matchCase: true,
      searchDirection: "Forward"
    });
    ResultType.load("address");
    await context.sync();
    console.log("DrawBridge 2")

    //往下两行，获得Result数据单元格
    let ResultCell = ResultType.getOffsetRange(2, 0);
    ResultCell.load("numberFormat"); // 获得单元格的数据格式

    // 将数据格式应用到 Bridge 数据范围
    // BridgeFormats.copyFrom(
    //   ResultCell,
    //   Excel.RangeCopyType.formats // 只复制格式
    // );

    await context.sync();

    console.log("ResultCell Formats is " + ResultCell.numberFormat[0][0]);
    console.log("dataRange is ", dataRange.address);
    console.log("xAxisRange is ", xAxisRange.address);
    console.log("BaseRange is ", BlankRange.address);
    console.log("GreenRange is ", GreenRange.address);
    console.log("RedRange is ", RedRange.address);
    console.log("AccRange is ", AccRange.address);

    //设置每个单元格的公式
    // BlankRange.getCell(0, 0).formulas = [["=C3"]];
    // BlankRange.getCell(0, 0)
    //   .getOffsetRange(BridgeRange.rowCount - 1, 0)
    //   .copyFrom(BlankRange.getCell(0, 0));
    // BlankRange.getCell(1, 0).formulas = [
    //   ["=IF(AND(G4<0,G3>0),G4,IF(AND(G4<0,G3<0,C4<0),G4-C4,IF(AND(G4<0,G3<0,C4>0),G3+C4,SUM(C$3:C3)-F4)))"]
    // ];
    // BlankRange.getCell(0, 0)
    //   .getOffsetRange(1, 0)
    //   .getAbsoluteResizedRange(BridgeRange.rowCount - 2, 1)
    //   .copyFrom(BlankRange.getCell(1, 0));

    // AccRange.getCell(0, 0).formulas = [["=SUM($C$3:C3)"]];
    // AccRange.getCell(0, 0)
    //   .getAbsoluteResizedRange(BridgeRange.rowCount - 1, 1)
    //   .copyFrom(AccRange.getCell(0, 0));
    // AccRange.getCell(BridgeRange.rowCount - 1, 0).copyFrom(BlankRange.getCell(BridgeRange.rowCount - 1, 0), Excel.RangeCopyType.values);

    // GreenRange.getCell(0, 0).getOffsetRange(1, 0).formulas = [
    //   ["=IF(AND(G3<0,G4<0,C4>0),-C4,IF(AND(G3<0,G4>0,C4>0),C4+D4,IF(C4>0,C4,0)))"]
    // ];
    // GreenRange.getCell(0, 0)
    //   .getOffsetRange(1, 0)
    //   .getAbsoluteResizedRange(BridgeRange.rowCount - 2, 1)
    //   .copyFrom(GreenRange.getCell(0, 0).getOffsetRange(1, 0));
    // RedRange.getCell(0, 0).getOffsetRange(1, 0).formulas = [
    //   ["=IF(AND(G3>0,G4<0,C4<0),D3,IF(AND(G3<0,G4<0,C4<0),C4,IF(C4>0,0,-C4)))"]
    // ];
    // RedRange.getCell(0, 0)
    //   .getOffsetRange(1, 0)
    //   .getAbsoluteResizedRange(BridgeRange.rowCount - 2, 1)
    //   .copyFrom(RedRange.getCell(0, 0).getOffsetRange(1, 0));

    // 删除已有的图表，避免重复创建
    let charts = sheet.charts;
    console.log("DrawBridge_onlyChart 2.1")
    charts.load("items/name");
    await context.sync();

    // 检查并删除名为 "BridgeChart" 的图表（如果存在）
    for (let i = 0; i < charts.items.length; i++) {
      if (charts.items[i].name === "BridgeChart") {
        charts.items[i].delete();
        break;
      }
    }
    console.log("DrawBridge_onlyChart 2.2")
    // 插入组合图表（柱状图和折线图）
    let chart = sheet.charts.add(Excel.ChartType.columnStacked, dataRange, Excel.ChartSeriesBy.columns);
    chart.name = "BridgeChart"; // 设置图表名称，便于后续查找和删除
    
        // 隐藏图表图例
    chart.legend.visible = false;

    // 定义目标单元格位置（例如 D5）

    // 设置图表位置，左上角对应单元格
    chart.setPosition("B12");


    // 设置图表的位置和大小
    // chart.top = 50;
    // chart.left = 50;
    chart.width = 500;
    chart.height = 300;

    await context.sync();

    // 设置横轴标签
    chart.axes.categoryAxis.setCategoryNames(xAxisRange.values);

    // 将轴标签位置设置为底部
    //chart.axes.valueAxis.position = "Automatic"; // 这里设置为Minimun 也只能在0轴的位置，不能是最低的负值下方
    let valueAxis = chart.axes.valueAxis;
    valueAxis.load("minimum");
    await context.sync();
    chart.axes.valueAxis.setPositionAt(valueAxis.minimum);

    // 获取图表的数据系列
    
    const seriesD = chart.series.getItemAt(0); // Base列
    const seriesE = chart.series.getItemAt(1); // 获取Green列的数据系列
    const seriesF = chart.series.getItemAt(2); // 获取Red列的数据系列
    const seriesLine = chart.series.getItemAt(3); // Bridge列

    seriesLine.chartType = Excel.ChartType.line; //插入Line
    //seriesLine.dataLabels.showValue = true;
    // 设置线条颜色为透明
    //seriesLine.format.line.color = "blue" ;
    seriesLine.format.line.lineStyle  = "None";

    seriesLine.points.load("count"); //这一步必须

    await context.sync();

    //设置线条的各种数据标签的颜色和位置等
    for (let i = 0; i < seriesLine.points.count; i++) {
      let CurrentBridgeRange = BridgeDataRange.getCell(i, 0);
      CurrentBridgeRange.load("values,text");
      await context.sync();
      //seriesLine.points.getItemAt(i).dataLabel.text = String(CurrentBridgeRange.values[0][0]);
      
      if (i == 0 || i == seriesLine.points.count -1){
        seriesLine.points.getItemAt(i).dataLabel.text = CurrentBridgeRange.text[0][0];
        seriesLine.points.getItemAt(i).dataLabel.numberFormat = ResultCell.numberFormat[0][0]; //设置数据格式
        seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#0070C0"  // 蓝色
        if(CurrentBridgeRange.values[0][0] >= 0){
          seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.top;
        }else{
          seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.bottom;
        }
        
      }else if (CurrentBridgeRange.values[0][0] > 0) {
        seriesLine.points.getItemAt(i).dataLabel.text = CurrentBridgeRange.text[0][0];
        seriesLine.points.getItemAt(i).dataLabel.numberFormat = ResultCell.numberFormat[0][0]; //设置数据格式
        seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#00B050"  //绿色
        seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.top;
      } else if (CurrentBridgeRange.values[0][0] < 0) {
        seriesLine.points.getItemAt(i).dataLabel.text = CurrentBridgeRange.text[0][0];
        seriesLine.points.getItemAt(i).dataLabel.numberFormat = ResultCell.numberFormat[0][0]; //设置数据格式
        seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#FF0000" //红色
        seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.bottom;
      } else {
        // seriesLine.points.getItemAt(i).dataLabel.format.font.color = "#000000"  //黑色
        // seriesLine.points.getItemAt(i).dataLabel.position = Excel.ChartDataLabelPosition.top;
      }
    }


    seriesD.points.load("items");
    seriesE.points.load("items");
    seriesF.points.load("items");

    await context.sync();

    // 为 D 列的数据点设置填充颜色
    for (let i = 0; i < seriesD.points.items.length; i++) {
      let BeforeAccRange = AccRange.getCell(i - 1, 0);
      let CurrentAccRange = AccRange.getCell(i, 0);
      BeforeAccRange.load("values");
      CurrentAccRange.load("values");

      await context.sync();

      if (i == 0 || i == seriesD.points.items.length - 1) {
        seriesD.points.items[i].format.fill.setSolidColor("#0070C0"); // 设置为起始和终点颜色
        //seriesD.points.items[i].dataLabel.showValue = true;
        //seriesD.points.items[i].dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
      } else if (i > 0 && BeforeAccRange.values[0][0] > 0 && CurrentAccRange.values[0][0] < 0) {
        seriesD.points.items[i].format.fill.setSolidColor("#FF0000"); // 设置为红色
      } else if (i > 0 && BeforeAccRange.values[0][0] < 0 && CurrentAccRange.values[0][0] > 0) {
        seriesD.points.items[i].format.fill.setSolidColor("#00B050"); // 设置为绿色
      } else {
        seriesD.points.items[i].format.fill.clear(); // 设置为无填充
      }
    }

    //seriesE.dataLabels.showValue = true;
    //seriesE.dataLabels.position = Excel.ChartDataLabelPosition.insideBase ;

    await context.sync();
    // 为E列数据点设置绿色
    for (let i = 0; i < seriesE.points.items.length; i++) {
      let CurrentGreenRange = GreenRange.getCell(i, 0);
      CurrentGreenRange.load("values");
      await context.sync();

      seriesE.points.items[i].format.fill.setSolidColor("#00B050");
      if (CurrentGreenRange.values[0][0] !== 0) {
        //seriesE.points.items[i].dataLabel.showValue = true;
        //seriesE.points.items[i].dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
      }
    }

    // 为F列数据点设置红色
    for (let i = 0; i < seriesF.points.items.length; i++) {
      let CurrentRedRange = RedRange.getCell(i, 0);
      CurrentRedRange.load("values");
      await context.sync();

      seriesF.points.items[i].format.fill.setSolidColor("#FF0000");
      if (CurrentRedRange.values[0][0] !== 0) {
        //seriesF.points.items[i].dataLabel.showValue = true;
        //seriesF.points.items[i].dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
      }
    }
    activateWaterfallSheet(); // 最后需要active waterfall 这个工作表
    
    await context.sync();
  });
}


//删除第一行中再次运行的时候需要删除的ProcessSum, Null等列
async function deleteProcessSum() {
  await Excel.run(async (context) => {
    console.log("Enter deleteProcessSum");
    const sheet = context.workbook.worksheets.getItem("Bridge Data");

    // 获取工作表的 usedRange，并加载其第一行的值
    const usedRange = sheet.getUsedRange();
    usedRange.load("rowCount, columnCount"); // 加载范围信息
    await context.sync();
    console.log("DeleteProcessSum 2");

    // 获取 usedRange 的第一行
    const firstRow = usedRange.getRow(0);
    firstRow.load("values"); // 加载第一行的值
    await context.sync();

    // 获取第一行的值
    const values = firstRow.values[0];
    console.log("First row values:", values);

    // 找到值为 "ProcessSum" 或 "Null" 的列索引
    const columnsToDelete = [];
    values.forEach((value, index) => {
      if (value === "ProcessSum" || value === "Null") {
        columnsToDelete.push(index + 1); // Excel 列索引从 1 开始
      }
    });

    console.log("Columns to delete:", columnsToDelete);

    // 按列索引删除列，从最后一列开始删除以避免索引错位
    columnsToDelete.reverse().forEach((colIndex) => {
      const columnRange = sheet.getRangeByIndexes(0, colIndex - 1, usedRange.rowCount, 1);
      columnRange.delete(Excel.DeleteShiftDirection.left);
    });

    await context.sync();
    console.log("Selected columns deleted successfully.");
  }).catch((error) => {
    console.error(error);
  });
};


//检测是否存在某个工作表，返回布尔值
// 使用示例
// (async () => {
//   const exists = await doesSheetExist("Bridge Data");
//   console.log("工作表 'Bridge Data' 是否存在: " + exists);
// })();
async function doesSheetExist(sheetName) {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;

      // 获取所有工作表
      const sheets = workbook.worksheets;
      sheets.load("items/name");

      await context.sync(); // 同步数据

      // 检查工作表是否存在
      const sheetExists = sheets.items.some(sheet => sheet.name === sheetName);

      console.log(sheetExists); // 输出结果
      return sheetExists; // 返回布尔值
    });
  } catch (error) {
    console.error("检测工作表时出错: ", error);
    return false; // 如果发生错误，返回 false
  }
}

async function TaskPaneStart(SheetName) {
  try {
    return await Excel.run(async (context) => {
      console.log("TaskPaneStart 开始")
      //判断是否存在"Bridge Data"工作表
      // let BridgeCheck = await doesSheetExist("Bridge Data");
      let BridgeCheck = await doesSheetExist(`${SheetName}`); //从Bridge Data 修改成Data
      console.log(`Check ${SheetName} is `);
      console.log(BridgeCheck);                                                 
      //若不存在"Bridge Data"工作表

      if(!BridgeCheck) {
        //生成Bridge Data 工作表空表
        await createSourceData(SheetName);
        return 0
      }
      console.log("TaskPaneStart 完成")
      return 1; // 返回布尔值
    });
  } catch (error) {
    console.error("TaskPaneStartError: ", error);
    return 2; // 如果发生错误，返回 false
  }
}

// //////------------检查是否在第一行里有Key----------------------
// //////------------检查 Key 的函数-----------------------------
// async function CheckKey() {
//   try {
//       return await Excel.run(async (context) => {
//           const sheet = context.workbook.worksheets.getItem("Bridge Data");
//           const usedRange = sheet.getUsedRange();
//           usedRange.load("values");

//           await context.sync();

//           // 检查第一行是否包含 "Key"
//           const firstRow = usedRange.values[0];
//           const hasKey = firstRow.includes("Key");

//           if (!hasKey) {
//               // 显示警告并禁用其他容器
//               showKeyWarning();
//               return false;
//           } else {
//               // 隐藏警告并恢复界面
//               hideKeyWarning();
//               return true;
//           }
//       });
//   } catch (error) {
//       console.error("Error checking Key:", error);
//   }
// }

// // 显示 Key 警告
// function showKeyWarning() {
//   document.querySelector("#keyWarningContainer").style.display = "flex";
//   document.querySelector("#modalOverlay").style.display = "block";
//   document.querySelector(".container").classList.add("disabled");
// }

// // 隐藏 Key 警告
// function hideKeyWarning() {
//   document.querySelector("#keyWarningContainer").style.display = "none";
//   document.querySelector("#modalOverlay").style.display = "none";
//   document.querySelector(".container").classList.remove("disabled");
// }

// //////------------检查是否在第一行里有Key----------------------


////-----------------保存Data中的字段和类型到TempVar中-----
async function createFieldTypeMapping() {
  await Excel.run(async (context) => {
    const workbook = context.workbook;

    // 获取 Data 工作表的 usedRange
    const bridgeSheet = workbook.worksheets.getItem("Data");
    const usedRange = bridgeSheet.getUsedRange();
    usedRange.load("values"); // 加载所有单元格的值

    await context.sync();

    // 从第二列开始的第一行和第二行获取值
    const values = usedRange.values;
    const headers = values[1].slice(1); // 第一行从第二列开始的值
    const types = values[0].slice(1); // 第二行从第二列开始的值

    // 构建 FieldType 对象
    const FieldType = {};
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && types[i] && types[i] != "ProcessSum" && types[i] != "Null") {
        FieldType[headers[i]] = types[i];
      }
    }

    console.log("FieldType object:", FieldType);

    const sheet = workbook.worksheets.getItem("TempVar");
    sheet.getRange("D1").values = [["Field"]];
    sheet.getRange("E1").values = [["Type"]];

    const fields = Object.keys(FieldType);
    const typesValues = Object.values(FieldType);

    const fieldsRange = sheet.getRange(`D2:D${fields.length + 1}`);
    const typesRange = sheet.getRange(`E2:E${typesValues.length + 1}`);

    fieldsRange.values = fields.map((field) => [field]);
    typesRange.values = typesValues.map((type) => [type]);

    await context.sync();

    console.log("FieldType mapping written to TempVar worksheet.");
  });
}


////-----------------对比Data中的字段和类型 和 TempVar中的已有数据-----
async function compareFieldType() {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;

    // 获取 Data 工作表的 usedRange
    const bridgeSheet = workbook.worksheets.getItem("Data");
    const bridgeRange = bridgeSheet.getUsedRange();
    bridgeRange.load("values"); // 加载所有单元格的值

    // 获取 TempVar 工作表的 usedRange
    const tempVarSheet = workbook.worksheets.getItem("TempVar");
    const tempVarRange = tempVarSheet.getUsedRange();
    tempVarRange.load("values");

    await context.sync();

    // 从 Data 中构建新的 FieldType 对象
    const bridgeValues = bridgeRange.values;
    const headers = bridgeValues[1].slice(1); // 第二行从第二列开始的值作为 headers
    const types = bridgeValues[0].slice(1); // 第一行从第二列开始的值作为 types

    const newFieldType = {};
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && types[i] && types[i] != "ProcessSum" && types[i] != "Null") {
        newFieldType[headers[i]] = types[i];
      }
    }

    console.log("New FieldType:", newFieldType);

    // 从 TempVar 中提取旧的 FieldType 数据
    const tempVarValues = tempVarRange.values;
    const oldFieldType = {};
    for (let i = 1; i < tempVarValues.length; i++) {
      const field = tempVarValues[i][3]; // 第 D 列（索引 3）
      const type = tempVarValues[i][4]; // 第 E 列（索引 4）
      if (field && type) {
        oldFieldType[field] = type;
      }
    }

    console.log("Old FieldType:", oldFieldType);

    // 比较新旧 FieldType 对象
    const newHeaders = [];
    const changedHeaders = [];
    const removedHeaders = [];

    // 检查新的 headers 和 types
    for (const header of Object.keys(newFieldType)) {
      if (!oldFieldType.hasOwnProperty(header)) {
        newHeaders.push(header); // 新的 header
      } else if (oldFieldType[header] !== newFieldType[header]) {
        changedHeaders.push(header); // header 的 type 发生变化
      }
    }

    // 检查被移除的 headers
    for (const header of Object.keys(oldFieldType)) {
      if (!newFieldType.hasOwnProperty(header)) {
        removedHeaders.push(header); // 被移除的 header
      }
    }
    
    console.log("compareField Here");
    // 返回结果
    if (newHeaders.length === 0 && changedHeaders.length === 0 && removedHeaders.length === 0) {
      return 0; // 无变化
    } else if (newHeaders.length > 0) {
      return { result: 1, newHeaders }; // 有新的 headers
    } else if (changedHeaders.length > 0) {
      return { result: 2, changedHeaders }; // headers 的 types 发生变化
    } else if (removedHeaders.length > 0) {
      return { result: 3, removedHeaders }; // 有被移除的 headers
    }
  });
}

//监控判断数据的维度类型和维度有没有变化
async function handleCompareFieldType() {
  return await Excel.run(async (context) => {
      const workbook = context.workbook;

      // 检查是否存在 TempVar 工作表
      const sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetNames = sheets.items.map(sheet => sheet.name);
      if (!sheetNames.includes("TempVar")) {
          console.log("TempVar 工作表不存在。");
          return false; // 如果 TempVar 不存在，直接返回
      }

      // 调用 compareFieldType 函数
      const result = await compareFieldType();

      if (result === 0) {
          console.log("No changes detected.");
          await CreateDropList(); //没有变化则直接生成下拉菜单
          await updateDropdownsFromSelectedValues(); //生成下拉根据临时保存变量选中之前已经选中的选项
          console.log("根据TempVar生成下拉菜单")
          return false;
      }

      // 准备提示内容
      // let message = "";
      // if (result.result === 1) {
      //     message = `有新的 headers: ${result.newHeaders.join(", ")}，是否要重新生成Waterfall?`;
      // } else if (result.result === 2) {
      //     message = `headers 的类型发生变化: ${result.changedHeaders.join(", ")}，是否要重新生成Waterfall?`;
      // } else if (result.result === 3) {
      //     message = `有被移除的 headers: ${result.removedHeaders.join(", ")}，是否要重新生成Waterfall?`;
      // }

      let message = "";
      if (result.result >0 ) {
          message = `数据源有变化，是否要重新生成Waterfall? <a href="#" id="detailLink">Detail</a>`;
      }

      // 更新提示框内容
      const promptElement = document.getElementById("dynamicWaterfallPrompt");
      // promptElement.querySelector(".waterfall-message").textContent = message;
      promptElement.querySelector(".waterfall-message").innerHTML = message;

      // 显示提示框
      const container = document.querySelector(".container");

      modalOverlay.style.display = "block";
      promptElement.style.display = "flex";
      container.classList.add("disabled");

      // 绑定 Detail 超链接点击事件
      const detailLink = document.getElementById("detailLink");
      if (detailLink) {
          detailLink.addEventListener("click", async (e) => {
              e.preventDefault();
              await handleDetail();
          });
      }

      // 处理用户确认或取消操作
      await new Promise((resolve) => {
          const confirmButton = document.getElementById("confirmDynamicWaterfall");
          const cancelButton = document.getElementById("cancelDynamicWaterfall");

          const handleConfirm = () => {
              GblComparison = true; // 检测是否被对比过表头，避免循环调用
              
              console.log("Confirmed. Proceeding to regenerate Waterfall...");
              hidePrompt();
              //runProgramHandler();
              // 这里不再直接调用 runProgramHandler，而是返回状态
              resolve(true);
          };

          const handleCancel = () => {
              console.log("Canceled. No action taken.");
              GblComparison = false; // 检测是否被对比过表头，避免循环调用
              hidePrompt();
              resolve(false);
          };

          confirmButton.addEventListener("click", handleConfirm, { once: true });
          cancelButton.addEventListener("click", handleCancel, { once: true });
      });

      // 隐藏提示框函数
      function hidePrompt() {
          modalOverlay.style.display = "none";
          promptElement.style.display = "none";
          container.classList.remove("disabled");
      }
      
  });
}

//------------删除特定的工作表-------------
async function deleteSheetsIfExist(sheetNames) {
  await Excel.run(async (context) => {
      console.log("Enter deleteSheets");
      const workbook = context.workbook;
      const sheets = workbook.worksheets;
      sheets.load("items/name"); // 加载所有工作表的名称

      await context.sync(); // 同步以确保工作表信息加载完成

      const existingSheetNames = sheets.items.map(sheet => sheet.name);

      for (const sheetName of sheetNames) {
          if (existingSheetNames.includes(sheetName)) {
              console.log(`Deleting sheet: ${sheetName}`);
              const sheet = sheets.getItem(sheetName);
              sheet.delete(); // 删除工作表
          } else {
              console.log(`Sheet not found: ${sheetName}`);
          }
      }

      await context.sync(); // 确保删除操作同步到 Excel
      console.log("Specified sheets checked and deleted if found.");
  }).catch((error) => {
      console.error("Error deleting sheets:", error);
  });
}


// 处理 Detail 功能
async function handleDetail() {
  await Excel.run(async (context) => {
      const workbook = context.workbook;

      // 检查是否存在 "Data Change" 工作表
      const sheets = workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetNames = sheets.items.map(sheet => sheet.name);

      // 如果存在 "Data Change" 工作表，则删除
      if (sheetNames.includes("Data Change")) {
          console.log("Deleting existing 'Data Change' sheet.");
          sheets.getItem("Data Change").delete();
          await context.sync();
      }

      // 创建新的 "Data Change" 工作表
      console.log("Creating new 'Data Change' sheet.");
      const newSheet = sheets.add("Data Change");

      // 跳转到 "Data Change" 的 B3 单元格
      const targetCell = newSheet.getRange("B3");
      targetCell.select();

      await context.sync();
  }).catch((error) => {
      console.error("Error handling detail link:", error);
  });
}


//设置跳转到Contribution的链接
async function insertHyperlink(hyperlinkText, worksheetName, linkCell) {
  try {
    await Excel.run(async (context) => {

      let VarTempSheet = context.workbook.worksheets.getItem("TempVar");
      let ContributionStart = VarTempSheet.getRange("B10");
      ContributionStart.load("address,values");
      await context.sync();

      let targetCellAddress = ContributionStart.values[0][0];
      // 获取工作表
      const sheet = context.workbook.worksheets.getItem(worksheetName);

      // 设置超链接的目标单元格
      const targetCellFullAddress = `#${targetCellAddress}`;;

      // 获取目标单元格范围
      const cell = sheet.getRange(linkCell);

      // 设置新的超链接
      cell.values = [[hyperlinkText]]; // 设置显示名称
      cell.hyperlink = { textToDisplay: hyperlinkText, address: targetCellFullAddress }; // 设置跳转地址

      // 加载更改并同步
      await context.sync();

      console.log("Hyperlink inserted successfully.");
    });
  } catch (error) {
    console.error("Error inserting hyperlink:", error);
  }
}

async function GetBaseLabel() {
  await Excel.run(async (context) => {

    let VarTempSheet = context.workbook.worksheets.getItem("TempVar");
    let BaseLabel = VarTempSheet.getRange("B13");
    BaseLabel.load("values");
    await context.sync();
    console.log("BaseLable is " + BaseLabel.values[0][0]); //从TempVar工作表中获取地址

    // Get the range by address and ensure it has one row
    let worksheet = context.workbook.worksheets.getItem("Process");
    let range = worksheet.getRange(BaseLabel.values[0][0]);
    range.load("address, values");
    await context.sync();

    if (range.values.length !== 1) {
      console.error("The range must contain only one row.");
      return;
    }

    // Move the entire range up by two rows
    const rangeAbove = range.getOffsetRange(-2, 0);
    rangeAbove.load("values");
    await context.sync();

    const filteredAddresses = [];
    let startAddress = null;
    let endAddress = null;
    
    // Loop through the cells in the range
    for (let colIndex = 0; colIndex < range.values[0].length; colIndex++) {
      let valueAbove = rangeAbove.values[0][colIndex];
      console.log("valueAbove is " + valueAbove);

            // Set start and end addresses based on condition
      if (!(valueAbove === "ProcessSum" || valueAbove === "NULL")) {
        let cell = range.getCell(0, colIndex);
        cell.load("address");
        await context.sync();

        if (!startAddress) {
          startAddress = cell.address;
        }
        endAddress = cell.address;
      }
      console.log("startAddress is " + startAddress);
      console.log("endAddress is " + endAddress);
    }

    // Create a continuous range from start to end
    if (startAddress && endAddress) {
      let BaseLabelRange = worksheet.getRange(`${startAddress.split("!")[0]}!${startAddress.split("!")[1].split(":")[0]}:${endAddress.split("!")[1].split(":")[0]}`);
      BaseLabelRange.load("address,values");
      await context.sync();

      console.log("BaseLabelRange address is " + BaseLabelRange.address);
      
    } else {
      console.log("No cells meet the criteria.");
    }
  });
}


async function GetVarianceRange() {
  await Excel.run(async (context) => {
    console.log("Enter GetVariance");
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    // let TempBaseRange = TempVarSheet.getRange("B2");
    // console.log("Enter GetVariance 2");
    // TempBaseRange.load("values"); // 获取临时变量工作表中的BaseRange的变量
    // await context.sync();

    // console.log("BaseRange.address is " + TempBaseRange.values[0][0]);

    let ProcessSheet = context.workbook.worksheets.getItem("Process");
    let ProcessUsedRange = ProcessSheet.getUsedRange();
    // let ProcessFirstRow = ProcessUsedRange.getRow(0);
    // let ProcessSecondRow = ProcessUsedRange.getRow(1);
    // let ProcessThirdRow = ProcessUsedRange.getRow(2);
    ProcessUsedRange.load("address,values,rowCount,columnCount");
    // ProcessFirstRow.load("address,values,rowCount,columnCount");
    // ProcessSecondRow.load("address,values,rowCount,columnCount");
    // ProcessThirdRow.load("address,values,rowCount,columnCount");
    // let BaseRange = ProcessSheet.getRange(TempBaseRange.values[0][0]);
    // BaseRange.load("address,values,rowCount,columnCount"); //获取BaseRange在Process中的Range
    await context.sync();

    // console.log("BaseRange is " + BaseRange.address);

    // let BaseRangeStart = BaseRange.getCell(0,0);
    // let BaseKey = BaseRange.getColumn(0);
    // BaseKey.load("address, values, rowCount, columnCount"); //获取BaseKey的RangeRange
    // let BaseRangeTitle = BaseRangeStart.getOffsetRange(0,1).getAbsoluteResizedRange(1,BaseRange.columnCount-1);
    // BaseRangeTitle.load("address,values,rowCount,columnCount"); //获取BaseTitle的Range
    // await context.sync();

    // console.log("BaseRangeTitle is " + BaseRangeTitle.address);
    // console.log("BaseKey is " + BaseKey.address);

    //----下面开始去除BaseRange中ProcessSum 和 NULL的对应变量
    // let BaseRangeTitleStart = BaseRangeTitle.getCell(0,0); 
    // let BaseTitleTypeStart = BaseRangeTitleStart.getOffsetRange(-2,0); 
    // let BaseRangeTitleCell = null;
    // let BaseTitleTypeCell = null;
    // let PreviousTypeCell = null;
    // for (i = 0; i < BaseRangeTitle.columnCount-1;i++){
    //   BaseRangeTitleCell = BaseRangeTitleStart.getCell(0, i); //BaseRangeTitle单元格循环遍历
    //   BaseTitleTypeCell = BaseTitleTypeStart.getOffsetRange(0, i); //每个变量在第一行对应的数据类型
    //   BaseRangeTitleCell.load("address, values");
    //   BaseTitleTypeCell.load("address, values");
    //   await context.sync();

    //   console.log("BaseRangeTitleCell is " + BaseRangeTitleCell.values[0][0]);
    //   console.log("BaseTitleTypeCell is " + BaseTitleTypeCell.values[0][0]);

    //   if ((BaseTitleTypeCell.values[0][0] == "ProcessSum" || BaseTitleTypeCell.values[0][0] == "NULL")){
    //     break;
    //   }
    //   PreviousTypeCell = ProcessSheet.getRange(BaseTitleTypeCell.address);
    // }
    
    // PreviousTypeCell.load("address");
    // await context.sync();

    //----获取Varaicne需要的在Process的Range-----
    // let VarianceRight = getRangeDetails(PreviousTypeCell.address).rightColumn;
    // let Varianceleft  = getRangeDetails(BaseRange.address).leftColumn;
    // let VarianceTop = getRangeDetails(BaseRange.address).topRow;
    // let VarianceBottom = getRangeDetails(BaseRange.address).bottomRow;
    // let VarianceRange = ProcessSheet.getRange(`${Varianceleft}${VarianceTop}:${VarianceRight}${VarianceBottom}`);
    // VarianceRange.load("address,values,rowCount,columnCount");
    // await context.sync();
    
    // console.log("VarianceRange is " + VarianceRange.address);

    //---获取Watarfall工作表中ContributionTable的地址 ---
    let ContributionTableAddress = TempVarSheet.getRange("B16");
    ContributionTableAddress.load("values");
    await context.sync();

    console.log("ContributionTableKey is " + ContributionTableAddress.values[0][0]);

    let WaterfallSheet = context.workbook.worksheets.getItem("Waterfall");
    let ContributionTable = WaterfallSheet.getRange(ContributionTableAddress.values[0][0]);
    ContributionTable.load("rowCount,columnCount");
    let ContributionTableAddressDetail = getRangeDetails(ContributionTableAddress.values[0][0]);
    let ContributionLeft = ContributionTableAddressDetail.leftColumn;
    let ContributionBottom = ContributionTableAddressDetail.bottomRow;
     //---Variance在Waterfall中的起点---
    let VarianceTableName = WaterfallSheet.getRange(`${ContributionLeft}${ContributionBottom}`).getOffsetRange(3,0);
    VarianceTableName.values = [["Variance"]];
    let VarianceTableStart = VarianceTableName.getOffsetRange(1,0);
    VarianceTableName.load("address");
    await context.sync();
    
    console.log("VarianceTableStart is " + VarianceTableName.address);

    //---将Waterfall 中的 ContributionTable拷贝到 下方中 ---
    VarianceTableStart.copyFrom(ContributionTable,Excel.RangeCopyType.formats);
    VarianceTableStart.copyFrom(ContributionTable,Excel.RangeCopyType.values);
    await context.sync();

    //获取VarianceTable的Title和Key
    let VarianceTable = VarianceTableStart.getAbsoluteResizedRange(ContributionTable.rowCount,ContributionTable.columnCount);
    let VarianceTitle = VarianceTableStart.getOffsetRange(0,1).getAbsoluteResizedRange(1,ContributionTable.columnCount -1);
    let VarianceKey = VarianceTableStart.getOffsetRange(1,0).getAbsoluteResizedRange(ContributionTable.rowCount -1,1)
    VarianceTable.load("address,values,rowCount,columnCount");
    VarianceTitle.load("address,values,rowCount,columnCount");
    VarianceKey.load("address,values,rowCount,columnCount");
    await context.sync();

    console.log("VarianceTable is " + VarianceTable.address);
    console.log("VarianceTitle is " + VarianceTitle.address);
    console.log("VarianceKey is " + VarianceKey.address);
    
    
    for(let TitleIndex = 0; TitleIndex < VarianceTitle.values[0].length; TitleIndex++){     //Variance的变量表头
      console.log("TitleIndex is " + TitleIndex);
      KeyLoop: 
      for(let KeyIndex = 0; KeyIndex < VarianceKey.values.length; KeyIndex++){           //Variance的Key部分循环

        console.log("KeyIndex is " + KeyIndex);
        for(let ColumnIndex = 0; ColumnIndex < ProcessUsedRange.values[1].length; ColumnIndex++){    //Process的第二行
          
          if(ProcessUsedRange.values[1][ColumnIndex] === "TargetPT" && ProcessUsedRange.values[2][ColumnIndex] === VarianceTitle.values[0][TitleIndex]){
            console.log("ColumnIndex is " + ColumnIndex);
            

            for(let RowIndex = 0; RowIndex < ProcessUsedRange.rowCount; RowIndex++){ //Process工作表的第一列
              console.log("RowIndex 0 is " + RowIndex);
              if(ProcessUsedRange.values[RowIndex][0] === VarianceKey.values[KeyIndex][0]){
                console.log("RowIndex 1 is " + RowIndex);
                console.log("ProcessUsedRange.values[RowIndex][0] is " + ProcessUsedRange.values[RowIndex][0]);
                console.log("VarianceKey.values[KeyIndex][0] is " + VarianceKey.values[KeyIndex][0]);
                let TargetVariable = ProcessUsedRange.values[RowIndex][ColumnIndex];    //获取Process中的Target中对应的变量
                console.log("TargetVariable is " + TargetVariable);
                //----------寻找Process中Base对应的变量----------------
                  for(let BaseColumnIndex = 0; BaseColumnIndex < ProcessUsedRange.values[1].length; BaseColumnIndex++){    //Process的第二行
                    console.log("BaseColumnIndex 0 is " + BaseColumnIndex);
                    if(ProcessUsedRange.values[1][BaseColumnIndex] === "BasePT" && ProcessUsedRange.values[2][BaseColumnIndex] === VarianceTitle.values[0][TitleIndex]){
                        console.log("BaseColumnIndex 1 is " + BaseColumnIndex);
                        for(let BaseRowIndex = 0; BaseRowIndex < ProcessUsedRange.rowCount; BaseRowIndex++){ //Process工作表的第一列
                          console.log("BaseRowIndex 0 is " + BaseRowIndex);
                          if(ProcessUsedRange.values[BaseRowIndex][0] === VarianceKey.values[KeyIndex][0]){
                            console.log("BaseRowIndex 1 is " + BaseRowIndex);
                            let BaseVariable = ProcessUsedRange.values[BaseRowIndex][BaseColumnIndex];    //获取Process中的Base中对应的变量
                            console.log("BaseVariable is " + BaseVariable);
                            //  let Variance = Number(TargetVariable) - Number(BaseVariable); //求得差异
                            let Variance = TargetVariable - BaseVariable; //求得差异
                            console.log("Variance is " + Variance);
                            let CurrentVarianceCell = VarianceTable.getCell(KeyIndex + 1,TitleIndex + 1);           
                            CurrentVarianceCell.values = [[Variance]]; //将差异放到VarianceTable 对应的单元格
                            CurrentVarianceCell.copyFrom(ProcessUsedRange.getCell(BaseRowIndex,BaseColumnIndex),Excel.RangeCopyType.formats)
                              await context.sync();
                              continue KeyLoop; // 跳到最外层循环的下一次迭代
                            }
          
                        }            
          
                    }
          
                  }
          
              }

            }            

          }

        }

      }

    }
    VarianceTable.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.continuous;
    VarianceTable.format.borders.getItem('EdgeTop').weight = Excel.BorderWeight.thin;
    VarianceTable.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.continuous;
    VarianceTable.format.borders.getItem('EdgeBottom').weight = Excel.BorderWeight.thin;
    VarianceTable.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.continuous;
    VarianceTable.format.borders.getItem('EdgeLeft').weight = Excel.BorderWeight.thin;
    VarianceTable.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.continuous;
    VarianceTable.format.borders.getItem('EdgeRight').weight = Excel.BorderWeight.thin;

  });
}

// -----直接使用Process工作表生成Variance Table-----
async function CreateVarianceTable() {
  await Excel.run(async (context) => {
    console.log("Enter GetVariance");
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    let TempBaseRange = TempVarSheet.getRange("B2");
    console.log("Enter GetVariance 2");
    TempBaseRange.load("values"); // 获取临时变量工作表中的BaseRange的变量

    await context.sync();

    // console.log("BaseRange.address is " + TempBaseRange.values[0][0]);
    //获取Process表格中的数据
    let ProcessSheet = context.workbook.worksheets.getItem("Process");
    let ProcessUsedRange = ProcessSheet.getUsedRange();
    let ProcessFirstRow = ProcessUsedRange.getRow(0);
    let ProcessSecondRow = ProcessUsedRange.getRow(1);
    let ProcessThirdRow = ProcessUsedRange.getRow(2);
    ProcessUsedRange.load("address,values,rowCount,columnCount");
    ProcessFirstRow.load("address,values,rowCount,columnCount");
    ProcessSecondRow.load("address,values,rowCount,columnCount");
    ProcessThirdRow.load("address,values,rowCount,columnCount");
    await context.sync();
    console.log("Enter GetVariance 3");
    let ProcessUsedRangeAddress = await GetRangeAddress("Process", ProcessUsedRange.address); //获得每个单元格的地址
    // console.log("BaseRange is " + BaseRange.address);

    // 用于存储BaseRange中不是ProcessSum和NULL的第一个列字符
    let BaseRightColumns = null;
    let ResultVar = null; // 数据类型是Result的，需要删除
    //除去掉ProcessSum和NULL的数据类型
    for(let ColumnIndex = 0; ColumnIndex < ProcessUsedRange.columnCount; ColumnIndex++){
      let secondRowValue = ProcessSecondRow.values[0][ColumnIndex]; // 获取第二行的值
      let firstRowValue = ProcessFirstRow.values[0][ColumnIndex]; // 获取第一行的值
      let thirdRowvalue = ProcessThirdRow.values[0][ColumnIndex]; //获取第三行的值

      if (
        secondRowValue === "BasePT" &&
        (firstRowValue === "ProcessSum" || firstRowValue === "NULL")
      ) {
        // 如果符合条件，返回当前单元格的前一列的列字符
        if (ColumnIndex > 0) { // 确保有前一列
          BaseRightColumns = getRangeDetails(ProcessUsedRangeAddress[1][ColumnIndex - 1]).rightColumn
          break;
        }
      }else if(firstRowValue === "Result" ){  //找到Result类型的变量，最后删除不出现在VarianceTable中******这里只能有一个Result
          ResultVar = thirdRowvalue;
          console.log("ResultVar is " + ResultVar);
      }

    }

    // let OldBaseRange = ProcessSheet.getRange(TempBaseRange.values[0][0]);
    //获取BaseRange在Process中的Range
    let TempBaseRangeAddress = getRangeDetails(TempBaseRange.values[0][0]);
    let LeftColumn = TempBaseRangeAddress.leftColumn;
    let TopRow = TempBaseRangeAddress.topRow;
    let BottomRow = TempBaseRangeAddress.bottomRow;
    let BaseRange = ProcessSheet.getRange(`${LeftColumn}${TopRow}:${BaseRightColumns}${BottomRow}`);
    BaseRange.load("address,values,rowCount,columnCount"); 
    await context.sync();
    console.log("BaseRange is " + BaseRange.address);

    let BaseRangeStart = BaseRange.getCell(0,0);
    let BaseKey = BaseRange.getColumn(0);
    BaseKey.load("address, values, rowCount, columnCount"); //获取BaseKey的Range
    let BaseRangeTitle = BaseRangeStart.getOffsetRange(0,1).getAbsoluteResizedRange(1,BaseRange.columnCount-1);
    BaseRangeTitle.load("address,values,rowCount,columnCount"); //获取BaseTitle的Range
    let BaseTitleData = BaseRangeStart.getOffsetRange(0,1).getAbsoluteResizedRange(BaseRange.rowCount,BaseRange.columnCount-1);
    BaseTitleData.load("address,values,rowCount,columnCount"); //获取BaseRange中除去Key以外的单元格
    let BaseData = BaseRangeStart.getOffsetRange(1,1).getAbsoluteResizedRange(BaseRange.rowCount-1,BaseRange.columnCount-1);//获取BaseRange中除去Key和Title以外的数据Range
    BaseData.load("address,values,rowCount,columnCount");
    await context.sync();

    console.log("BaseRangeTitle is " + BaseRangeTitle.address);
    console.log("BaseKey is " + BaseKey.address);

    let ContributionEnd = TempVarSheet.getRange("B19");
    ContributionEnd.load("values");
    await context.sync();

    let VarianceStart = ProcessSheet.getRange(ContributionEnd.values[0][0]).getOffsetRange(0,1); // 往右移动一格，作为Variance的起始地址
    VarianceStart.load("address");
    VarianceStart.copyFrom(BaseKey,Excel.RangeCopyType.formats);//拷贝BaseKey
    VarianceStart.copyFrom(BaseKey,Excel.RangeCopyType.values);
    let VarianceKey = VarianceStart.getOffsetRange(1,0).getAbsoluteResizedRange(BaseRange.rowCount -1,1);
    VarianceKey.load("address,values,rowCount,columnCount");
    
    let VarianceTitleData = VarianceStart.getOffsetRange(0,1).getAbsoluteResizedRange(BaseTitleData.rowCount,BaseTitleData.columnCount); //获得Title部分

    VarianceTitleData.copyFrom(BaseTitleData,Excel.RangeCopyType.formats);
    VarianceTitleData.copyFrom(BaseTitleData,Excel.RangeCopyType.values);
    let VarianceData = VarianceStart.getOffsetRange(1,1).getAbsoluteResizedRange(BaseData.rowCount,BaseData.columnCount); //获得数据部分Range
    VarianceTitleData.load("address,values,rowCount,columnCount"); //需要放在Copy 后面才有数值
    VarianceData.load("address,values,rowCount,columnCount");
    VarianceData.clear(Excel.ClearApplyTo.contents); // 只清除数据，保留格式

    await context.sync();
    console.log("VarianceKey is " + VarianceKey.address);
    console.log("VarianceTitleData is " + VarianceTitleData.address);
    console.log("VarianceData is " + VarianceData.address);

    // 准备一个空的二维数组
    const formulaArray = Array.from({ length: VarianceData.rowCount }, () => new Array(VarianceData.columnCount));
    //整体把所有的公式写到数组里，一次性赋值
    for(let TitleIndex = 0;TitleIndex < VarianceTitleData.columnCount; TitleIndex++){
      KeyLoop:
      for(let KeyIndex = 0;KeyIndex < VarianceKey.rowCount;KeyIndex++){
      for(let ProcessColumnIndex = 0;ProcessColumnIndex < ProcessUsedRange.columnCount;ProcessColumnIndex++){

        if(ProcessUsedRange.values[2][ProcessColumnIndex] === VarianceTitleData.values[0][TitleIndex] && ProcessUsedRange.values[1][ProcessColumnIndex] === "TargetPT" ){
          console.log("ProcessUsedRange.values 2 is " + ProcessUsedRange.values[2][ProcessColumnIndex]);


            for(let ProcessRowIndex = 0;ProcessRowIndex < ProcessUsedRange.rowCount; ProcessRowIndex++){
              if(ProcessUsedRange.values[ProcessRowIndex][0] === VarianceKey.values[KeyIndex][0]){
                let TargetAddress = ProcessUsedRangeAddress[ProcessRowIndex][ProcessColumnIndex];
                console.log("TargetAddress is " + TargetAddress); 
                //查找Base对应的单元格

                  for(let BaseProcessColumnIndex = 0; BaseProcessColumnIndex < ProcessUsedRange.columnCount; BaseProcessColumnIndex++){
                    if(ProcessUsedRange.values[2][BaseProcessColumnIndex] === VarianceTitleData.values[0][TitleIndex] && ProcessUsedRange.values[1][BaseProcessColumnIndex] === "BasePT" ){
                      //因为是和Target的变量再同一行，不需要比较RowIndex
                      let BaseAddress = ProcessUsedRangeAddress[ProcessRowIndex][BaseProcessColumnIndex];
                      console.log("BaseAddress is " + BaseAddress);
                      formulaArray[KeyIndex][TitleIndex] = `=${TargetAddress}-${BaseAddress}`;
                      continue KeyLoop;
                    } 
            
                  }
                  
              }


            }

          }

        } 

      }
      
    }

    VarianceData.formulas = formulaArray;
    await context.sync();

    let ResultCol = null;
    //删除掉数据类型是Result的列，不显示在Variance中
    let VarianceTitleDataAddress = await GetRangeAddress("Process", VarianceTitleData.address); //获得每个单元格的地址
    for(let col = 0; col <VarianceTitleData.columnCount;col++){
      if(VarianceTitleData.values[0][col] === ResultVar){
          ResultCol = getRangeDetails(VarianceTitleDataAddress[0][col]).leftColumn
          let ResultColRange = ProcessSheet.getRange(`${ResultCol}:${ResultCol}`);
          // 删除列，右侧列会向左移动
          ResultColRange.delete(Excel.DeleteShiftDirection.left);
          await context.sync();
          console.log(`Column ${ResultCol} deleted successfully.`);
      }

    }
    console.log("ResultCol is " + ResultCol);
    console.log("VarianceTitleData.address is " + VarianceTitleData.address);
    //删除Result列后新的地址：
    let NewVarianceTitleDataAddress = getShiftedRangeAfterRemoving(VarianceTitleData.address,ResultCol);
    console.log("NewVarianceTitleDataAddress is " + NewVarianceTitleDataAddress);
    //获取Process中VarianceTableRange
    let VarianceStartAddress = getRangeDetails(VarianceStart.address);
    let VarianceLeftColumn = VarianceStartAddress.leftColumn;
    let VarianceTopRow = VarianceStartAddress.topRow;
    let NewVarianceTitleDataAddressDetail = getRangeDetails(NewVarianceTitleDataAddress);
    let VarianceRightColumn = NewVarianceTitleDataAddressDetail.rightColumn;
    let VarianceBottomRow = NewVarianceTitleDataAddressDetail.bottomRow;
    let VarianceRange = ProcessSheet.getRange(`${VarianceLeftColumn}${VarianceTopRow}:${VarianceRightColumn}${VarianceBottomRow}`);
    VarianceRange.load("address");
    await context.sync();

    let TempVarianceRangeName = TempVarSheet.getRange("B21");
    TempVarianceRangeName.values = [["VarianceTable"]];
    let TempVarianceRange = TempVarSheet.getRange("B22");
    TempVarianceRange.values = [[VarianceRange.address]]; //将Process中的VarianceTableRange保存在TempVar工作表中
    await context.sync();

  });
}


//将在Process生成Variance的Table贴入到Waterfall工作表中
async function WaterfallVarianceTable() {
  await Excel.run(async (context) => {
    console.log("enter WaterfallVariance");
    let ProcessSheet = context.workbook.worksheets.getItem("Process");
    let WaterfallSheet = context.workbook.worksheets.getItem("Waterfall");
    let TempVarSheet = context.workbook.worksheets.getItem("TempVar");
    let VarianceTableVar = TempVarSheet.getRange("B22");
    VarianceTableVar.load("values");
    let ContributionVar = TempVarSheet.getRange("B16");
    ContributionVar.load("values");
    await context.sync();

    let VarianceTable = ProcessSheet.getRange(VarianceTableVar.values[0][0]); //在Process中的VarianceTable
    VarianceTable.load("rowCount,columnCount");
    let ContributionVarAddress = getRangeDetails(ContributionVar.values[0][0]);
    let ContributionTableLeft = ContributionVarAddress.leftColumn;
    let ContributionTableBottom = ContributionVarAddress.bottomRow;
    
    let WaterfallVarianceName = WaterfallSheet.getRange(`${ContributionTableLeft}${ContributionTableBottom}`).getOffsetRange(2,0); //往下移动两格
    WaterfallVarianceName.values = [["Variance"]];
    let WaterfallVarianceStart = WaterfallVarianceName.getOffsetRange(1,0);
    WaterfallVarianceStart.copyFrom(VarianceTable,Excel.RangeCopyType.formats);
    WaterfallVarianceStart.copyFrom(VarianceTable,Excel.RangeCopyType.values);
    await context.sync();

    let WaterfallVarianceTable = WaterfallVarianceStart.getAbsoluteResizedRange(VarianceTable.rowCount,VarianceTable.columnCount);
    WaterfallVarianceTable.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.continuous;
    WaterfallVarianceTable.format.borders.getItem('EdgeTop').weight = Excel.BorderWeight.thin;
    WaterfallVarianceTable.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.continuous;
    WaterfallVarianceTable.format.borders.getItem('EdgeBottom').weight = Excel.BorderWeight.thin;
    WaterfallVarianceTable.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.continuous;
    WaterfallVarianceTable.format.borders.getItem('EdgeLeft').weight = Excel.BorderWeight.thin;
    WaterfallVarianceTable.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.continuous;
    WaterfallVarianceTable.format.borders.getItem('EdgeRight').weight = Excel.BorderWeight.thin;
    await context.sync();
  });
}








//将某个Range中的所有单元格的地址存放到数组中
async function GetRangeAddress(SheetName, TargetRange) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(SheetName);
    const range = sheet.getRange(TargetRange);

    // 第一次只需要知道 Range 的总行数和总列数
    range.load("rowCount,columnCount");
    await context.sync();

    const rowCount = range.rowCount;
    const colCount = range.columnCount;

    // 建立一个二维结构，用来存储每个 cell 对象
    const cellObjs2D = [];

    // 1. 分行列循环，依次对每个单元格 load("address")
    for (let r = 0; r < rowCount; r++) {
      const rowCells = [];
      for (let c = 0; c < colCount; c++) {
        const cell = range.getCell(r, c);
        cell.load("address");
        rowCells.push(cell);
      }
      cellObjs2D.push(rowCells);
    }

    // 2. 等所有 cell 都加载完之后，一次性 sync
    await context.sync();

    // 3. 把每个 cell 的 address 取出来，做成和 TargetRange 一样维度的二维数组
    const addresses2D = [];
    for (let r = 0; r < rowCount; r++) {
      const rowAddresses = [];
      for (let c = 0; c < colCount; c++) {
        rowAddresses.push(cellObjs2D[r][c].address);
      }
      addresses2D.push(rowAddresses);
    }

    // console.log(addresses2D);
    console.log("GetRangeAddress End");
    return addresses2D;
    // addresses2D 的结构类似：
    // [
    //   ["Sheet1!A1", "Sheet1!B1", ...],
    //   ["Sheet1!A2", "Sheet1!B2", ...],
    //   ...
    // ]
  });
}

//将原有的Range 中间删除不定数量的列后，返回删除后的地址，例如A1:G10,删除中间两列返回A1:E10
function getShiftedRangeAfterRemoving(originalRange, ...columnsToRemove) {
  // 1) 如果 originalRange 带有工作表名 (如 "Sheet1!A1:G10")
  //    则提取 sheetName 和 rangePart
  let sheetName = "";
  let rangePart = originalRange;

  // 判断是否包含 "!"
  if (originalRange.includes("!")) {
    const parts = originalRange.split("!");
    sheetName = parts[0];        // 如 "Sheet1"
    rangePart = parts[1];        // 如 "A1:G10"
  }

  // 2) 用正则解析范围部分，如 "A1:G10"
  //    ^([A-Z]+)(\d+):([A-Z]+)(\d+)$
  //    如果范围不匹配，抛出错误
  //   const rangeMatch = rangePart.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    const rangeMatch = rangePart.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
  if (!rangeMatch) {
    throw new Error(
      `Invalid range format. Expected like 'A1:G10' or 'Sheet1!A1:G10', but got '${originalRange}'`
    );
  }

  const startCol = rangeMatch[1];            // "A"
  const startRow = parseInt(rangeMatch[2]);  // 1
  const endCol = rangeMatch[3];              // "G"
  const endRow = parseInt(rangeMatch[4]);    // 10

  // 3) 辅助函数：列字母 -> 数值索引
  function columnToIndex(col) {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      // 'A' = ASCII 65，所以 'A' - 64 = 1, 'B' - 64 = 2, ...
      index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index;
  }

  // 4) 辅助函数：数值索引 -> 列字母
  function indexToColumn(index) {
    let col = "";
    while (index > 0) {
      const remainder = (index - 1) % 26;
      col = toColumnLetter(remainder) + col;
      index = Math.floor((index - 1) / 26);
    }
    return col;
  }

  // 5) 计算原始范围的列索引
  const startColIndex = columnToIndex(startCol); // e.g. A => 1
  const endColIndex = columnToIndex(endCol);     // e.g. G => 7

  // 6) 原有的列宽
  const rangeWidth = endColIndex - startColIndex + 1; // e.g. 7

  // 7) 处理传入的 columnsToRemove，可能包含 "!": 只取列字母
  //    比如 "Sheet1!E" => "E"
  const extractColumnLetters = (colString) => {
    if (colString.includes("!")) {
      // 去掉前面的 sheetName!
      return colString.split("!")[1];
    }
    return colString;
  };

  const removeIndices = columnsToRemove.map((col) => {
    const onlyCol = extractColumnLetters(col);
    return columnToIndex(onlyCol);
  });

  // 8) 计算要删除的列中，确实在 [startColIndex..endColIndex] 范围内的数量
  let removeCount = 0;
  for (const colIndex of removeIndices) {
    if (colIndex >= startColIndex && colIndex <= endColIndex) {
      removeCount++;
    }
  }

  // 9) 新的列宽 = 原始列宽 - 删除列数
  const newWidth = rangeWidth - removeCount;
  if (newWidth <= 0) {
    throw new Error("No columns left after removal!");
  }

  // 10) 新的结束列索引 = 起始列索引 + 新的列宽 - 1
  const newEndColIndex = startColIndex + newWidth - 1;

  // 11) 转回列字母
  const newEndCol = indexToColumn(newEndColIndex);

  // 12) 拼出新的范围地址
  //     如果原始带有表名，则拼回去 "Sheet1!A1:E10"
  const newRangePart = `${startCol}${startRow}:${newEndCol}${endRow}`;
  if (sheetName) {
    return `${sheetName}!${newRangePart}`;
  }
  return newRangePart;
}

async function setFormat(sheetName) {
  await Excel.run(async (context) => {
      try {
          // 获取工作表 Waterfall
          const sheet = context.workbook.worksheets.getItem(sheetName);

          // 获取整个工作表的范围
          const usedRange = sheet.getUsedRange();
          usedRange.format.font.name = "Calibri"; // 设置字体为 Calibri

          await context.sync(); // 同步到 Excel
          console.log("All cells in the Waterfall worksheet are now set to Calibri font.");
      } catch (error) {
          console.error("Error setting font to Calibri:", error);
      }
  });
}

function convertToA1Addresses(cellIndices) {
  // 辅助函数：将列索引转换为列字母
  function indexToColumn(colIndex) {
    let col = "";
    colIndex++; // 转为 1-based
    while (colIndex > 0) {
      const remainder = (colIndex - 1) % 26;
      col = String.fromCharCode(65 + remainder) + col; // 根据 remainder 动态生成字符
      colIndex = Math.floor((colIndex - 1) / 26);
    }
    return col;
  }

  // 判断输入是二维数组还是一维数组
  const isSinglePair = !Array.isArray(cellIndices[0]);

  // 如果是一维数组，包装为二维数组
  const indicesArray = isSinglePair ? [cellIndices] : cellIndices;

  // 转换为 A1 地址
  return indicesArray.map(([rowIndex, colIndex]) => {
    const columnLetter = indexToColumn(colIndex);
    return `${columnLetter}${rowIndex + 1}`; // 转为 A1 样式
  });
}

//-------------提示用户不要修改单元格，在工作表中插入一个长方形----------------
async function DoNotChangeCellWarning(SheetName) {
  await Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem(SheetName);
    var shapes = sheet.shapes;

    // 创建一个长方形
    var rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    rectangle.left = 50;
    rectangle.top = 50;
    rectangle.width = 300;
    rectangle.height = 100;

    // 设置填充颜色为白色
    rectangle.fill.setSolidColor("white");

    // 添加文字并设置颜色为黑色
    var textRange = rectangle.textFrame.textRange;
    textRange.text = "请不要增加或修改单元格的内容";
    textRange.font.color = "red";
    rectangle.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    rectangle.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
    console.log(`${SheetName} Warning for not changing`);

    return context.sync();
  }).catch(function (error) {
    console.log(error);
  });
}


























//----------------检查是否有不合适的运算符---------------------
async function CheckValidOperator() {
  return await Excel.run(async (context) => {
    let Arr = await getColumnAddresses(); //
    console.log("Arr is ");
    console.log(Arr);

    for (let i = 0; i < Arr.length; i++) {
      let InValidOperator = await CheckOperator("Data", Arr[i]); // 如果返回false，说明有不合法的运算符号
      console.log("InValidOperator is " + InValidOperator);
      // 碰到第一个不合法的运算符号就退出循环 
      if (InValidOperator) {
         return true;
      }else{
        return false;
      }

    }
  });
}


// 获取"SumY", "SumN", "Result"对应的单元格的地址，为后面的运算符判断做准备
async function getColumnAddresses() {
  try {
    return await Excel.run(async (context) => {
      console.log("getColumnAddress 1");
      const sheet = context.workbook.worksheets.getItem("Data");
      const usedRange = sheet.getUsedRange().getRowsAbove(-200);
      usedRange.load(["values", "columnCount", "rowCount"]);
      await context.sync();
      console.log("getColumnAddress 2");

      const headerRow = usedRange.values[0];
      const targetColumns = ["SumY", "SumN", "Result"];
      let currentColumnRange = null;
      let columnRanges = [];
      let tempRange = [];
      console.log("getColumnAddress 3");

      // 逐列检查头部内容，判断是否匹配目标字段
      for (let i = 0; i < headerRow.length; i++) {
        const header = headerRow[i];
        if (targetColumns.includes(header)) {
          // 如果是目标字段，检查是否是连续的
          if (currentColumnRange === null || currentColumnRange === i - 1) {
            // 连续列，添加到临时范围
            tempRange.push(i);
            currentColumnRange = i;
          } else {
            // 不连续列，把之前的范围存入数组，并清空临时范围
            if (tempRange.length > 0) {
              columnRanges.push(tempRange);
            }
            tempRange = [i];
            currentColumnRange = i;
          }
        }
      }

      // 最后一组连续列需要额外保存
      if (tempRange.length > 0) {
        columnRanges.push(tempRange);
      }
      console.log("getColumnAddress 4");
      // 打印结果，将连续列存入数组中
      let addressArray = [];
      for (let colRange of columnRanges) {
        if (colRange.length === 1) {
          // 单列不连续，直接返回列地址
          addressArray.push(`${toColumnLetter(colRange[0])}1:${toColumnLetter(colRange[0])}${usedRange.rowCount}`);
        } else {
          // 连续列，返回范围地址
          let startCol = toColumnLetter(colRange[0]);
          let endCol = toColumnLetter(colRange[colRange.length - 1]);
          addressArray.push(`${startCol}1:${endCol}${usedRange.rowCount}`);
        }
      }

      console.log(addressArray); // 在控制台输出地址数组
      return addressArray;
    });
  } catch (error) {
    console.error("Error:", error);
  }
}

//-------------检测单元格是否有"$", "=", "+", "-", "*", "/", "(", ")" 以外的符号，如果有则跳出提示窗口------------
async function CheckOperator(TargetSheetName, TargetRange) {
  return await Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem(TargetSheetName);
    var range = sheet.getRange(TargetRange);
    range.load("formulas,values,rowCount,columnCount");

    await context.sync();

    var hasInvalidCharacters = false;

    // 从第三行开始检测
    for (let colIndex = 0; colIndex < range.columnCount; colIndex++) {
      for (let rowIndex = 2; rowIndex < range.rowCount; rowIndex++) {

        let rowLetter = toColumnLetter(rowIndex);
        let columnLetter = toColumnLetter(colIndex);
        const formula = range.formulas[rowIndex][colIndex];
        const value = range.values[rowIndex][colIndex];

        // 如果单元格是数字，直接返回true
        if (formula === value) {
          // console.log(`单元格 ${TargetRange} 的 ${rowIndex + 1}, ${colIndex + 1} 为数字`);
          hasInvalidCharacters = false;
          continue;
        }
        // console.log("CheckOperator 2");
        // 如果是公式，检查是否包含无效字符
        if (formula) {
          let cleanedFormula = formula.replace(/[\$\=\+\-\*\/\(\)]/g, ""); // 移除允许的符号
          // console.log("公式清理后: " + cleanedFormula);

          // 移除单元格引用的字符（例如 A2, AD10 等）
          cleanedFormula = cleanedFormula.replace(/[A-Za-z]+\d+/g, "");

          // 新增步骤：移除所有单纯的数字
          cleanedFormula = cleanedFormula.replace(/\b\d+\b/g, "");
          // console.log("清理后的公式: " + cleanedFormula);
          // console.log("CheckOperator 3")
          // 检查是否还有其他字符
          if (cleanedFormula.trim().length > 0) {
            hasInvalidCharacters = true;
            console.log(`${range.values[1][colIndex]}一列的公式 "${formula}" 包含无效字符: "${cleanedFormula}"`);

          // 在网页上显示警告
          const warningMessage = `${range.values[1][colIndex]}一列的公式 "${formula}" 包含无效字符: "${cleanedFormula}"`;
          const keyWarningPrompt = document.getElementById("keyWarningPrompt");
          const modalOverlay = document.getElementById("modalOverlay");
          const container = document.querySelector(".container");

          // 动态更新警告消息
          const warningElement = document.querySelector("#keyWarningPrompt .waterfall-message");
          warningElement.textContent = warningMessage;

          // 显示模态遮罩和提示框
          modalOverlay.style.display = "block";
          keyWarningPrompt.style.display = "flex";
          container.classList.add("disabled");

          // 等待用户点击确认按钮
          await new Promise((resolve) => {
            const confirmButton = document.getElementById("confirmKeyWarning");

            confirmButton.addEventListener(
              "click",
              function () {
                keyWarningPrompt.style.display = "none";
                modalOverlay.style.display = "none";
                container.classList.remove("disabled");
                resolve();
              },
              { once: true } // 确保事件只触发一次
            );
          });

            break;
          }
        }
      }

      if (hasInvalidCharacters) {
        break;
      }
    }


    if (!hasInvalidCharacters) {
      console.log(`所有公式中均未发现无效字符，且没有单元格为数字。`);
      return true;
    } else {
      console.log("已发现包含无效字符的公式或数字，详情请查看控制台输出。");
      return false;
    }
  }).catch(function (error) {
    console.log(error);
  });
}

// 检查第Data第二行是否存在重复字段
async function CheckDuplicateHeaders() {
  return await Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Data");
    var usedRange = sheet.getUsedRange();
    usedRange.load("values,rowCount,columnCount");

    await context.sync();

    if (usedRange.rowCount < 2) {
      console.log("数据行不足两行，无法检查重复字段。");
      return;
    }

    let secondRow = usedRange.values[1]; // 第二行索引为1
    let seenHeaders = new Set();
    let duplicateHeaders = new Set();

    for (let colIndex = 0; colIndex < usedRange.columnCount; colIndex++) {
      let header = secondRow[colIndex];
      if (seenHeaders.has(header)) {
        duplicateHeaders.add(header);
      } else {
        seenHeaders.add(header);
      }
    }

    if (duplicateHeaders.size > 0) {
      let duplicatesList = Array.from(duplicateHeaders).join(", ");
      console.log(`第二行的标题中存在同样的字段名: ${duplicatesList}，请修改`);

      // 在网页上显示警告
      const warningMessage = `第二行的标题中存在同样的字段名: ${duplicatesList}，请修改`;
      const keyWarningPrompt = document.getElementById("keyWarningPrompt");
      const modalOverlay = document.getElementById("modalOverlay");
      const container = document.querySelector(".container");

      // 更新警告消息内容
      const warningElement = document.querySelector("#keyWarningPrompt .waterfall-message");
      warningElement.textContent = warningMessage;

      // 显示模态遮罩和提示框
      modalOverlay.style.display = "block";
      keyWarningPrompt.style.display = "flex";
      container.classList.add("disabled");

      // 等待用户点击确认按钮
      await new Promise((resolve) => {
        const confirmButton = document.getElementById("confirmKeyWarning");

        confirmButton.addEventListener(
          "click",
          function () {
            keyWarningPrompt.style.display = "none";
            modalOverlay.style.display = "none";
            container.classList.remove("disabled");
            resolve(); // 继续 Promise
          },
          { once: true } // 确保事件只触发一次
        );
      });

      return false;
    } else {
      console.log("第二行标题中没有重复字段。");
      return true;
    }
  }).catch(function (error) {
    console.log(error);
  });
}


// 分解Bridge data 中result的公式，创建FormulasBreakdown，并在其中分解，并复制到BridgeData
async function FormulaBreakDown2() {
  return await Excel.run(async (context) => {
    console.log("Enter FormulaBreakDown2");

    // 获取工作簿中所有工作表
    const sheets = context.workbook.worksheets;
    // 尝试获取名为 "FormulasBreakdown" 的工作表，如果不存在则返回 null 对象
    const targetSheet = sheets.getItemOrNullObject("FormulasBreakdown");
    
    // 同步加载 targetSheet 的 isNullObject 属性
    await context.sync();

    // 如果工作表不存在，则创建该工作表
    if (targetSheet.isNullObject) {
      await copyAndModifySheet("Bridge Data", "FormulasBreakdown"); // 创建FormulasBreakdown工作表
      console.log("FormulasBreakdown 工作表已创建。");
    } else {
      console.log("FormulasBreakdown 工作表已存在，跳过创建。");

    }

    // await copyAndModifySheet("Bridge Data", "FormulasBreakdown"); // 创建FormulasBreakdown工作表
    let FormulaSheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    let FormulaRange = FormulaSheet.getUsedRange();
    let FirstRow = FormulaRange.getRow(0); // 获取Type行，找到Result
    FirstRow.load("address,values");
    await context.sync();

    console.log("FirstRow.address is " + FirstRow.address);

    // 找到result单元格
    let ResultType = FirstRow.find("Result", {
      completeMatch: true,
      matchCase: true,
      searchDirection: "Forward"
    });
    ResultType.load("address");
    await context.sync();

    console.log("ResultType is " + ResultType.address);

    //往下两行，获得Result对应的公式
    let ResultCell = ResultType.getOffsetRange(2, 0);
    ResultCell.load("address,formulas");
    let ResultTitle =ResultType.getOffsetRange(1, 0);
    ResultTitle.load("values");
    await context.sync();
    console.log("ResultTitle is " + ResultTitle.values[0][0]);

    const newValue = ResultTitle.values[0][0];
    if (!ArrVarPartsForPivotTable.includes(newValue)) { // 保证是唯一的变量
      ArrVarPartsForPivotTable.push(newValue); // 给数据透视表筛选变量
    }
    // ArrVarPartsForPivotTable.push(ResultTitle.values[0][0]); //给数据透视表筛选变量

    await FindNextFormulas(ResultCell.address); // 1>>>>>查找公式单元格中是否还有进一步引用的公式, 并最终反应在第一个单元格中, 这里已经去掉$固定符号

    //下面需要重新load 一次，不然后面的代码不知道上一步已经改变了单元格内容。
    ResultCell.load("address,formulas");
    await context.sync();

  //--------------------------临时-------------------
  // let FormulaSheet = context.workbook.worksheets.getItem("FormulasBreakdown 2");
  // let ResultCell = FormulaSheet.getRange("S3");
  // ResultCell.load("formulas");
  // await context.sync();
  //--------------------------临时------------------
  let formula = ResultCell.formulas[0][0].replace(/=|\$/g, ""); //删除所有的固定符号$和=号

  //清除掉+---等多余的正负号
  formula = CleanOperator(formula);
  console.log("after clean operator, formula is ");
  console.log(formula);
  
  ////取出掉公式里没有必要的括号，返回的formula没有"="号
  formula = await removeUnnecessaryParentheses(formula); 
  console.log("Remove ParenTheses is :" + formula);
  //   ResultCell.formulas = [[formula]];
  //   await context.sync();

  //清除掉+---等多余的正负号
  // formula = CleanOperator(formula);
  // console.log("after clean operator, formula is ");
  // console.log(formula);

  //-------建立一个数组包含所有变量元素和运算符的对象，
  //包含公式里的变量名，变量对应的单元格地址，是否是运算符号，是否是SumY, 是否是replace, others 表示其他情况，其他的自定义变量？？
  // let FormulaTokens = [];

  //第一次先将原始formula中的所有Tokens放入对象中，这里还不包含replace的情况
  let tokenPattern = /([+\-*/()])/; 
  let tokens = formula.split(tokenPattern).map(t => t.trim()).filter(Boolean); //分解除公式中的所有元素
  console.log("tokens is");
  console.log(tokens);
  
  for(let i = 0; i<tokens.length; i++){
      let CellAddress = tokens[i].trim().match(/^([A-Za-z]+)(\d+)$/); //匹配A1样式变量
      let NumberPattern = tokens[i].trim().match(/^(\d+|\.\d+)(\.\d+)?%?$/); //匹配整数、小数、百分数以及省略整数部分的小数
      let isOperator = /[+\-*/()]/.test(tokens[i]); //测试是否是运算符号, 如果不是返回false
      let ReplaceFactor = tokens[i].trim().match(/^__replace__(\d+)$/); //ReplaceFactor[0]: 整个匹配到的字符串,ReplaceFactor[1]捕获组 (\d+) 提取的数值部分
      let SumType = null;
      if (CellAddress) { //如果是A1格式，则需要去第一行找是否是SumY
          console.log("CellAddress is " + CellAddress);
          let colLetters = CellAddress[1]; // "A" / "B" / "C"
          console.log("colLetters is " + colLetters);
          // 构造我们要去读取的 “类型定义单元格”: 比如 A3 -> A1
          let typeCellAddress = colLetters + "1";
          let titleCellAddress = colLetters + "2";
          let valueCellAddress = colLetters + "3";
          
          console.log("typeCellAddress is " + typeCellAddress);
          console.log("titleCellAddress is " + titleCellAddress);
          console.log("valueCellAddress is " + valueCellAddress);

          let typeCell = FormulaSheet.getRange(typeCellAddress);
          let titleCell = FormulaSheet.getRange(titleCellAddress);
          let valueCell = FormulaSheet.getRange(valueCellAddress);
          typeCell.load("values");
          titleCell.load("values");
          valueCell.load("values");
          await context.sync();

          let cellValue = typeCell.values[0][0];
          let titleValue = titleCell.values[0][0];
          // 假设用户确保在 A1, B1, C1 等位置写的就是 "SumY" 或 "SumN"
          console.log("cellValue is " + cellValue);
          console.log("titleValue is " + titleValue);
          // 判断是SumY还是SumN ******目前还是"SumN"，这里要修改成SumY 和 SumN最后的状态，因此下面的if判断不需要
          if(cellValue === "SumY"){
            SumType = "SumY";  // 把数组中的原变量替换成type，是否是SumY或者是SumN
          }else if(cellValue === "SumN"){
            SumType = "SumN";
          }

          //存放入FormulaTokens中
          FormulaTokens.push({
              Token: tokens[i],
              TokenName: titleValue,
              SumType: SumType,
              ReplaceIndex:null,
              isCell: true,
              isOperator: isOperator,
              isNumber: false,
              isReplace:false,
              isOthers: false,
              TermToReplace: null
          });


      }
      else if (NumberPattern){ // 这里判断如果是数字的话，例如10, 0.5, 或者0.5%等
          console.log("NumberPattern is " + NumberPattern[0]);
          // 如果是单纯的数字，那么这里先认为是SumY类型的 ****这里需要进一步思考
          //存放入FormulaTokens中
          FormulaTokens.push({
              Token: tokens[i],
              TokenName: titleValue,
              SumType: "SumY",
              ReplaceIndex:null,
              isCell: false,
              isOperator: isOperator,
              isNumber: true,
              isReplace:false,
              isOthers: false,
              TermToReplace: null
          });
      }
      else if(isOperator){
          //存放入FormulaTokens中
          FormulaTokens.push({
            Token: tokens[i],
            TokenName: tokens[i],
            SumType: null,
            ReplaceIndex:null,
            isCell: false,
            isOperator: isOperator,
            isNumber: false,
            isReplace:false,
            isOthers: false,
            TermToReplace: null
        });

      }
      else{ //应该是自定义变量ABC等情况，但是理论上不允许

          //存放入FormulaTokens中
          FormulaTokens.push({
              Token: tokens[i],
              TokenName: titleValue,
              SumType: "SumY",
              ReplaceIndex:null,
              isCell: false,
              isOperator: isOperator,
              isNumber: false,
              isReplace:false,
              isOthers: true,
              TermToReplace: null
          });

      }
      
  }

  console.log("First time put in FormulaTokens:")
  console.log(JSON.stringify(FormulaTokens, null, 2));

  let innerMostParenthesesRegex = /\(([^()]*)\)/g; // match[0] 包含括号，match[1] 不包含括号，仅括号内内容
  let match;
  let ReplaceIndex = 0; //用来替换formula中的部分公式，不然的的话永远不能替换到最外层括号，导致无线循环
  formula = `(${formula})`; //必须要在循环前加入括号而且要替代原有的formula，
  while ((match = innerMostParenthesesRegex.exec(formula)) !== null) {
      //match[0] 包含括号，match[1] 不包含括号，仅括号内内容
      let MatchTokens = match[1].split(tokenPattern).map(t => t.trim()).filter(Boolean); //分解除公式中的所有元素
      console.log("MatchTokens is" + MatchTokens);
      console.log("match[1] is " + match[1]);
       /**
        * 子函数1：扫描公式中的连续乘除链，
        * 如果该链中所有运算符均为除法（"/"），则调用 Func1。
        */
      
       //下面返回found是否有连除公式，以及经过处理的公式subFormula
      let PureDivisionResult = processPureDivisionSegments(match[1],FormulaTokens); 
      if(PureDivisionResult.found){
        //进到if里面说明有连除，则需要处理更新后的SubFormula 计算结果是SumY还是SumN，并且替换到原始formula的部分，并放入对象
          console.log("连续除法 formula is " + PureDivisionResult.formula);
          let SumResult = checkFormulaResultType(PureDivisionResult.formula,FormulaTokens);
          console.log("连续除法 SubResult is " + SumResult);

          FormulaTokens.push({
            Token: `__Replace__M${ReplaceIndex}`,
            TokenName: `__Replace__M${ReplaceIndex}`,
            SumType: SumResult,
            ReplaceIndex:ReplaceIndex,
            isCell: false,
            isOperator: false,
            isNumber: false,
            isReplace:true,
            isOthers: false,
            TermToReplace: `(${PureDivisionResult.formula})` //被替换的经过调整的公式,需要单独加上括号，保证和match[0]一样有括号
        });
          // match[0]包括括号，如果进入到这里说明SumY的结果已经全部求完，包括括号都可以替换
          console.log("match[0] in pureDivision is " + match[0]);
          formula = formula.replace(match[0],`__Replace__M${ReplaceIndex}`) // match[0]
          console.log("after pureDivision formula is " + formula);
          console.log(JSON.stringify(FormulaTokens, null, 2));
          ReplaceIndex++; //必须放在每一次替换执行后面添加
          innerMostParenthesesRegex.lastIndex = 0; //每一次while 循环后添加

          continue; //每一次的连除等类型判断完后，直接进入下一轮while循环，后面的循环代码不需要
      }

      //如果上面的连除没有找到，没有进入if则继续判断是否有连乘和乘除连续的公式部分
      let MixMulDivResult = processMixedMulDivSegments(match[1], FormulaTokens);
      console.log("MixMulDivResult.formula is " + MixMulDivResult.formula);
      if(MixMulDivResult.found){

        console.log("连续乘除 formula is " + MixMulDivResult.formula);
        let SumResult = checkFormulaResultType(MixMulDivResult.formula, FormulaTokens);
        console.log("连续乘除 SubResult is " + SumResult);

        FormulaTokens.push({
          Token: `__Replace__M${ReplaceIndex}`,
          TokenName: `__Replace__M${ReplaceIndex}`,
          SumType: SumResult,
          ReplaceIndex: ReplaceIndex,
          isCell: false,
          isOperator: false,
          isNumber: false,
          isReplace: true,
          isOthers: false,
          TermToReplace: `(${MixMulDivResult.formula})`
        });

        // match[0]包括括号，如果进入到这里说明SumY的结果已经全部求完，包括括号都可以替换
        console.log("match[0] in MixedMulDiv is " + match[0]);
        formula = formula.replace(match[0], `__Replace__M${ReplaceIndex}`) // match[0]
        console.log("after MixedMulDiv formula is " + formula);
        console.log(JSON.stringify(FormulaTokens, null, 2));
        ReplaceIndex++; //必须放在每一次替换执行后面添加
        innerMostParenthesesRegex.lastIndex = 0; //每一次while 循环后添加

        continue; //每一次的连除等类型判断完后，直接进入下一轮while循环，后面的循环代码不需要

      }

      //乘除法都解决完了以后，就剩下加减法，调用计算SumY的函数执行, 然后直接替换
      let SumResult = checkFormulaResultType(match[1],FormulaTokens);
      FormulaTokens.push({
        Token: `__Replace__M${ReplaceIndex}`,
        TokenName: `__Replace__M${ReplaceIndex}`,
        SumType: SumResult,
        ReplaceIndex: ReplaceIndex,
        isCell: false,
        isOperator: false,
        isNumber: false,
        isReplace: true,
        isOthers: false,
        TermToReplace: match[0] // 直接用match[0]就应该可以，因为match[0]带括号
      });
      

      console.log("match[0] normal is " + match[0]);
      formula = formula.replace(match[0], `__Replace__M${ReplaceIndex}`) // match[0]
      console.log("after MixedMulDiv formula is " + formula);
      console.log(JSON.stringify(FormulaTokens, null, 2));
      ReplaceIndex++; //必须放在每一次替换执行后面添加
      innerMostParenthesesRegex.lastIndex = 0; //每一次while 循环后添加
  }

  //while循环结束以后，将replace后的字符串逆向替换会公式
  formula = restoreFinalFormula(formula, FormulaTokens);
  console.log("Last formula is" + formula);

  //所有的连除，乘除混合，排序都已经完成，剩下清除0-A-B这个步骤
  formula = simplifyExpression(formula)
  formula = formula.replace(" ","");
  console.log("在清除零之后：" + formula);

  ResultCell.formulas= [[`=${formula}`]];//放回单元格
  await context.sync();

  await processFormulaObj(ResultCell.address); // 不返回任何的值，函数修改成完全为了数据透视表筛选数据用
  
  //如果已经在检测公式checkType2的循环中了，则不需要再进一步检测
  // console.log("checkFormulaGlobalVar 2 is ");
  // console.log(checkFormulaGlobalVar);
  // if(!checkFormulaGlobalVar){
  let result = await CheckFormula(ResultCell.address); //检查并处理公式中的特殊部分
  // 如果第一次返回的是 Error，则直接退出
  if(result === "Error") {
    console.log("CheckFormula 返回 Error，终止 main 函数执行");
    return "Error";
}

// 当返回 "CheckType_1_or_2" 时，循环调用 CheckFormula 直到不再返回 "CheckType_1_or_2"
while(result === "CheckType_1_or_2") {
  console.log("CheckType_1_or_2 类型")
  result = await CheckFormula(ResultCell.address);
  if(result === "Error") {
    console.log("CheckFormula 返回 Error，终止 main 函数执行");
    return "Error";
  }
}
  // }

  // 不返回任何的值，函数修改成完全为了数据透视表筛选数据用, 再执行一次，防止上一步有的变量，例如SumN+SumN 中的变量有单独在别的地方被引用但是上一步从数透表变量中删除了
    await processFormulaObj(ResultCell.address); 

  await processFormula(ResultCell.address); //2>>>>>>>>>>> 对公式里的运算符和优先级，从左到右加上括号
  await SplitFormula(ResultCell.address); //3

  });
}


/**
 * 根据 FormulaTokens 中的替换记录，递归替换 finalFormula 中的 token
 * 直到 finalFormula 中不再包含任何 token 为止
 *
 * @param {string} finalFormula - 替换后的公式字符串，可能还包含 token 字符串
 * @param {Array} formulaTokens - 包含替换信息的数组，每个对象必须具有 Token 和 TermToReplace 属性
 * @returns {string} - 最终还原成原始公式的字符串
 */
function restoreFinalFormula(finalFormula, formulaTokens) {
  // 标识是否在当前循环中进行了替换
  let hasToken = true;

  // 循环直到没有匹配的 token 出现
  while (hasToken) {
    // 每次循环开始时先将标识置为 false
    hasToken = false;

    // 遍历所有的替换记录
    formulaTokens.forEach(tokenObj => {
      // 判断当前公式字符串是否包含 tokenObj.Token 这个子字符串
      if (tokenObj.isReplace === true && finalFormula.indexOf(tokenObj.Token) !== -1) {
        console.log("finalFormula is "+ finalFormula);
        console.log("tokenObj.Token is " + tokenObj.Token);
        // 由于 token 可能包含正则表达式中具有特殊含义的字符，
        // 因此需要对 tokenObj.Token 进行转义，
        // 转义后的字符串才能安全用于构造正则表达式进行匹配替换
        const escapedToken = tokenObj.Token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        // 构造一个全局匹配的正则表达式，确保公式中所有出现该 token 的地方都能被替换
        const regExp = new RegExp(escapedToken, 'g');

        // 用 tokenObj.TermToReplace 替换所有匹配的 token
        finalFormula = finalFormula.replace(regExp, tokenObj.TermToReplace);

        // 标记本次循环中有进行替换操作，以便继续下一轮检查
        hasToken = true;
      }
    });
  }
  return finalFormula;
}




//取出掉公式里没有必要的括号
function removeUnnecessaryParentheses(formula) {
  const precedence = {
    '+': 1,
    '-': 1,
    '*': 2,
    '/': 2
  };

  console.log("Remove formula is " + formula);

  let tempFormulas = {}; // 用于存储不能去掉的括号及其公式
  let formulaCounter = 1;

  function getPrecedence(op) {
    return precedence[op] || 0;
  }

  let innerMostParenthesesRegex = /\([^()]*\)/g; // 找到最内层的括号
  let match;

  console.log("Remove formula 1 is " + formula);

  //`(${formula})` 这里再公式的最外边临时加上一对括号，让最外层至少也循环一次，避免没有括号不循环的情况
  while ((match = innerMostParenthesesRegex.exec(formula)) !== null) {
    let innerExpr = match[0]; //包括括号
    let innerContent = innerExpr.slice(1, -1); // 去掉括号获取内部内容
    console.log("Remove formula 1.1 is " + formula);
    console.log("match[0] is " + match[0]);
    console.log("match.index is " + match.index);
    console.log("innerContent is " + innerContent);
    console.log("/^\d+$/.test(innerContent) is " + /^\d+$/.test(innerContent));

    const numberPattern = /^(\d+|\.\d+)(\.\d+)?%?$/; //匹配整数、小数、百分数以及省略整数部分的小数, 不支持前面有运算符号
    // const numberPattern2 = /^[+-]*(\d+|\.\d+)(\.\d+)?%?$/; //匹配整数、小数、百分数以及省略整数部分的小数,并支持前面任意数量的+号和-号
    // 正则表达式匹配单元格（例如 A10、AA34 或 abc123）
    const cellPattern = /^[A-Za-z]+[0-9]+$/;
    let operators = innerContent.match(/[+\-*/]/g);
    let canRemove = false;
    // if (!(numberPattern.test(innerContent) || cellPattern.test(innerContent))){ //判断如果不是（10）这种单纯的数字 或者(Q3）才往下
    if(operators){ 
          // 查找括号X内部的运算符优先级M 这里如果是单纯的(10)则不会返回任何的operators, 程序无法进行下去，需要修改
          //已经修改为包含运算符号才进行下面的判断，如果不包含运算符号，则直接取消括号，因为可能会有各种变量名等不包含括号
          // let operators = innerContent.match(/[+\-*/]/g);
          let M = Math.min(...operators.map(getPrecedence)); //这里取得括号里优先级最低的运算符作为M，来对比左边和右边的运算符

          console.log("M is " + M);
          // 查找括号X左边和右边的运算符
          let leftPart = formula.slice(0, match.index).trim();
          let rightPart = formula.slice(match.index + innerExpr.length).trim();

          console.log("leftPart is " + leftPart);
          console.log("rightPart is " + rightPart);

          let L = leftPart ? getPrecedence(leftPart[leftPart.length - 1]) : null;
          let R = rightPart ? getPrecedence(rightPart[0]) : null;

          console.log("L is " + L);
          console.log("R is " + R);

          // 判断左边和右边是否为运算符
          let isLeftOperator = L !== null && precedence.hasOwnProperty(leftPart[leftPart.length - 1]);
          let isRightOperator = R !== null && precedence.hasOwnProperty(rightPart[0]);


          // 2. 如果括号X的相邻左边和相邻右边都是括号，去掉X
          if (leftPart && rightPart &&
            leftPart[leftPart.length - 1] === '(' && rightPart[0] === ')') {
            canRemove = true;
            console.log("2. 如果括号X的相邻左边和相邻右边都是括号，去掉X")
          }

          console.log("Remove formula 1.6 is " + formula);
          
          // 3-1. 左边是运算符且M优先级高于L
          if (!canRemove && isLeftOperator && M > L) {
            canRemove = true;
            console.log(" 3-1. 左边是运算符且M优先级高于L");
          }
          // 3-2. 右边是运算符且M优先级高于R
          else if (!canRemove && isRightOperator && M > R) {
            canRemove = true;
            console.log("3-2. 右边是运算符且M优先级高于R");
          }

          // 3-3. 如果M优先级等于L，并且等于R（如果R存在）, 如果L是除号，则括号不去掉
          else if (!canRemove && isLeftOperator && M === L && (!isRightOperator || M === R)) {
            // if (['-', '/'].includes(leftPart[leftPart.length - 1])) {
            //   // 3-3-1. L是- 或 / 号，内部符号需要反转
            //   innerContent = innerContent.replace(/[+\-*/]/g, function (op) {
            //     return { '+': '-', '-': '+', '*': '/', '/': '*' }[op];
            //   });
              
            if (['-'].includes(leftPart[leftPart.length - 1])) {
              // 3-3-1. L是- 或 / 号，内部符号需要反转，这里 - 和 / 必须要分开成两部分，
              innerContent = innerContent.replace(/[+\-*/]/g, function (op) {
                console.log("3-3-1. L是-号，内部符号需要反转");
                return { '+': '-', '-': '+' }[op];
              });      


              canRemove = true;
            } 
            else if (['/'].includes(leftPart[leftPart.length - 1])) {
              // 3-3-1. L是- 或 / 号，内部符号需要反转，这里 - 和 / 必须要分开成两部分
              innerContent = innerContent.replace(/[+\-*/]/g, function (op) {
                console.log("3-3-2. L是 / 号，内部符号需要反转");
                return { '*': '/', '/': '*' }[op];
              });      


              canRemove = true;
            } 
            else if (['+', '*'].includes(leftPart[leftPart.length - 1])) {
              // 3-3-2. L是+ 或 * 号，直接去掉括号X
              canRemove = true;
              console.log("// 3-3-3. L是+ 或 * 号，直接去掉括号X");
            }
          }

          // 3-4. 如果右边是运算符且M优先级等于R，并且也等于L（左边没有运算符则不需要比较L)
          else if (!canRemove && isRightOperator && M === R && (!isLeftOperator || M === L)) {
            canRemove = true;
            console.log(" 3-4. 如果右边是运算符且M优先级等于R，并且也等于L");
          }

          // 3-5. 如果括号X的左边和右边都没有字符了，可以直接去掉
          else if (!canRemove && !leftPart && !rightPart) {
            canRemove = true;
            console.log("3-5. 如果括号X的左边和右边都没有字符了");
          }

          console.log("Remove formula 2 is " + formula);

      
    }else{  //如果是（10）这样的纯数字，则直接去掉括号

      canRemove = true;

    }

    console.log("canRemove is " + canRemove);

    if (canRemove) {
      // 去掉括号，替换公式中的部分
      formula = formula.slice(0, match.index) + innerContent + formula.slice(match.index + innerExpr.length);
    } else {
      // 4. 括号不能去掉，将其替换为键并存入TempFormulas
      let key = `_M${formulaCounter++}_`;
      tempFormulas[key] = innerExpr;  // 存储的是包括括号在内的完整表达式
      formula = formula.slice(0, match.index) + key + formula.slice(match.index + innerExpr.length);
    }

    // 重置正则表达式的搜索位置
    innerMostParenthesesRegex.lastIndex = 0;
  }

  console.log("Remove formula 3 is " + formula);

  // 7. 从TempFormulas的最后一个开始替换
  let keys = Object.keys(tempFormulas).reverse();
  keys.forEach(key => {
    formula = formula.replace(key, tempFormulas[key]);
  });

  // formula = "=" + formula

  console.log("Remove End formula is" + formula);
  return formula;
}


//////////////////////////////-----------------CleanOperator---------------Start----------//////////////////////////////////
///////////////////////////////////////////////////////////////
// （A）对外入口：给定表达式字符串，返回 "SumY" / "SumN" 或错误
//////////////////////////////////////////////////////////////
function CleanOperator(expression) {
  let formula = expression.replace("=","");

  let innerMostParenthesesRegex = /\([^()]*\)/g; // 找到最内层的括号
  let match;
  let i = 1;
  let ReplaceFactor = {}; //存放分解替换的公式括号里的内容 
  while ((match = innerMostParenthesesRegex.exec(formula)) !== null) {
      let innerExpr = match[0];
      let innerContent = innerExpr.slice(1, -1); // 去掉括号获取内部内容
      console.log("innerExpr is " + innerExpr);
      console.log("innerContent is " + innerContent);
      let AdjFormula = CleanOperatorExe(innerContent); //清理+-符号
      console.log("after clean formula is " + AdjFormula);
      
      //将+-号清除号的公式存入对象中
      let key = `__replace__${i}`;
      ReplaceFactor[key] = "(" +AdjFormula+ ")"; //要把括号也存进Key中
      //替换原公式的括号的内容为Key
      // 这里要替换带括号的部分，不能替换括号里的部分，会造成无限循环
      formula = formula.replace(innerExpr,key); 

      // 重置正则表达式的搜索位置
      innerMostParenthesesRegex.lastIndex = 0;
      i++;
  }
  formula = CleanOperatorExe(formula); //在最外层可能没有括号，单独使用一次

  //ReplaceFactor的最后一个开始替换
  let keys = Object.keys(ReplaceFactor).reverse();
  keys.forEach(key => {
      formula = formula.replace(key, ReplaceFactor[key]);
  });

  // formula = "=" + formula

  console.log("After clean final formula is" + formula);
  return formula;

}


function CleanOperatorExe(expression) {
  try {
    // 1) 预处理、拆分出初步 Token
    let tokens = tokenizeFormula(expression);

    // 2) 将“一元符号”转成括号形式  (0 - X) / (X)
    tokens = transformUnarySigns(tokens);
    console.log("tokens is ");
    console.log(tokens);
    return tokens;
  } catch (err) {
    console.error(err);
    return "Error";
  }
}

////////////////////////////////////////////////////////////////////
//// （B）tokenizeExpression:
////     给运算符前后插入空格，再 split，得到初步的 Token 数组
////     e.g. "SumY*---SumN" -> ["SumY","*","-","-","-","SumN"]
///////////////////////////////////////////////////////////////////
function tokenizeExpression(formula) {
  // 给 + - * / 前后插入空格
  let tokenPattern = /([+\-*/()])/;
  return formula.split(tokenPattern).map(t => t.trim()).filter(Boolean);
}

//////////////////////////////////////////////////////////////
///// （C）transformUnarySigns:
/////     扫描 tokens，将连续的 +/− (一元符号) 转成括号表达式
/////    e.g. ["SumY","*","-","-","-","SumN"]
////       -> 遇到 "*","-": 说明这是一元符 => 3次负 => "(0 - SumN)"
////      -> 最终 => ["SumY","*","(","0","-","SumN",")"]
///////////////////////////////////////////////////////////////
function transformUnarySigns(tokens) {
  let result = [];
  let i = 0;

  while (i < tokens.length) {
    const t = tokens[i];

    if ((t === '+' || t === '-')) {
      // 判断是否是这种情况A---B 或者是 )--+B 或者 )--+-(
      // if (result.length === 0 || isOperator(result[result.length - 1])) {
      if(i !== 0 && tokens[i-1] !=="(" && tokens[i-1] !=="+" && tokens[i-1] !=="-" && tokens[i-1] !=="*" && tokens[i-1] !=="/"){
          // 连续读取后面的 +/-，直到遇到非 +/- 为止 
          let minusCount = 0;
          while ((tokens[i] === '+' || tokens[i] === '-') && i < tokens.length) {
              if (tokens[i] === '-') minusCount++;
              i++;
          }
          // minusCount 为出现的 '-' 数量 (奇数=>负,偶数=>正)
          if (i >= tokens.length) {
              // 表达式不完整 => 报错/直接返回
              break;
          }
          const isNegative = (minusCount % 2 === 1);
          const operand = tokens[i]; // 接下来的操作数 (可能是 "SumY"/"SumN" 或 "(")

          // 我们用括号包裹: (0 - 操作数) 或 (操作数)
          if (isNegative) {
              // 一元负 => - X

              result.push("-");

          } else {
              // 一元正 => + X 
              result.push("+");
          }
          // i++;
      }
      // 判断是否是一元符号？如果是第一个字符，或者后面的符号是+或者-号,考察这种情况A*--B 或者(---B)
      else if(result.length === 0 || ((i+1)<tokens.length && (tokens[i+1] ==='+' || tokens[i+1] ==='-'))) {
        // 连续读取后面的 +/-，直到遇到非 +/- 为止 
        let minusCount = 0;
        while ((tokens[i] === '+' || tokens[i] === '-') && i < tokens.length) {
          if (tokens[i] === '-') minusCount++;
          i++;
        }
        // minusCount 为出现的 '-' 数量 (奇数=>负,偶数=>正)
        if (i >= tokens.length) {
          // 表达式不完整 => 报错/直接返回
          break;
        }
        const isNegative = (minusCount % 2 === 1);
        const operand = tokens[i]; // 接下来的操作数 (可能是 "SumY"/"SumN" 或 "(")

        // 我们用括号包裹: (0 - 操作数) 或 (操作数)
        if (isNegative) {
          // 一元负 => (0 - X)
          result.push("(");
          result.push("0");
          result.push("-");
          result.push(operand);
          result.push(")");
        } else {
          // 一元正 => ( X )
          result.push("(");
          result.push(operand);
          result.push(")");
        }
        i++;
      } 
      //如果都不是以上情况，那么就直接放入
      else {
        // 二元运算符 => 直接放入
        result.push(t);
        i++;
      }
    } else {
      // 正常 token => 直接放入结果
      result.push(t);
      i++;
    }
  }
  let expression = result.join("");
  return expression;
}

/////////////////////////////////////////////////////////////
////// （F）判断某个 token 是否是运算符
/////////////////////////////////////////////////////////////
function isOperator(token) {
  return (token === '+' || token === '-' || token === '*' || token === '/' || token === '(' || token === ')');
}

//////////////////////////////-----------------CleanOperator---------------End----------//////////////////////////////////




/////////////////////////////---------------连续乘除判断----------------Start-----------///////////////////////////////////

// 示例：两个函数（子函数），后续可以根据实际业务扩展内部逻辑
function Func1(segment) {
  console.log("Func1 called on pure-division segment:" + segment);
}

function Func2(segment) {
  console.log("Func2 called on mixed mul/div segment:" + segment);
}

function Func3(segment) {
  console.log("Func3 called on mixed mul/div segment:" + segment);
}

// 辅助函数：判断 token 是否为运算符
function isOperator(token) {
  return token === '+' || token === '-' || token === '*' || token === '/';
}

// 辅助函数：判断 token 是否为乘法或除法运算符
function isMulDivOperator(token) {
  return token === '*' || token === '/';
}

// 词法拆分：将公式拆分成操作数（例如 A、B、C/D 等）和运算符
// 此处假设操作数为由字母、数字组成的字符串；运算符为 + - * /
// 例如："A/D+B/C*E/F/G+H/I/J/K"  将拆分成 ["A", "/", "D", "+", "B", "/", "C", "*", "E", "/", "F", "/", "G", "+", "H", "/", "I", "/", "J", "/", "K"]
function tokenizeFormula(formula) {
  // 匹配字母数字序列 或 单个运算符
  let tokenPattern = /([+\-*/()])/;
  return formula.split(tokenPattern).map(t => t.trim()).filter(Boolean);
}

/**
 * 子函数1：扫描公式中的连续乘除链，
 * 如果该链中所有运算符均为除法（"/"），则调用 Func1。
 */
function processPureDivisionSegments(formula,FormulaTokens) {
  console.log("PureDivision 1");
  console.log("formula is " + formula);
  const tokens = tokenizeFormula(formula);
  console.log("PureDivision tokens is " + tokens);
  let found = false;  // 标记是否找到纯除法段
  let SubFormula = null; //返回最后处理过的公式的一部分
  let i = 0;
  while (i < tokens.length) {
    // 如果当前 token 是操作数，则可能开启一个乘除链
    // 每次只找公式中有可能的一部分连除，例如 1+A/B/C+D/E/F，循环替代
    // 因为传递进来的formula是一个括号里的公式，因此这里的while循环会全部替换掉连除相关公式
    // 最后会变成1+A/(B*C)+D/(E*F)
    if (!isOperator(tokens[i])) {
      // 以当前操作数开始，收集后续连续的乘除运算部分
      let segmentTokens = [tokens[i]];
      // let j = i + 1;
      // 只要后续的 token 是 "*" 或 "/"，并且后面跟着一个操作数，就加入链中
      console.log("PureDivision 2");

      // while (j < tokens.length && isMulDivOperator(tokens[j]) && (j + 1) < tokens.length && !isOperator(tokens[j + 1])) {
      //   segmentTokens.push(tokens[j]);       // 运算符
      //   segmentTokens.push(tokens[j + 1]);     // 下一个操作数
      //   j += 2;
      // }
      //下面一定要用isMulDivOperator来判断是否是乘号和除号混合，因为有这种情况A*B/C/D/E*F，
      // 这样的话如果没有下一步，就会直接匹配到B/C/D/E，
      // 而A*B/C/D/E*F应该用processMixedMulDivSegments进行混合判断，先调整成为A*F*B/C/D/E，
      // 然后再进一步改调用本函数变A*F*B/(C*D*E)
      for (j = i + 1; j < tokens.length;j+=2){
        if (isMulDivOperator(tokens[j]) && (j + 1) < tokens.length && !isOperator(tokens[j + 1])){
            segmentTokens.push(tokens[j]);       // 运算符
            segmentTokens.push(tokens[j + 1]);     // 下一个操作数
        }else{

          break; // 如果后面的符号不是乘除，则直接退出此次循环
        }
      }

      console.log("PureDivision 3");
      // 如果至少有一个运算符（即 segmentTokens 长度大于 4），则判断是否仅为除法
      //下面的1，3，5 位置判断是否是除法，可以把上一步的乘法和除法混合给排除掉，得到纯粹的除法连乘
      if (segmentTokens.length > 4) {
        let allDivision = true;
        // 运算符一般位于索引 1, 3, 5, …（假设格式为：operand op operand op operand …）
        for (let k = 1; k < segmentTokens.length; k += 2) {
          if (segmentTokens[k] !== '/') {
            allDivision = false;
            break;
          }
        }
        console.log("PureDivision 4");

        if (allDivision) {
          // 用 join 拼接成字符串，调用 Func1
          const segStr = segmentTokens.join('');
          console.log("segStr is " + segStr );
          //返回处理过的公式的一部分，如果没有SumY处理则原样返回, 1+A/B/C/D 变成 1+A/(B*C*D)的A/(B*C*D)部分
          SubFormula = transformDivisionChain(segStr, FormulaTokens);  // SubFormula = A/(B*C*D)
          console.log("SubFormula 1 is " + SubFormula);

          //这里需要把传进来的formula的相应部分替换掉，不然后面不能匹配。1+A/B/C/D 变成 1+A/(B*C*D)
          formula = formula.replace(segStr, SubFormula);
          console.log("formula 7 is " + formula);

          //这里返回的SubFormula公式是类似于A/(B*C*D)，需要提取括号内的部分进一步处理乘法排序
          //这里应该不需要while循环，应该只可能有一个括号
          let innerMostParenthesesRegex = /\(([^()]*)\)/g; // match[0] 包含括号，match[1] 不包含括号，仅括号内内容
          let match;
          //如果有括号，说明连除已经变成了括号带乘法的形式，则进行下一步，在括号内对乘法SumY排序
          if((match = innerMostParenthesesRegex.exec(SubFormula)) !== null){   
            //match[0] 包含括号，match[1] 不包含括号，仅括号内内容
            //经过了上一步，无论什么公式结果，都根据SumY调整连乘的顺序
              console.log("match[1] is " + match[1]);

              AdjMultipleFormula = transformMulDivChain(match[1], FormulaTokens);
              console.log("AdjMultipleFormula is " + AdjMultipleFormula);

              formula = formula.replace(match[1], AdjMultipleFormula); //1+A/B/C/D 变成 1+A/(B*C*D)
              console.log("SubFormula 2 is " + formula);

          }


          found = true;
          // 如果只要求选择执行其中之一，可以选择在这里提前返回 true
          // return true;
        }
      }
      // 移动到下一个 token（跳过本次连续链已处理的部分）
      i = j;
      // console.log("PureDivision 5");
    } else {
      // 当前 token 为运算符，则直接跳过
      i++;
      // console.log("PureDivision 6");
    }
  }
  console.log("PureDivision 7");
  return {found, formula};
}

/**
 * 子函数2：扫描公式中的连续乘除链，
 * 如果该链中同时存在乘法("*")和除法("/")（也就是混合出现），则调用 Func2。
 */
function processMixedMulDivSegments(formula,FormulaTokens) {
  const tokens = tokenizeFormula(formula);
  console.log("tokens is " + tokens);
  let found = false;
  let i = 0;
  let SubFormula;
  while (i < tokens.length) {
    if (!isOperator(tokens[i])) {
      let segmentTokens = [tokens[i]];
      // let j = i + 1;
      let hasMul = false;
      let hasDiv = false;

      //可能有这样的传递进来的formula,N3*R3/J3/O3/P3+R3*L3*J3, 因此需要for 循环，分成几段获取
      // while (j < tokens.length && isMulDivOperator(tokens[j]) && (j + 1) < tokens.length && !isOperator(tokens[j + 1])) {
      //   // 检查运算符类型
      //   if (tokens[j] === '*') hasMul = true;
      //   if (tokens[j] === '/') hasDiv = true;
      //   segmentTokens.push(tokens[j]);
      //   segmentTokens.push(tokens[j + 1]);
      //   j += 2;
      // }

      for(j = i + 1; j<tokens.length; j += 2 ){
        if(isMulDivOperator(tokens[j]) && (j + 1) < tokens.length && !isOperator(tokens[j + 1])){
          if (tokens[j] === '*') hasMul = true;
          if (tokens[j] === '/') hasDiv = true;
          segmentTokens.push(tokens[j]);
          segmentTokens.push(tokens[j + 1]);
        }else{
          break;
        }

      }
      // 如果当前链中至少有一个运算符，并且同时包含乘法和除法，则调用 Func2
      // 如果只有除法的连除，应该在处理连除的函数中处理完毕了，这一步在之后，应该要处理含有乘法的公式
      // if (segmentTokens.length > 4 && hasMul && hasDiv) {
      console.log("MixedMulDiv 1");
      if (segmentTokens.length > 4 && hasMul) {
        const segStr = segmentTokens.join('');
        console.log("segStr is " + segStr );
        // Func2(segStr);
        //这里找到的公式类似于A*B/C/D*E，不包括括号，因此可以不用while循环去找括号

        SubFormula = transformMulDivChain(segStr, FormulaTokens); //返回原始formula的一部分，经过乘法排序A*E*B/C/D的部分
        console.log("SubFormula 2 is " + SubFormula);
        //先把SubFormula在传入的formua中替换掉
        formula = formula.replace(segStr,SubFormula); //将传入的公式，例如1+A*B/C/D*E，替换掉A*B/C/D*E这部分变成1+A*E*B/C/D
        console.log("formula 11 is " + formula);

        //接下来要处理A*E*B/C/D 的 /C/D部分
        let result = processPureDivisionSegmentsNew(SubFormula, FormulaTokens);
        formula = formula.replace(SubFormula,result.formula); //1+A*E*B/C/D 变成1+A*E*B/(C*D)
        // formula = `(${formula})`; // 这里要加括号，不然作为分母的时候就会有问题
        console.log("formula 12 is " + formula);

        found = true;
        // 如果只要求选择执行其中之一，可以选择在这里提前返回 true
        // return true;
      }


    }
    i++
  }
  console.log("MixedMulDiv 2");
  return {found,formula};
}




/**
* 新的子函数：识别公式中纯除法的连续部分（纯由 "/" 运算符构成）
* 例如，对于公式 "D*H/I/J/K" 会识别出纯除法部分 "H/I/J/K"
* 并调用 Func1(segStr) 进行处理。
*
* @param {string} formula - 例如 "D*H/I/J/K"
*/
function processPureDivisionSegmentsNew(formula, FormulaTokens) {
  const tokens = tokenizeFormula(formula);
  let found = false;  // 标记是否找到纯除法段
  let SubFormula = null; //返回最后处理过的公式的一部分
  let i = 0;
  console.log(" processPureDivision 1");
  while (i < tokens.length) {
    // 如果当前 token 是操作数，则可能开启一个除法链
    if (!isOperator(tokens[i])) {
      let j = i + 1;
      // 从当前操作数开始，收集纯除法链
      let segmentTokens = [tokens[i]];

      // 判断后续是否以除法 "/" 开始, 因为经过processMixedMulDivSegments处理，所有的乘号都应该放在了前面，后面的应该只有连除的可能
      for(j = i+1; j < tokens.length; j += 2){
          if (tokens[j] === '/' && (j + 1) < tokens.length && !isOperator(tokens[j + 1])) {
              // 加入除法符号及其后面的操作数
              segmentTokens.push(tokens[j]);       // "/"
              segmentTokens.push(tokens[j + 1]);     // 操作数
          }else{
              break; //碰到不是除法的符号就停止循环
          }
       }
            // 如果纯除法链至少包含一个除法运算（即至少三个 token，如 "H/I/J" 则长度为 5，
            // 这里要求长度 >= 3 以确保有连续的除法）
      if (segmentTokens.length >= 5) {
          //这里应该直接调用纯粹连续除法
          const segStr = segmentTokens.join('');
          console.log("segStr 3 is " + segStr);

          //能进到这一步，就应该肯定有连除，因此不需要判断result.found
          let result = processPureDivisionSegments(segStr,FormulaTokens); 
          console.log("result.SubFormula is " + result.formula);

          formula = formula.replace(segStr,result.formula);
          console.log("formula 7 is " + formula);



        // const segStr = segmentTokens.join('');
        // console.log("Pure division segment found: " + segStr);
        // // Func3(segStr);
        // SubFormula = transformDivisionChain(segStr, FormulaTokens); //返回处理过的公式的一部分，如果没有SumY处理则原样返回
        // console.log("SubFormula 1 is " + SubFormula);

        // //这里返回的SubFormula公式是类似于N3/(R3*O3*P3*J3)，需要提取括号内的部分进一步处理
        // //这里应该不需要while循环，应该只可能有一个括号
        // let innerMostParenthesesRegex = /\(([^()]*)\)/g; // match[0] 包含括号，match[1] 不包含括号，仅括号内内容
        // let match;
        // //如果有括号，说明连除已经变成了括号带乘法的形式，则进行下一步，在括号内对乘法SumY排序
        // if ((match = innerMostParenthesesRegex.exec(SubFormula)) !== null) {
        //   //match[0] 包含括号，match[1] 不包含括号，仅括号内内容
        //   //经过了上一步，无论什么公式结果，都根据SumY调整连乘的顺序
        //   console.log("match[0] is " + match[1]);

        //   AdjMultipleFormula = transformMulDivChain(match[1], FormulaTokens);
        //   console.log("AdjMultipleFormula is " + AdjMultipleFormula);

        //   SubFormula = SubFormula.replace(match[1], AdjMultipleFormula);
        //   console.log("SubFormula 2 is " + SubFormula);
        // }

        found = true;
        // 如果只要求选择执行其中之一，可以选择在这里提前返回 true
        // return true;
      }
            // 跳过已处理的这段链
            // i = j;
            // continue;

    }
    i++;
  }
  console.log(" processPureDivision 2");
  return { found, formula };
}

/////////////////////////////---------------连续乘除判断----------------End-------------///////////////////////////////////

////////////////////////////----------------连续除法加括号变乘法-----------Start--------///////////////////////////////////
/**
 * 根据公式以及 FormulaTokens 中的变量对应关系（SumType 只有 SumY 和 SumN），
 * 对连除公式进行转换：
 * 如果第一个除号后面的部分中任一操作数对应的 SumType 为 SumY，
 * 则将右侧部分用括号括起来，并将其中的除法改为乘法。
 *
 * @param {string} formula - 原始公式，例如 "A/B/C"
 * @param {Array} formulaTokens - 对象数组，每个对象形如：
 *   {
 *     Token: "A",          // 变量名
 *     TokenName: "Token A",
 *     SumType: "SumY",       // 或 "SumN"
 *     ReplaceIndex: null,
 *     isCell: true,
 *     isOperator: false,
 *     isNumber: false,
 *     isReplace: false,
 *     isOthers: false
 *   }
 * @returns {string} - 转换后的公式
 */
function transformDivisionChain(formula, formulaTokens) {
  // 按除号拆分公式，拆分后不包含除号
  let tokens = formula.split('/').map(t => t.trim());
  
  // 如果没有除法或只有一个操作数，直接返回原公式
  if (tokens.length < 2) {
    return formula;
  }
  
  // 取出右侧所有操作数（即第一个除号之后的部分）
  let rightTokens = tokens.slice(1);
  
  // 检查右侧是否存在 SumType 为 "SumY" 的变量
  let containsSumY = rightTokens.some(token => {
    // 在 FormulaTokens 数组中查找 token 对应的对象
    let obj = formulaTokens.find(item => item.Token === token);
    return obj && obj.SumType === "SumY";
  });
  
  // 如果右侧存在 SumY，则对右侧部分进行转换
  if (containsSumY) {
    // 将右侧部分的各操作数用乘法连接，并用括号括起来
    return tokens[0] + "/(" + rightTokens.join('*') + ")";
  } else {
    // 否则返回原公式
    return formula;
  }
}

////////////////////////////----------------连续除法加括号变乘法-----------Start--------///////////////////////////////////



////////////////////////////----------------混合乘除移动SumY------------Start----------///////////////////////////////////
/**
 * 转换乘除混合公式：
 * 1. 解析公式，将第一个操作数放入分子，
 *    后续遇到 "*" 的操作数归入分子，遇到 "/" 的操作数归入分母。
 * 2. 对分子中的因子，如果该因子是乘法产生的（即非第一个操作数）
 *    且在 FormulaTokens 中对应的 SumType 为 "SumY"，则视为“可移动”，
 *    将所有“可移动”的因子移到分子最前面，其余因子保持原来的相对顺序。
 * 3. 按 “分子因子乘积” / “分母因子乘积” 重构公式（如果分母为空，则只输出分子）。
 *
 * @param {string} formula - 乘除混合公式，例如 "A/B/C*D/E*F/G"
 * @param {Array} formulaTokens - 对象数组，每个对象至少包含：
 *       { Token: "变量名", SumType: "SumY" 或 "SumN", ... }
 * @returns {string} - 转换后的公式字符串
 */
function transformMulDivChain(formula, formulaTokens) {
  // 利用正则将公式拆分为操作数和运算符（保留 "*" 和 "/"）
  // 例如 "A/B/C*D/E*F/G" 拆分结果为：
  // ["A", "/", "B", "/", "C", "*", "D", "/", "E", "*", "F", "/", "G"]
  let tokens = formula.split(/([*/])/).map(t => t.trim()).filter(Boolean);
  if (tokens.length === 0) return formula;
  
  // 分别存放分子和分母因子，每个因子保存的信息包括：
  // { operand, index, movable }
  // 其中 index 表示该因子在原分子中的顺序，用于在排序时保持相对顺序。
  let numerator = [];
  let denominator = [];
  
  // 第一项始终属于分子，且视为不可移动（即无乘法运算符产生）
  numerator.push({
    operand: tokens[0],
    index: 0,
    movable: false  // 第一项不管其 SumType 如何，保持原位置
  });
  
  // 记录分子中后续因子的序号（用于排序时记录原始顺序）
  let numIndex = 1;
  let denomIndex = 0;
  
  // 从第二个 token 开始，每两个 token构成一组：[运算符, 操作数]
  for (let i = 1; i < tokens.length; i += 2) {
    let op = tokens[i];         // "*" 或 "/"
    let operand = tokens[i + 1];
    
    if (op === "*") {
      // 乘法产生的因子归入分子
      // 判断该因子是否可移动：仅当对应的 SumType 为 "SumY" 时视为可移动
      let tokenObj = formulaTokens.find(item => item.Token === operand);
      let sumType = tokenObj ? tokenObj.SumType : "SumN";
      let movable = (sumType === "SumY");
      numerator.push({
        operand: operand,
        index: numIndex,
        movable: movable
      });
      numIndex++;
    } else if (op === "/") {
      // 除法产生的因子归入分母（不可移动）
      denominator.push({
        operand: operand,
        index: denomIndex
      });
      denomIndex++;
    }
  }
  
  // 重新排序分子：将所有“可移动”的因子（乘法产生且 SumType 为 SumY）移到最前面，
  // 保持各组内原始顺序不变。
  // 这里对分子因子设置排序键：若 movable 为 true，则 key = 0，否则 key = 1。
  let sortedNumerator = numerator.slice().sort((a, b) => {
    let keyA = a.movable ? 0 : 1;
    let keyB = b.movable ? 0 : 1;
    if (keyA !== keyB) {
      return keyA - keyB;
    } else {
      return a.index - b.index; // 保持原始顺序
    }
  });
  
  // 重构分子字符串：将各因子用 "*" 连接
  let numeratorStr = sortedNumerator.map(f => f.operand).join('*');
  // 重构分母字符串：保持原顺序，用 "/" 连接各因子
  let denominatorStr = denominator.map(f => f.operand).join('/');
  
  // 如果有分母，则构造形如 "numeratorStr/denominatorStr" 的公式；
  // 否则仅输出 numeratorStr。
  let newFormula = denominatorStr ? (numeratorStr + "/" + denominatorStr) : numeratorStr;
  return newFormula;
}

////////////////////////////----------------混合乘除移动SumY------------End----------///////////////////////////////////


///////////////////////////-----------------计算SumY-------------------Start--------///////////////////////////////////

/**
 * 从单元格 D3 读取公式（例如 "=A3+B3*C3"），
 * 从 A1, B1, C1 读取每个引用对应的类型（"SumY" / "SumN"），
 * 然后调用 determineExpressionType() 分析四则运算类型。
 *
 * 注意：此示例较简易，未处理复杂公式的各种情况。
 * 有可能出现A/(B*C*D) 这样内部带括号的输入
 */
function checkFormulaResultType(formulaString,FormulaTokens) {
  // try {
  //   await Excel.run(async (context) => {
      // // 假设在当前工作表 Sheet1 中
      // let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");

      // // 1. 读取公式所在的单元格（例如 D3）
      // let formulaRange = sheet.getRange(ResultAddress);
      // formulaRange.load("formulas"); // 加载公式
      // await context.sync();

      // // formulas 是一个二维数组，这里只取 [0][0] 对应这个单元格的公式字符串
      // let formulaString = formulaRange.formulas[0][0]; 
      // 例如 formulaString == "=A3+B3*C3"
      console.log("Read formula:" + formulaString);

      // 找最内侧的单元格match[0] 包含括号，match[1] 不包含括号，仅括号内内容
      let innerMostParenthesesRegex = /\(([^()]*)\)/g; 

      // 下面在最内层的单元格中找到公式，先进行排序，然后计算是否是SumN
      let i = 1;
      let match;
      let tokenPattern = /([+\-*/()])/g; 
      let resultType = null;
        //`(${formulaString})` 这里再公式的最外边临时加上一对括号，让最外层至少也循环一次，避免没有括号不循环的情况
      formulaString = "(" + formulaString + ")";
      console.log("formulaString before is " + formulaString);
      while ((match = innerMostParenthesesRegex.exec(formulaString)) !== null) {
        console.log("match[0] start is " + match[0]); // match[1] 不包括括号
        console.log("match[1] start is " + match[1]); // match[1] 不包括括号
      
      //   // 1. 分离出运算符和操作数。这里做一个非常粗糙的 split，然后再拼接运算符
      //   //    也可以用更复杂的正则或解析器。这里只是演示思路。
      //   //    以 + - * / 作为分隔符，并保留分隔符（用于后续组装）。
      //   //    例如 "A3+B3*C3" -> ["A3", "+", "B3", "*", "C3"]。
        
        // split 之后会把分隔符也拆出来
        let tokens = match[1].split(tokenPattern).map(t => t.trim()).filter(Boolean);
        // tokens => ["A3", "+", "B3", "*", "C3"]
        console.log("tokens C is ");
        console.log(tokens);
        // 2. 对每个单元格引用（例如 "A3", "B3"）进行类型替换
        for (let i = 0; i < tokens.length; i++) {
          let t = tokens[i];
          // 如果是运算符 (+ - * /)，就直接跳过
          console.log("t is " + t);
          if (t === "+" || t === "-" || t === "*" || t === "/" || t === "(" || t === ")" || t === "SumY" || t === "SumN") {
              continue;
            };

          let obj = FormulaTokens.find(item => item.Token === t);
          tokens[i] =  obj.SumType;  //获取对象中的SumY或SumN

        }
      
        // 3. 组装成一个最终字符串（用空格隔开，便于后续 determineExpressionType() 解析）
        //    例如 ["SumY", "+", "SumN", "*", "SumY"] -> "SumY + SumN * SumY"
        let expression = tokens.join(" "); //这里是用空格隔开了各个变量
        console.log("Parsed expression:" + expression);

        // 4. 调用我们前面写好的 determineExpressionType(expression) 函数，得到最终类型
        resultType = determineExpressionType(expression);
        console.log("括号内的 Type :" + resultType);
        console.log("match[0] is " + match[0]);
        console.log("formulaString is " + formulaString);

        // // 先去掉空格，保证匹配
        // let cleanedMatch = match[0].replace(/\s+/g, "");

        // // 使用正则替换
        // formulaString = formulaString.replace("(O3*P3*R3*J3)", resultType);
        formulaString = formulaString.replace(match[0],resultType); //将带有括号的match[0]用SumY或SumN替代掉  
        console.log("formulaString AAA is " + formulaString);
        innerMostParenthesesRegex.lastIndex = 0;
      }

      console.log("处理完括号后 formulaString is " + formulaString);

      // //剩下的部分没有括号，重新处理一遍
      // let tokens = formulaString.split(tokenPattern).map(t => t.trim()).filter(Boolean);
      // for (let i = 0; i < tokens.length; i++) {
      //   let t = tokens[i];
      //   // 如果是运算符 (+ - * /)，就直接跳过，并且跳过上一步带括号已经计算出来的SumY和SumN
      //   console.log("t is " + t);
      //   if (t === "+" || t === "-" || t === "*" || t === "/" || t === "(" || t === ")" || t === "SumY" || t === "SumN") {
      //       continue;
      //     };

      //   let obj = FormulaTokens.find(item => item.Token === t);
      //   tokens[i] =  obj.SumType;  //获取对象中的SumY或SumN
      // }

      // let expression = tokens.join(" "); //这里是用空格隔开了各个变量
      // console.log("Parsed 2 expression:" + expression);
      // resultType = determineExpressionType(expression);
      // console.log("final Typpe is " + resultType);

      return resultType;

}

/**
 * =========================
 * 以下是前面已有的函数
 * =========================
 */

/**
 * 将表达式字符串（"SumY + SumN * SumY"）解析计算最终类型
 * @param {string} expression
 * @returns {string} "SumY" or "SumN"
 */
function determineExpressionType(expression) {
  try {
    let tokens = expression.split(/([+\-*/()])/g).map(t => t.trim()).filter(Boolean);
    console.log("tokens D is " + tokens);
    // 先处理 * / 优先级
    handleMultiplyDivide(tokens);
    console.log("tokens D 2 is " + tokens);
    // 再处理 + -
    handleAddSubtract(tokens);
    console.log("tokens D 3 is " + tokens);
    // 最终应只剩一个
    if (tokens.length === 1) {
      return tokens[0];
    } else {
      return "表达式有误，无法判定";
    }
  } catch (err) {
    console.error(err);
    return "Error";
  }
}

/** 处理乘除 */
function handleMultiplyDivide(tokens) {
  let i = 0;
  while (i < tokens.length) {
    let token = tokens[i];
    if (token === "*" || token === "/") {
      let leftType = tokens[i - 1];
      let rightType = tokens[i + 1];
      let newType = combineTypes(leftType, rightType, token);
      tokens.splice(i - 1, 3, newType);
      i = i - 1;
    } else {
      i++;
    }
  }
}

/** 处理加减 */
function handleAddSubtract(tokens) {
  let i = 0;
  while (i < tokens.length) {
    let token = tokens[i];
    if (token === "+" || token === "-") {
      let leftType = tokens[i - 1];
      let rightType = tokens[i + 1];
      let newType = combineTypes(leftType, rightType, token);
      tokens.splice(i - 1, 3, newType);
      i = i - 1;
    } else {
      i++;
    }
  }
}

/**
 * 根据自定义规则，将两个类型与运算符合并得出结果
//  */
// function combineTypes(type1, type2, operator) {
//   // 加减规则不变
//   if (operator === "+" || operator === "-") {
//     // 只有两个都是 SumN 才返回 SumN，否则 SumY
//     if (type1 === "SumY" && type2 === "SumY") {
//       return "SumY";
//     } else if(type1 === "SumY" && type2 === "SumN") {
//       return "SumY";
//     } else if (type1 === "SumN" && type2 === "SumY") {
//       return "SumY";
//     } else {
//       return "SumN";
//     }
//   }

//   // 乘法规则
//   if (operator === "*") {
//     // - SumY * SumY → SumY
//     // - SumY * SumN → SumY
//     // - SumN * SumY → SumY
//     // - SumN * SumN → SumN
//     if (type1 === "SumY" && type2 === "SumY") {
//       return "SumY";
//     } else if (type1 === "SumY" && type2 === "SumN") {
//       return "SumY";
//     } else if (type1 === "SumN" && type2 === "SumY") {
//       return "SumY";
//     } else {
//       return "SumN";
//     }
//   }

//   // 除法规则
//   if (operator === "/") {
//     // - SumY / SumY → SumN
//     // - SumY / SumN → SumY
//     // - SumN / SumY → SumN
//     // - SumN / SumN → SumN
//     if (type1 === "SumY" && type2 === "SumY") {
//       return "SumN";
//     } else if (type1 === "SumY" && type2 === "SumN") {
//       return "SumY";
//     } else if (type1 === "SumN" && type2 === "SumY") {
//       return "SumN";
//     } else {
//       return "SumN";
//     }
//   }

//   // 其余情况，默认 SumN
//   return "SumN";
// }


///////////////////////////-----------------计算SumY-------------------End----------///////////////////////////////////

//////////////////////////------------清除0，包含乘除法的公式-----------Start---------///////////////////

// ─────────────────────────────
// 1. 词法分析：将输入字符串拆分为标记（token）数组
// ─────────────────────────────
function tokenize(input) {
  const regex = /\s*([A-Za-z]\w*|\d+(\.\d+)?|[+\-*/()]|\S)\s*/g;
  let tokens = [];
  let match;
  while ((match = regex.exec(input)) !== null) {
    let token = match[1];
    if (/^[A-Za-z]\w*$/.test(token)) {
      tokens.push({ type: 'Identifier', value: token });
    } else if (/^\d+(\.\d+)?$/.test(token)) {
      tokens.push({ type: 'Literal', value: parseFloat(token) });
    } else if (token === '+' || token === '-' || token === '*' || token === '/') {
      tokens.push({ type: 'Operator', value: token });
    } else if (token === '(' || token === ')') {
      tokens.push({ type: 'Paren', value: token });
    } else {
      throw new Error("Unknown token: " + token);
    }
  }
  return tokens;
}

// ─────────────────────────────
// 2. 解析器：构造 AST（抽象语法树）
// ─────────────────────────────

function parseExpression(tokens) {
  let node = parseTerm(tokens);
  while (peek(tokens) && peek(tokens).type === 'Operator' &&
         (peek(tokens).value === '+' || peek(tokens).value === '-')) {
    let op = consume(tokens).value;
    let right = parseTerm(tokens);
    node = { type: 'BinaryExpression', operator: op, left: node, right: right };
  }
  return node;
}

function parseTerm(tokens) {
  let node = parseFactor(tokens);
  while (peek(tokens) && peek(tokens).type === 'Operator' &&
         (peek(tokens).value === '*' || peek(tokens).value === '/')) {
    let op = consume(tokens).value;
    let right = parseFactor(tokens);
    node = { type: 'BinaryExpression', operator: op, left: node, right: right };
  }
  return node;
}

function parseFactor(tokens) {
  let token = peek(tokens);
  if (token && token.type === 'Operator' && (token.value === '+' || token.value === '-')) {
    let op = consume(tokens).value;
    let argument = parseFactor(tokens);
    return { type: 'UnaryExpression', operator: op, argument };
  } else if (token && token.type === 'Literal') {
    return consume(tokens);
  } else if (token && token.type === 'Identifier') {
    return consume(tokens);
  } else if (token && token.type === 'Paren' && token.value === '(') {
    consume(tokens); // consume "("
    let node = parseExpression(tokens);
    if (!peek(tokens) || peek(tokens).type !== 'Paren' || peek(tokens).value !== ')') {
      throw new Error("Expected closing parenthesis");
    }
    consume(tokens); // consume ")"
    return node;
  } else {
    throw new Error("Unexpected token: " + JSON.stringify(token));
  }
}

function peek(tokens) {
  return tokens[0];
}

function consume(tokens) {
  return tokens.shift();
}

// ─────────────────────────────
// 3. 加减法表达式的平铺与简化（相关零的处理）
// ─────────────────────────────

function flattenAdditive(node) {
  let result = [];
  function helper(n, currentSign) {
    if (n.type === 'BinaryExpression' &&
        (n.operator === '+' || n.operator === '-')) {
      helper(n.left, currentSign);
      let newSign = n.operator === '+' ? currentSign : (currentSign === '+' ? '-' : '+');
      helper(n.right, newSign);
    } else if (n.type === 'UnaryExpression' && n.operator === '-') {
      helper(n.argument, currentSign === '+' ? '-' : '+');
    } else {
      result.push({ sign: currentSign, node: n });
    }
  }
  helper(node, '+');
  return result;
}

function simplifyAST(node) {
  if (!node) return node;
  
  if (node.type === 'BinaryExpression' &&
      (node.operator === '+' || node.operator === '-')) {
    let terms = flattenAdditive(node);
    for (let term of terms) {
      term.node = simplifyAST(term.node);
    }
    let allLiterals = terms.every(term => term.node.type === 'Literal');
    
    if (terms.length === 2) {
      if (terms[0].node.type === 'Literal' && terms[0].node.value === 0) {
        // "0 - X" 保持不变
      } else if (terms[1].node.type === 'Literal' && terms[1].node.value === 0) {
        terms = [terms[0]];
      }
    }
    else if (terms.length > 2 &&
             terms[0].node.type === 'Literal' && terms[0].node.value === 0) {
      if (allLiterals) {
        if (terms[terms.length - 1].sign === '+') {
          if (terms.length >= 3) {
            let temp = terms[1];
            terms[1] = terms[2];
            terms[2] = temp;
          }
          terms.shift();
        } else {
          let combined = 0;
          for (let i = 1; i < terms.length; i++) {
            combined += (terms[i].sign === '+' ? terms[i].node.value : -terms[i].node.value);
          }
          return {
            type: 'BinaryExpression',
            operator: '-',
            left: terms[0].node,
            right: { type: 'Literal', value: Math.abs(combined) }
          };
        }
      } else {
        if (terms.length >= 3 && terms[1].sign === '-' && terms[2].sign === '+') {
          let temp = terms[1];
          terms[1] = terms[2];
          terms[2] = temp;
          terms.shift();
        }
        // 否则保持原顺序，保留首项 0
      }
    }
    
    if (terms.length === 1) {
      if (terms[0].sign === '-') {
        return { type: 'UnaryExpression', operator: '-', argument: terms[0].node };
      } else {
        return terms[0].node;
      }
    }
    let newNode = terms[0].node;
    for (let i = 1; i < terms.length; i++) {
      newNode = {
        type: 'BinaryExpression',
        operator: terms[i].sign,
        left: newNode,
        right: terms[i].node
      };
    }
    return newNode;
  } else if (node.type === 'BinaryExpression') {
    node.left = simplifyAST(node.left);
    node.right = simplifyAST(node.right);
    return node;
  } else if (node.type === 'UnaryExpression') {
    node.argument = simplifyAST(node.argument);
    return node;
  } else {
    return node;
  }
}

// ─────────────────────────────
// 4. 将 AST 转换回字符串
// ─────────────────────────────
function astToString(node) {
  if (!node) return "";
  switch (node.type) {
    case 'Literal':
      return String(node.value);
    case 'Identifier':
      return node.value;
    case 'UnaryExpression': {
      let argStr = astToString(node.argument);
      if (node.argument.type === 'BinaryExpression') {
        argStr = '(' + argStr + ')';
      }
      return node.operator + argStr;
    }
    case 'BinaryExpression': {
      function precedence(op) {
        if (op === '+' || op === '-') return 1;
        if (op === '*' || op === '/') return 2;
        return 0;
      }
      let leftStr = astToString(node.left);
      let rightStr = astToString(node.right);
      if (node.left.type === 'BinaryExpression' &&
          precedence(node.left.operator) < precedence(node.operator)) {
        leftStr = '(' + leftStr + ')';
      }
      // 【修改部分】：如果当前节点的运算符为 '/'，且右子树为二元表达式，则始终加括号
      if (node.right.type === 'BinaryExpression' &&
          (node.operator === '/' || precedence(node.right.operator) < precedence(node.operator))) {
        rightStr = '(' + rightStr + ')';
      }
      return leftStr + ' ' + node.operator + ' ' + rightStr;
    }
    default:
      return "";
  }
}

// ─────────────────────────────
// 5. 综合函数：解析、简化并返回新的表达式字符串
// ─────────────────────────────
function simplifyExpression(formula) {
  const tokens = tokenize(formula);
  const ast = parseExpression(tokens);
  const simplifiedAst = simplifyAST(ast);
  // 如果需要移除所有空格，可以在这里加 .replace(/\s+/g, "")
  return astToString(simplifiedAst).replace(/\s+/g, ""); // 去掉所有空格
}

// ─────────────────────────────
// 测试示例
// ─────────────────────────────
// console.log(simplifyExpression("Q3 * (0-2 + 1)"));            // 预期: "Q3 * (1 - 2)"
// console.log(simplifyExpression("Q3 * (0-2 - 1)"));            // 预期: "Q3 * (0 - 3)"
// console.log(simplifyExpression("Q3 * (0-2 + 1*5 + 1)"));        // 预期: "Q3 * (1*5 - 2 + 1)"
// console.log(simplifyExpression("L3*(0-P3*1+O3)"));             // 预期: "L3 * (O3 - P3 * 1)"
// console.log(simplifyExpression("L3*(0-P3*1-O3)"));             // 预期: "L3 * (0 - P3 * 1 - O3)"
// console.log(simplifyExpression("N3/(P3*O3*R3*J3)"));           // 预期: "N3/(P3*O3*R3*J3)"
// console.log(simplifyExpression("G-A-B+C-D"));                 // 保持 "G-A-B+C-D"



//////////////////////////------------清除0，包含乘除法的公式-----------Start---------///////////////////


/////////////////////////--------------处理和修正公式中错误的情况--------Start------------////////////////////

/**
 * 主入口函数
 * 处理公式，生成 AST 并检测是否存在 checkType1～checkType6 模式。
 * 当检测到 checkType1 时，会自动使用乘法分配律展开表达式，
 * 并记录展开后的公式。最终返回一个对象，其中包含：
 *   - finalType: 最终计算结果类型
 *   - finalFormula: 最终公式（经过展开后，如有）
 *   - detections: 检测到的模式信息数组（每个对象包括 type、formula，checkType1 还包括 expandedFormula）
 *   - expandedFormula: 如果存在 checkType1 展开，则返回展开后的公式字符串，否则为 null
 *
 * @param {string} expression 原始公式，例如 "A * (D + E)" 或 "(A + B * C * (D – E + D)) * C"
 * @param {Array} FormulaTokens 变量信息数组
 * @returns {Object} 结果对象
 */
function evaluateExpressionStepByStep(expression, FormulaTokens) {
  // 定义全局检测信息数组
  let detections = [];
  // 1. 分词生成初始 tokens，并传入全局 detections
  let tokens = tokenizeExpression(expression, FormulaTokens, detections);
  // 2. 递归处理括号及非括号部分，生成完整 AST
  tokens = evaluateTokensStepByStep(tokens, detections);
  // 3. 遍历最终 tokens 中每个 token 的 AST，收集检测信息
  tokens.forEach(token => {
    if (token.ast) {
      traverseAST(token.ast, null, false, detections);
    }
  });
  // 4. 将最终 tokens 的 AST 更新回 token 对象
  tokens.forEach(token => {
    if (token.ast) {
      token.origStr = token.ast.origStr;
      token.computedType = token.ast.sumType;
    }
  });
  // 5. 如果最终结果只有一个 token，则检查模式5
  if (tokens.length === 1) {
    checkType5Final(tokens[0], detections);
    console.log("最终计算结果: " + tokens[0].computedType + " 对应公式: " + tokens[0].origStr);
    console.log("检测信息：", detections);
    let expandedFormula = null;
    for (let det of detections) {
      if (det.type === "checkType1" && det.expandedFormula) {
        expandedFormula = det.expandedFormula;
        break;
      }
    }
    return {
      finalType: tokens[0].computedType,
      finalFormula: tokens[0].origStr,
      detections: detections,
      expandedFormula: expandedFormula
    };
  } else {
    console.log("表达式有误，无法计算");
    return { error: "表达式有误，无法计算" };
  }
}

/* ============ 分词及 AST 构造 ============ */
/**
 * 将公式字符串拆分成 token 数组。
 * 对于变量，从 FormulaTokens 中获取 SumType 与显示名称；
 * 对于运算符（包括括号）直接生成 token 对象。
 *
 * 每个 token 包含：
 *   - token: 原始符号
 *   - origStr: 显示字符串
 *   - computedType: 变量初始的 SumType（SumY 或 SumN）；运算符无此属性
 *   - isOperator: 是否为运算符
 *   - ast: 对应的 AST 节点（变量节点类型为 "variable"）
 *
 * @param {string} expression
 * @param {Array} FormulaTokens
 * @param {Array} detections 全局检测信息数组（中间阶段不使用，但统一传递）
 * @returns {Array} tokens 数组
 */
function tokenizeExpression(expression, FormulaTokens, detections) {
  let rawTokens = expression.split(/([+\-*/()])/g)
    .map(t => t.trim())
    .filter(Boolean);
  let tokens = [];
  rawTokens.forEach(tok => {
    if (["+", "-", "*", "/", "(", ")"].includes(tok)) {
      tokens.push({
        token: tok,
        origStr: tok,
        isOperator: true
      });
    } else {
      let found = FormulaTokens.find(item => item.Token === tok);
      let computed = found ? found.SumType : "SumN";
      let display = found ? (found.Token || tok) : tok;
      tokens.push({
        token: tok,
        origStr: display,
        computedType: computed,
        isOperator: false,
        ast: { type: "variable", token: tok, sumType: computed, origStr: display }
      });
    }
  });
  console.log("初始 tokens: " + tokens.map(t => t.origStr).join(" "));
  checkTypeAll(tokens, true, detections);
  return tokens;
}

/* ============ AST 构造与求值 ============ */
/**
 * 递归处理括号及非括号部分的求值，返回合并后的 tokens 数组。
 */
function evaluateTokensStepByStep(tokens, detections) {
  tokens = evaluateWithParentheses(tokens, detections);
  tokens = evaluateNoParentheses(tokens, detections);
  return tokens;
}

/**
 * 处理括号部分：
 * 找到最内层括号，将括号内 tokens 递归求值，
 * 构造新 token 时保留括号内原始 tokens（rawInnerTokens），生成 AST 节点类型 "parentheses"。
 */
function evaluateWithParentheses(tokens, detections) {
  while (tokens.some(t => t.isOperator && t.token === "(")) {
    let openIndex = -1;
    for (let i = 0; i < tokens.length; i++) {
      if (tokens[i].isOperator && tokens[i].token === "(") {
        openIndex = i;
      }
    }
    if (openIndex === -1) break;
    let closeIndex = tokens.findIndex((t, idx) => idx > openIndex && t.isOperator && t.token === ")");
    if (closeIndex === -1) break;
    let rawInnerTokens = tokens.slice(openIndex + 1, closeIndex);
    console.log("Evaluating parentheses expression: " + rawInnerTokens.map(t => t.origStr).join(" "));
    let evaluatedSubTokens = evaluateTokensStepByStep(rawInnerTokens, detections);
    if (evaluatedSubTokens.length !== 1) {
      console.error("括号内表达式计算错误");
      return tokens;
    }
    let subResult = evaluatedSubTokens[0];
    let newToken = {
      token: subResult.computedType,
      origStr: "(" + rawInnerTokens.map(t => t.origStr).join(" ") + ")",
      computedType: subResult.computedType,
      isOperator: false,
      isParentheses: true,
      rawInnerTokens: rawInnerTokens,
      ast: {
        type: "parentheses",
        inner: subResult.ast,
        origStr: "(" + rawInnerTokens.map(t => t.origStr).join(" ") + ")",
        sumType: subResult.computedType
      }
    };
    tokens.splice(openIndex, closeIndex - openIndex + 1, newToken);
    console.log("After evaluating parentheses: " + tokens.map(t => t.origStr).join(" "));
    checkTypeAll(tokens, true, detections);
  }
  return tokens;
}

/**
 * 处理非括号部分的表达式，先处理乘除，再处理加减。
 * 每次合并时构造新的 AST 节点（类型 "operator"），记录原始公式部分。
 */
function evaluateNoParentheses(tokens, detections) {
  // 先处理乘除
  while (tokens.length > 1 && tokens.some(t => t.isOperator && (t.token === "*" || t.token === "/"))) {
    checkTypeAll(tokens, true, detections);
    let processed = false;
    for (let i = 0; i < tokens.length; i++) {
      if (tokens[i].isOperator && (tokens[i].token === "*" || tokens[i].token === "/")) {
        let operator = tokens[i].token;
        if (i - 1 < 0 || i + 1 >= tokens.length) continue;
        let left = tokens[i - 1];
        let right = tokens[i + 1];
        let newComputedType = combineTypes(left.computedType, right.computedType, operator);
        let newOrigStr = left.origStr + " " + operator + " " + right.origStr;
        let newAST = {
          type: "operator",
          op: operator,
          left: left.ast,
          right: right.ast,
          origStr: newOrigStr,
          sumType: newComputedType
        };
        let newToken = {
          token: newComputedType,
          origStr: newOrigStr,
          computedType: newComputedType,
          isOperator: false,
          ast: newAST
        };
        tokens.splice(i - 1, 3, newToken);
        console.log("After evaluating: " + left.origStr + " " + operator + " " + right.origStr +
          " -> " + tokens.map(t => t.origStr).join(" "));
        processed = true;
        checkTypeAll(tokens, true, detections);
        break;
      }
    }
    if (!processed) break;
  }
  
  // 再处理加减
  while (tokens.length > 1 && tokens.some(t => t.isOperator && (t.token === "+" || t.token === "-"))) {
    checkTypeAll(tokens, true, detections);
    let processed = false;
    for (let i = 0; i < tokens.length; i++) {
      if (tokens[i].isOperator && (tokens[i].token === "+" || tokens[i].token === "-")) {
        let operator = tokens[i].token;
        if (i - 1 < 0 || i + 1 >= tokens.length) continue;
        let left = tokens[i - 1];
        let right = tokens[i + 1];
        let newComputedType = combineTypes(left.computedType, right.computedType, operator);
        let newOrigStr = left.origStr + " " + operator + " " + right.origStr;
        let newAST = {
          type: "operator",
          op: operator,
          left: left.ast,
          right: right.ast,
          origStr: newOrigStr,
          sumType: newComputedType
        };
        let newToken = {
          token: newComputedType,
          origStr: newOrigStr,
          computedType: newComputedType,
          isOperator: false,
          ast: newAST
        };
        tokens.splice(i - 1, 3, newToken);
        console.log("After evaluating: " + left.origStr + " " + operator + " " + right.origStr +
          " -> " + tokens.map(t => t.origStr).join(" "));
        processed = true;
        checkTypeAll(tokens, true, detections);
        break;
      }
    }
    if (!processed) break;
  }
  return tokens;
}

/* ================= 检测函数 ================= */
/**
 * 遍历 tokens 数组中每个 token 的 AST，
 * 对每个 AST 节点调用 traverseAST 并传入父运算符（初始为 null）。
 * 参数 skipType6 为 true 时跳过检测模式6（中间阶段）。
 * 如果传入 detections 数组，则将检测到的模式信息保存进去。
 */
function checkTypeAll(tokens, skipType6, detections) {
  tokens.forEach(token => {
    if (token.ast) {
      traverseAST(token.ast, null, skipType6, detections);
    }
  });
}

/**
 * 递归遍历 AST，检测以下模式：
 *
 * 模式1 (checkType1)：SumY * ( ...加法或减法... )
 *    当检测到时，将检测信息 {type:"checkType1", formula, expandedFormula}
 *    保存到 detections 数组，并利用乘法分配律展开表达式，
 *    替换当前节点。
 *
 * 模式2 (checkType2)：SumY / (SumN ± SumN)
 * 模式3 (checkType3)：SumN / SumN
 * 模式4 (checkType4)：SumN / SumY
 * 模式6 (checkType6)：SumN * SumN（仅当父运算符为空或为加/减时检测）
 *
 * @param {Object} node AST 节点
 * @param {string|null} parentOp 父节点的运算符（无则为 null）
 * @param {boolean} skipType6 若为 true，则跳过检测模式6（中间阶段）
 * @param {Array} detections 检测结果数组，保存检测信息对象 { type, formula, [expandedFormula] }
 */
function traverseAST(node, parentOp = null, skipType6 = false, detections = []) {
  if (!node) return;

  // 模式1检测：SumY * ( ...加法或减法... )
  if (node.type === "operator" && node.op === "*") {
    if (node.left && node.left.sumType === "SumY" &&
        node.right && node.right.type === "parentheses" &&
        node.right.inner && node.right.inner.type === "operator" &&
        (node.right.inner.op === "+" || node.right.inner.op === "-")) {
      // 扁平化括号内的加减，得到所有加数及符号
      let addends = flattenAddSub(node.right.inner);
      // 如果所有加数的 term 均为 SumN，则符合模式
      if (addends.every(item => item.term.sumType === "SumN")) {
        let originalFormula = node.origStr;
        console.log("检测到特定类型1公式部分: " + originalFormula);
        // 展开：利用乘法分配律展开 L * (X1 ± X2 ± ... ± Xn)
        let expanded = expandCheckType1(node, addends);
        console.log("展开后公式: " + expanded.origStr);
        detections.push({ type: "checkType1", formula: originalFormula, expandedFormula: expanded.origStr });
        // 替换当前节点为展开后的结果
        node.type = expanded.type;
        node.op = expanded.op;
        node.left = expanded.left;
        node.right = expanded.right;
        node.origStr = expanded.origStr;
        node.sumType = expanded.sumType;
      }
    }
    // 模式6检测：SumN * SumN
    if (!skipType6 && node.left && node.left.sumType === "SumN" &&
        node.right && node.right.sumType === "SumN") {
      if (parentOp === null || parentOp === "+" || parentOp === "-") {
        console.log("检测到特定类型6公式部分: " + node.origStr);
        detections.push({ type: "checkType6", formula: node.origStr });
      }
    }
  }

  // 模式2检测：SumY / (SumN ± SumN)
  if (node.type === "operator" && node.op === "/") {
    if (node.left && node.left.sumType === "SumY" &&
        node.right && node.right.type === "parentheses" &&
        node.right.inner && node.right.inner.type === "operator" &&
        (node.right.inner.op === "+" || node.right.inner.op === "-") &&
        node.right.inner.left && node.right.inner.left.sumType === "SumN" &&
        node.right.inner.right && node.right.inner.right.sumType === "SumN") {
      console.log("检测到特定类型2公式部分: " + node.origStr);
      detections.push({ type: "checkType2", formula: node.origStr, denominatorPart: node.right.origStr });
    }
    // 模式3检测：SumN / SumN
    if (node.left && node.left.sumType === "SumN" &&
        node.right && node.right.sumType === "SumN") {
      console.log("检测到特定类型3公式部分: " + node.origStr);
      detections.push({ type: "checkType3", formula: node.origStr });
    }
    // 模式4检测：SumN / SumY
    if (node.left && node.left.sumType === "SumN" &&
        node.right && node.right.sumType === "SumY") {
      console.log("检测到特定类型4公式部分: " + node.origStr);
      detections.push({ type: "checkType4", formula: node.origStr });
    }
  }

  // 递归遍历左右子树，传入当前节点的运算符作为父运算符
  if (node.left) traverseAST(node.left, node.op, skipType6, detections);
  if (node.right) traverseAST(node.right, node.op, skipType6, detections);
}

/**
 * 将加减 AST 节点扁平化，返回一个数组。
 * 每个元素是对象 { term: 节点, op: '+' 或 '-' }，其中第一个元素默认 op 为 '+'。
 * 例如，对于表达式 ((D - E) + D) 会返回:
 *    [ { term: D, op: '+' }, { term: E, op: '-' }, { term: D, op: '+' } ]
 *
 * @param {Object} node 加减 AST 节点（type==="operator" 且 op==="+"或"-"）
 * @returns {Array} 数组，每个元素为 { term, op }
 */
function flattenAddSub(node) {
  let result = [];
  function helper(n, inheritedOp = '+') {
    if (n.type === "operator" && (n.op === "+" || n.op === "-")) {
      // 对左子树，保留当前 inheritedOp
      helper(n.left, inheritedOp);
      // 对右子树，若当前运算符为 '-', 则翻转 inheritedOp（'+'->'-', '-'->'+'）
      let newOp = n.op === "+" ? inheritedOp : (inheritedOp === '+' ? '-' : '+');
      helper(n.right, newOp);
    } else {
      result.push({ term: n, op: inheritedOp });
    }
  }
  helper(node, '+');
  return result;
}

/**
 * 当检测到 checkType1 模式时，利用乘法分配律展开表达式。
 * 将形如 L * (X1 ± X2 ± ... ± Xn) 展开为 (L * X1 ± L * X2 ± ... ± L * Xn)。
 *
 * @param {Object} node 匹配 checkType1 的 AST 节点（类型 "operator"，op 为 "*"）
 * @param {Array} addends 数组，每个元素为 { term, op }（由 flattenAddSub 得到）
 * @returns {Object} 展开后的 AST 节点
 */
function expandCheckType1(node, addends) {
  let L = node.left; // 外侧因子（SumY）
  // 构造每一项：L * term，每项带上对应的 op
  let multipliedNodes = addends.map(item => {
    return {
      type: "operator",
      op: "*",
      left: L,
      right: item.term,
      origStr: L.origStr + " * " + item.term.origStr,
      sumType: combineTypes(L.sumType, item.term.sumType, "*"),
      // 注意：展开后每一项单独看均为乘法
      // 不需要进一步展开
    };
  });
  // 将所有乘积按 addends 中的 op 依次连接
  let expanded = multipliedNodes[0];
  for (let i = 1; i < multipliedNodes.length; i++) {
    // 取第 i 个加数的运算符
    let currentOp = addends[i].op; // '+' 或 '-'
    expanded = {
      type: "operator",
      op: currentOp,
      left: expanded,
      right: multipliedNodes[i],
      origStr: "(" + expanded.origStr + " " + currentOp + " " + multipliedNodes[i].origStr + ")",
      sumType: combineTypes(expanded.sumType, multipliedNodes[i].sumType, currentOp)
    };
  }
  return expanded;
}

/**
 * 检查最终结果是否符合模式5：
 * 模式5 (checkType5)：SumN ± SumN（这里允许加或减），且该加减为最终结果（无后续运算步骤）。
 *
 * @param {Object} token 最终 token 对象
 * @param {Array} detections 检测结果数组
 */
function checkType5Final(token, detections) {
  if (token.ast && token.ast.type === "operator" && (token.ast.op === "+" || token.ast.op === "-")) {
    if (token.ast.left && token.ast.left.sumType === "SumN" &&
        token.ast.right && token.ast.right.sumType === "SumN") {
      console.log("检测到特定类型5公式部分: " + token.origStr);
      detections.push({ type: "checkType5", formula: token.origStr });
    }
  }
}

/**
 * 根据运算符和左右操作数的类型，返回合并后的结果类型。
 *
 * 规则：
 *   - 加减：若两个操作数均为 SumY 则结果为 SumY，否则为 SumN
 *   - 乘法：SumY*SumY→SumY, SumY*SumN或SumN*SumY→SumY, SumN*SumN→SumN
 *   - 除法：SumY/SumY→SumN, SumY/SumN→SumY, SumN/SumY→SumN, SumN/SumN→SumN
 *
 * @param {string} type1
 * @param {string} type2
 * @param {string} operator
 * @returns {string}
 */
function combineTypes(type1, type2, operator) {
  if (operator === "+" || operator === "-") {
    if (type1 === "SumY" && type2 === "SumY") return "SumY";
    else if ((type1 === "SumY" && type2 === "SumN") || (type1 === "SumN" && type2 === "SumY"))
      return "SumY";
    else return "SumN";
  }
  if (operator === "*") {
    if (type1 === "SumY" && type2 === "SumY") return "SumY";
    else if ((type1 === "SumY" && type2 === "SumN") || (type1 === "SumN" && type2 === "SumY"))
      return "SumY";
    else return "SumN";
  }
  if (operator === "/") {
    if (type1 === "SumY" && type2 === "SumY") return "SumN";
    else if (type1 === "SumY" && type2 === "SumN") return "SumY";
    else if (type1 === "SumN" && type2 === "SumY") return "SumN";
    else return "SumN";
  }
  return "SumN";
}

/* ================= 测试代码 ================= */

// 示例 FormulaTokens 数组：
// 这里为了测试方便，令 A 为 SumY，D 和 E 为 SumN。
// let FormulaTokens = [
//   { Token: "A", TokenName: "A", SumType: "SumY", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
//   { Token: "D", TokenName: "D", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
//   { Token: "E", TokenName: "E", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null }
// ];

// /*
//   测试 1：检测 checkType1（展开）
//   表达式 "A * (D + E)" 中，
//   A 为 SumY，D 与 E 均为 SumN；
//   应触发 checkType1，展开后变为 "(A * D + A * E)"，
//   返回结果中 expandedFormula 属性应为 "(A * D + A * E)"。
// */
// console.log("===== 测试 checkType1（展开） =====");
// let result1 = evaluateExpressionStepByStep("A * (D + E)", FormulaTokens);
// console.log("最终计算类型：" + result1.finalType);
// console.log("最终公式：" + result1.finalFormula);
// console.log("检测信息：", result1.detections);
// console.log("展开后公式：" + result1.expandedFormula);

// /*
//   测试 2：检测 checkType2
//   表达式 "A / (D - E)" 中，
//   允许括号内为减法，应触发 checkType2，返回分母部分为 "(D - E)"。
// */
// console.log("===== 测试 checkType2 =====");
// let result2 = evaluateExpressionStepByStep("A / (D - E)", FormulaTokens);
// console.log("最终计算类型：" + result2.finalType);
// console.log("最终公式：" + result2.finalFormula);
// console.log("检测信息：", result2.detections);
// let checkType2Det = result2.detections.find(d => d.type === "checkType2");
// if (checkType2Det) {
//   console.log("分母部分:", checkType2Det.denominatorPart); // 应输出 "(D - E)"
// }

// /*
//   测试 3：检测 checkType3
//   表达式 "D / E" 中，
//   应触发 checkType3。
// */
// console.log("===== 测试 checkType3 =====");
// let result3 = evaluateExpressionStepByStep("D / E", FormulaTokens);
// console.log("最终计算类型：" + result3.finalType);
// console.log("最终公式：" + result3.finalFormula);
// console.log("检测信息：", result3.detections);

// /*
//   测试 4：检测 checkType4
//   表达式 "D / A" 中，
//   应触发 checkType4。
// */
// console.log("===== 测试 checkType4 =====");
// let result4 = evaluateExpressionStepByStep("D / A", FormulaTokens);
// console.log("最终计算类型：" + result4.finalType);
// console.log("最终公式：" + result4.finalFormula);
// console.log("检测信息：", result4.detections);

// /*
//   测试 5：检测 checkType5
//   表达式 "D - E" 中，
//   应触发 checkType5。
// */
// console.log("===== 测试 checkType5 =====");
// let result5 = evaluateExpressionStepByStep("D - E", FormulaTokens);
// console.log("最终计算类型：" + result5.finalType);
// console.log("最终公式：" + result5.finalFormula);
// console.log("检测信息：", result5.detections);

// /*
//   测试 6：检测 checkType6
//   测试方案1：表达式 "D * E"
//     应触发 checkType6。
// */
// console.log("===== 测试 checkType6 测试方案1 =====");
// let result6a = evaluateExpressionStepByStep("D * E", FormulaTokens);
// console.log("最终计算类型：" + result6a.finalType);
// console.log("最终公式：" + result6a.finalFormula);
// console.log("检测信息：", result6a.detections);

// /*
//   测试方案2：表达式 "D * E - A"
//     左子表达式 "D * E" 应触发 checkType6。
// */
// console.log("===== 测试 checkType6 测试方案2 =====");
// let result6b = evaluateExpressionStepByStep("D * E - A", FormulaTokens);
// console.log("最终计算类型：" + result6b.finalType);
// console.log("最终公式：" + result6b.finalFormula);
// console.log("检测信息：", result6b.detections);

// /*
//   测试方案3：表达式 "D * E * A"
//     连乘中 "D * E" 不应单独触发 checkType6。
// */
// console.log("===== 测试 checkType6 测试方案3 =====");
// let result6c = evaluateExpressionStepByStep("D * E * A", FormulaTokens);
// console.log("最终计算类型：" + result6c.finalType);
// console.log("最终公式：" + result6c.finalFormula);
// console.log("检测信息：", result6c.detections);


//从输入的数组中排除掉运算符号和括号，得到唯一的变量名
function extractUniqueVariables(tokens) {
  const excludeTokens = ["+", "-", "*", "/", "(", ")"];
  // 不改变原数组，新建一个数组保存符合条件的变量名
  const uniqueVariables = [];
  tokens.forEach(token => {
    // 如果 token 不在排除列表中，并且 uniqueVariables 中尚未存在该 token，则加入结果数组
    if (!excludeTokens.includes(token) && uniqueVariables.indexOf(token) === -1) {
      uniqueVariables.push(token);
    }
  });
  return uniqueVariables;
}


//检测公式中是否有要处理的部分
async function CheckFormula(FormulaAddress) {
  // let FormulaTokens = [
  //   { Token: "N3", TokenName: "Room Revenue", SumType: "SumY", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "L3", TokenName: "Room Exp_AAA", SumType: "SumY", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "M3", TokenName: "Room Exp", SumType: "SumY", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "O3", TokenName: "Ava. Rooms", SumType: "SumY", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "S3", TokenName: "Occ%", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "P3", TokenName: "Test", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "Q3", TokenName: "Test 2", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "J3", TokenName: "ARR", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null },
  //   { Token: "K3", TokenName: "ARR2", SumType: "SumN", ReplaceIndex: null, isCell: false, isOperator: false, isNumber: false, isReplace: false, isOthers: true, TermToReplace: null }
  // ];


  return await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("FormulasBreakdown");
    let formulaCell = sheet.getRange(FormulaAddress); // 目标公式的地址
    formulaCell.load("address,formulas");
    await context.sync();

    let formula = formulaCell.formulas[0][0].replace("=", ""); //去掉等号
    formula = formula.replace(/\s/g, ""); //去掉空号
    console.log("the formula need to check is " + formula);
    let result = evaluateExpressionStepByStep(formula, FormulaTokens);

    console.log("最终计算类型：" + result.finalType);
    console.log("最终公式：" + result.finalFormula);
    console.log("检测信息：", result.detections);
    console.log("展开后公式：" + result.expandedFormula);

    let checkType2Denominator = null;
    let DenominatorTitle = null;

    //下面需要循环检测
    //类型1
    console.log("check 0");
    if (result.detections.length > 0 && result.detections[0].type == "checkType1") {
      console.log("原始公式:" + result.detections[0].formula);
      console.log("需要变为：" + result.expandedFormula);
      let oldFormula = result.detections[0].formula.replace(/\s/g, "");
      let newFormula = result.expandedFormula.replace(/\s/g, "");
      console.log("oldFormula is " + oldFormula);
      console.log("newFormula is " + newFormula);
      formula = formula.replace(oldFormula, newFormula);
      console.log("after change the formula is" + formula);

      //用更新的公式替换原来的公式
      formulaCell.formulas = [[`=${formula}`]];
      await context.sync();
      return "CheckType_1_or_2";
    }
    console.log("check 1");


    //类型2 SumY/(SumN+SumN)的处理
    if (result.detections.length > 0 && result.detections[0].type == "checkType2") {
      let checkType2Formula = result.detections[0].formula;
      checkType2Denominator = result.detections[0].denominatorPart;
      console.log("checkType2 原始公式:" + checkType2Formula);
      console.log("checkType2 分母:" + checkType2Denominator);

      // 分割公式，保留运算符和括号
      checkType2Denominator = checkType2Denominator.replace(/\s/g, "");
      let DenominatorParts = checkType2Denominator.split(/([+\-*/()])/g).filter(part => part.trim() !== "");
      console.log("DenominatorParts is ");
      console.log(DenominatorParts);

      let tempDenominatorParts = [...DenominatorParts]; // 临时存储选中的选项数据

      for (let i = 0; i < tempDenominatorParts.length; i++) {
        let part = tempDenominatorParts[i];
        let isOperator = /[+\-*/()]/.test(part);
        if (!isOperator) {
          // 在 FormulaTokens 中查找匹配的 Token
          const tokenObj = FormulaTokens.find(item => item.Token === part);

          // 如果找到匹配项则替换为 TokenName, A1+B1 替换成ARR+ARR2
          if (tokenObj) {
            tempDenominatorParts[i] = tokenObj.TokenName;
          }
        }
      }

      console.log("after change to title, DenominatorParts is ");
      console.log(tempDenominatorParts);

      DenominatorTitle = tempDenominatorParts.join("");
      console.log("DenominatorTitle is " + DenominatorTitle);

      let DenomonatorVarname = extractUniqueVariables(tempDenominatorParts); // 获取分母的单独变量，从数据透视表中去除
      console.log("DenomonatorVarname is " + DenomonatorVarname);

      //查找是否在第二行已经存在了同样的变量名的Title，可能是上一次运行已经生成，那么就不用进行下面的步骤，不然会重复两列一样的变量

      let sheetUsedRange = sheet.getUsedRange();
      let sheetSecondRow = sheetUsedRange.getRow(1);
      console.log("checkformula88888")
      sheetUsedRange.load("values");
      sheetSecondRow.load("values");
      await context.sync();

      // console.log("checkformula 101010");

      // let VarCheckFormula = sheetSecondRow.find(DenominatorTitle, {
      //   completeMatch: true,
      //   matchCase: true,
      //   searchDirection: "Forward"
      // });
      // console.log("checkformula ABABAB");
      // // VarCheckFormula.load("address");
      // // await context.sync();

      // console.log("checkformula 9999");

      // 通过数组方式查找 DenominatorTitle
      // sheetSecondRow.values 是一个二维数组，只包含一行数据，所以是 [ [val1, val2, ...] ]
      // 取第一行的元素再用 indexOf
      const secondRowValues = sheetSecondRow.values[0];
      const colIndex = secondRowValues.indexOf(DenominatorTitle);

      // 准备一个变量保存匹配到的单元格 Range
      let VarCheckFormula = null;
      
      // if (!VarCheckFormula.isNullObject) { //如果找到了，说明之前一次已经运行过一次SumN+SumN的处理
      // // if (!VarCheckFormula === null) { //如果找到了，说明之前一次已经运行过一次SumN+SumN的处理
      //     console.log("找到了匹配的值");


      if (colIndex !== -1) {
            // 找到了匹配的列索引
        VarCheckFormula = sheetSecondRow.getCell(0, colIndex);
      
        // 如果需要后续操作（如获取地址），可以 load 后再 sync
        // VarCheckFormula.load("address");
        // await context.sync();
        VarValue = VarCheckFormula.getOffsetRange(1,0); // 从ARR+ARR2 单元格往下移动，指向数值的单元格
        VarValue.load("address");
      await context.sync();

        DomoninatorFormula = DenominatorParts.join(""); // 将分母给合并成公式
        console.log("DomoninatorFormula is " + DomoninatorFormula);
        console.log("formula KK is " + formula);
        console.log("VarValue.address" + VarValue.address.split("!")[1]);
        formula = formula.replace(DomoninatorFormula,VarValue.address.split("!")[1]);
        console.log("CheckFormula 010101");

        formulaCell.formulas = [[`=${formula}`]];
        console.log("formulaCell.formulas is " + formulaCell.formulas[0][0]);
        await context.sync();
        // ArrVarPartsForPivotTable = ArrVarPartsForPivotTable.filter(item => !DenomonatorVarname.includes(item)); //去除掉分母中原有的单独变量
      // ArrVarPartsForPivotTable.push(DenominatorTitle); //给数据透视表筛选变量，用新的变量（SumN+SumnN)
      
        console.log("CheckFormula 020202");
        /////下面仅将最新的result公式拷贝到Bridge Data中
        let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
        let  BridgeDataUsedRange = BridgeDataSheet.getUsedRange();
        let BridgeFormulaCell = BridgeDataSheet.getRange(formulaCell.address.split("!")[1]);
        BridgeFormulaCell.formulas = [[formulaCell.formulas[0][0]]];
        console.log("CheckFormula 030303");
        BridgeDataUsedRange.load("rowCount");

        await context.sync();
        console.log("check 6666");
        //拷贝到一整列
        let BridgeResultRange = BridgeFormulaCell.getAbsoluteResizedRange(BridgeDataUsedRange.rowCount-2,1);

        console.log("check 7777");
        BridgeResultRange.copyFrom(BridgeFormulaCell);
        await context.sync();

      } else{

        //在UsedRange的最右边放上SumN, Title 和 对应的公式
        let CurrentUsedRange = sheet.getUsedRange();
        let UsedRangeSecondRow = CurrentUsedRange.getRow(1);
        let UsedRangeThirdRow = CurrentUsedRange.getRow(2);
        CurrentUsedRange.load("address");
        UsedRangeSecondRow.load("address,values");
        UsedRangeThirdRow.load("address,values");
        await context.sync();

        // let ExistingVar = UsedRangeSecondRow.find(DenominatorTitle, {
        //   completeMatch: true,
        //   matchCase: true,
        //   searchDirection: "Forward"
        // });
        // ExistingVar.load("values,address");



        let CurrentRightColumn = getRangeDetails(CurrentUsedRange.address).rightColumn;
        let UsedRightTopCell = sheet.getRange(`${CurrentRightColumn}1`);
        let NewRightFirstCell = UsedRightTopCell.getOffsetRange(0, 1);
        let NewRightSecondCell = UsedRightTopCell.getOffsetRange(1, 1);
        let NewRightThirdCell = UsedRightTopCell.getOffsetRange(2, 1);
        NewRightFirstCell.load("address");
        NewRightSecondCell.load("address");
        NewRightThirdCell.load("address");

        NewRightFirstCell.values = [[`SumN`]];
        NewRightSecondCell.values = [[`${DenominatorTitle}`]];
        NewRightThirdCell.values = [[`=${checkType2Denominator}`]];

        await context.sync();

        console.log("NewRightFirstCell is " + NewRightFirstCell.values[0][0]);
        console.log("NewRightSecondCell is " + NewRightSecondCell.values[0][0]);
        console.log("NewRightThirdCell is " + NewRightThirdCell.values[0][0]);

        let NewSumNAddress = NewRightThirdCell.address.split("!")[1];
        console.log("NewSumNAddress is " + NewSumNAddress);
        formulaCell.formulas = [[`=${formula.replace(checkType2Denominator, NewSumNAddress)}`]] //公式中的分母，直接替换到新的单元格地址
        await context.sync();

        //存放入FormulaTokens中
        FormulaTokens.push({
          Token: NewSumNAddress,
          TokenName: NewRightSecondCell.values[0][0],
          SumType: "SumN",
          ReplaceIndex:null,
          isCell: false,
          isOperator: false,
          isNumber: false,
          isReplace:false,
          isOthers: false,
          TermToReplace: null
        });

        /////////将新生成的NewRight 1-3行 拷贝到Bridge Data中
        let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
        let BridgeDataUsedRange = BridgeDataSheet.getUsedRange();
        console.log("check 3");
        let BridgeDataRightFirstCell = BridgeDataSheet.getRange(NewRightFirstCell.address.split("!")[1]);
        console.log("check 4");
        let BridgeDataRightSecondCell = BridgeDataSheet.getRange(NewRightSecondCell.address.split("!")[1]);
        let BridgeDataRightThirdCell = BridgeDataSheet.getRange(NewRightThirdCell.address.split("!")[1]);
        let BridgeFormulaCell = BridgeDataSheet.getRange(formulaCell.address.split("!")[1]);
        console.log("check 5");
        BridgeDataRightFirstCell.values = [[NewRightFirstCell.values[0][0]]];
        BridgeDataRightSecondCell.values = [[NewRightSecondCell.values[0][0]]];
        BridgeDataRightThirdCell.values = [[NewRightThirdCell.values[0][0]]];
        BridgeFormulaCell.formulas = [[formulaCell.formulas[0][0]]];

        //排除掉SumN 和 SumN等父母表里，使其不在数据透视表中，
        ArrVarPartsForPivotTable = ArrVarPartsForPivotTable.filter(item => !DenomonatorVarname.includes(item));


        const newValue = NewRightSecondCell.values[0][0];
        if (!ArrVarPartsForPivotTable.includes(newValue)) { // 保证是唯一的变量
          ArrVarPartsForPivotTable.push(newValue); // 给数据透视表筛选变量
        }
        // ArrVarPartsForPivotTable.push(NewRightSecondCell.values[0][0]); //给数据透视表筛选变量，用新的变量（SumN+SumnN)

        checkType2Var.push(NewRightSecondCell.values[0][0]); // 把这个新增的变量（SumN+SumnN)放到数组中，因为前一步删除了SumN的单独变量，因此在Process中获取有formula的单元格填充的时候就不能考虑，不然没有SumN等单独的变量形成公式

        BridgeDataUsedRange.load("rowCount");

        await context.sync();
        console.log("check 6");
        //拷贝到一整列
        let BridgeResultRange = BridgeFormulaCell.getAbsoluteResizedRange(BridgeDataUsedRange.rowCount-2,1);
        let BridgeDataRightRange = BridgeDataRightThirdCell.getAbsoluteResizedRange(BridgeDataUsedRange.rowCount-2,1);
        console.log("check 7");
        BridgeResultRange.copyFrom(BridgeFormulaCell);
        BridgeDataRightRange.copyFrom(BridgeDataRightThirdCell);
        BridgeDataRightRange.load("values");
        await context.sync();

        let values = BridgeDataRightRange.values; // 读取计算后的值
        BridgeDataRightRange.values = values;       // 粘贴为静态值
        await context.sync();
        
        console.log("check 8");
                // await deleteProcessSum(); //重新计算的情况，需要把之前
                // checkFormulaGlobalVar = true;
                // await runProgramHandler(); // 因为result的公式已经变化，需要把程序全部重新跑一遍
                // checkFormulaGlobalVar = false;
      } 
      return "CheckType_1_or_2";
    }
    console.log("check 2");
    

    //其他的checkType
    if (result.detections.length > 0 && result.detections[0].type == "checkType3") {
      await checkTypeWarning("公式中不能包含SumN/SumN这样的类型，请修改", result.detections[0].formula);
      return "Error";
    } else if (result.detections.length > 0 && result.detections[0].type == "checkType4") {
      await checkTypeWarning("公式中不能包含SumN/SumY这样的类型，请修改", result.detections[0].formula);
      return "Error";
    } else if (result.detections.length > 0 && result.detections[0].type == "checkType5") {
      await checkTypeWarning("计算的最后一步不能是SumN+SumN", result.detections[0].formula);
      return "Error";
    } else if (result.detections.length > 0 && result.detections[0].type == "checkType6") {
      await checkTypeWarning("计算的最后一步不能是SumN*SumN", result.detections[0].formula);
      return "Error";
    }
    console.log("check end");
  });


}

// 获得Range 四周的行数和列数的信息
//如果范围字符串是单个单元格（例如 AD9），则结束列和结束行与起始列和起始行相同。
//返回一个对象，包含 topRow、bottomRow、leftColumn 和 rightColumn 四个属性。
// function getRangeDetails(rangeStr) {
//   // 使用正则表达式提取列和行信息
//   //   const regex = /([A-Z]+)(\d+):?([A-Z]+)?(\d+)?/;
//   const regex = /([A-Za-z]+)(\d+):?([A-Za-z]+)?(\d+)?/;
//   const match = rangeStr.match(regex);

//   if (match) {
//     const startColumn = match[1];
//     const startRow = parseInt(match[2], 10);
//     const endColumn = match[3] ? match[3] : startColumn;
//     const endRow = match[4] ? parseInt(match[4], 10) : startRow;
//     return {
//       topRow: startRow,
//       bottomRow: endRow,
//       leftColumn: startColumn,
//       rightColumn: endColumn
//     };
//   } else {
//     throw new Error("Invalid range format");
//   }
// }


async function checkTypeWarning(info,Formula){

    const warningMessage = `在result一列的公式中，发现不合规的公式：${Formula}。${info}`;
    console.error(warningMessage);

    // 显示缺失标题的提示框
    const modalOverlay = document.getElementById("modalOverlay");
    const keyWarningPrompt = document.getElementById("keyWarningPrompt");
    const container = document.querySelector(".container");

    const warningElement = document.querySelector("#keyWarningPrompt .waterfall-message");
    warningElement.textContent = warningMessage;

    modalOverlay.style.display = "block";
    keyWarningPrompt.style.display = "flex";
    container.classList.add("disabled");

    await new Promise((resolve) => {
        const confirmButton = document.getElementById("confirmKeyWarning");
        confirmButton.addEventListener(
            "click",
            function () {
                keyWarningPrompt.style.display = "none";
                modalOverlay.style.display = "none";
                container.classList.remove("disabled");
                resolve(); // 继续执行
            },
            { once: true } // 确保事件只触发一次
        );
    });

    return true; // 表示存在缺失的标题


}





/////////////////////////--------------处理和修正公式中错误的情况--------End------------////////////////////



///////////////拷贝Data工作表到Bridge Data中///////////////////
async function CopyDataSheet() {
  await Excel.run(async (context) => {
    let DataSheet = context.workbook.worksheets.getItem("Data");
    let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
    let DataRange = DataSheet.getUsedRange();
    let BridgeStartRange = BridgeDataSheet.getRange("A1");
    BridgeStartRange.copyFrom(DataRange);
    await context.sync();
  });
}

/////////////清除BridgeData工作表中的内容//////////////////////
async function ClearBridgeData() {
  await Excel.run(async (context) => {
    let BridgeDataSheet = context.workbook.worksheets.getItem("Bridge Data");
    let BridgeUsedRange = BridgeDataSheet.getUsedRange();
    BridgeUsedRange.clear();
    await context.sync();
  });
}