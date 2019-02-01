(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // 每次加载新页面时都必须运行初始化函数
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // 初始化 FabricUI 通知机制并隐藏
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // 如果未使用 Excel 2016，请使用回退逻辑。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("this is a test to show the highest number in given range");
                $('#button-text').text("Test!");
                $('#button-desc').text("Test to highlight the largest number");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("this is a test to show the highest number in given range");
            $('#button-text').text("Test!");
            $('#button-desc').text("Test to highlight the largest number");

            $('#button-signin').text("Sign In");
            $('#button-goback').text("Sign Out and Back");
                
            loadSampleData();

            // 为突出显示按钮添加单击事件处理程序。
            $('#highlight-button').click(hightlightHighestValue);
            $('#signin-button').click(signIn);
            $('#goback-button').click(goback);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // 针对 Excel 对象模型运行批处理操作
        Excel.run(function (ctx) {
            // 为活动工作表创建代理对象
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // 将向电子表格写入示例数据的命令插入队列
            sheet.getRange("B3:D5").values = values;

            // 运行排队的命令，并返回承诺表示任务完成
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // 针对 Excel 对象模型运行批处理操作
        Excel.run(function (ctx) {
            // 创建选定范围的代理对象，并加载其属性
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // 运行排队的命令，并返回承诺表示任务完成
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // 找到要突出显示的单元格
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // 突出显示该单元格
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);

    }

    function signIn() {
        window.location.href = "https://nike.raymond.dgw.cloud/home2.html";
    }

    function goback() {
        window.location.href = "https://raymond.dgw.cloud/Home.html";
    }
    


    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('选定的文本为:', '"' + result.value + '"');
                } else {
                    showNotification('错误', result.error.message);
                }
            });
    }

    // 处理错误的帮助程序函数
    function errorHandler(error) {
        // 请务必捕获 Excel.run 执行过程中出现的所有累积错误
        showNotification("错误", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // 用于显示通知的帮助程序函数
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
