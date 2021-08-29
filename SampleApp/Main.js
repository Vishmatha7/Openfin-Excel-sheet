var excelInstance;
var currentWorkbook;
var currentWorksheet;
var rowLength = 27;
var columnLength = 12;
var isWorkBook = false;

function addEventListeners(){
	this.initializeExcelEvents();	
}
async function connectExcel() {
	// return new Promise fin.desktop.ExcelService.init()
    //       .then(this.checkConnectionStatus)
    //       .catch(err => console.error(err));

    let that = this;

    return new Promise((resolve, reject) => {
        fin.desktop.ExcelService.init()
            .then(()=>{
                console.error('init worked');
                that.checkConnectionStatus();
                resolve();
            })
            .catch(err => {
                console.error(err);
                reject();
            });
    });

    console.log('connect To Excel - Success');
}

function checkConnectionStatus() {
	let that = this;

    fin.desktop.Excel.getConnectionStatus(connected => {
      if (connected) {
        console.log('Already connected to Excel, synthetically raising event.');
        that.onExcelConnected();
      } else {
        console.log("Excel not connected");
        this.openExcel();
        // this.openExcelNotOpened();
        setTimeout(async function () {
            await fin.desktop.Excel.addWorkbook()
                .then(that.updateData)
                .catch(err => console.error(err));
            }, 5000)
        // setDisplayContainer(view.noConnectionContainer);
      }
    });
}

async function openExcelNotOpened() {
    await fin.desktop.Excel.run()
        .then(this.addWorkbook)
        .catch(err => console.error(err));
}
function onExcelConnected() {
	excelInstance = fin.desktop.Excel;
	
	console.log("Excel connected - onExcelConnected");
        excelInstance.addEventListener("workbookAdded", onWorkbookAdded);
        excelInstance.addEventListener("workbookOpened", onWorkbookAdded);
		
		fin.desktop.Excel.getWorkbooks(workbooks => {
            for (var i = 0; i < workbooks.length; i++) {
                // addWorkbookTab(workbooks[i].name);
                workbooks[i].addEventListener("workbookActivated", onWorkbookActivated);
                workbooks[i].addEventListener("sheetAdded", onWorksheetAdded);
            }

            if (workbooks.length) {
                selectWorkbook(workbooks[0]);
            }
        });
}

function onExcelDisconnected() {
	console.log("Excel disconnected - onExcelDisconnected");
}

 function openExcel() {
    console.log('open Excel');
    return fin.desktop.Excel.run();
}

function activateWorkbook() {
	var workbook = fin.desktop.Excel.getWorkbookByName('Book2');
    workbook.activate();
}

function sendDataToExcel() {
	var sheet1 = fin.desktop.Excel.getWorkbookByName('Book2').getWorksheetByName('Sheet1');

// A little fun with Pythagorean triples:
sheet1.setCells([
  ["A", "B", "C"],
  [  3,   4, "=SQRT(A2^2+B2^2)"],
  [  5,  12, "=SQRT(A3^2+B3^2)"],
  [  8,  15, "=SQRT(A4^2+B4^2)"],
], "A1");

// Write the computed values to the console:
sheet1.getCells("C2", 0, 2, cells => {
  console.log(cells[0][0].value);
  console.log(cells[1][0].value);
  console.log(cells[2][0].value);
});
}

function initializeExcelEvents() {
        fin.desktop.ExcelService.addEventListener("excelConnected", onExcelConnected);
        fin.desktop.ExcelService.addEventListener("excelDisconnected", onExcelDisconnected);
        console.log('initializeExcelEvents fired')
}
		
function selectWorksheet(sheet) {
        if (currentWorksheet === sheet) {
            return;
        }

        currentWorksheet = sheet;
}

function addWorkbook() {
	fin.desktop.Excel.addWorkbook();
}
    function addWorkbookTab(name) {
        currentWorkbook = getWorkbookTab(name);
        button.addEventListener("click", onWorkbookTabClicked);
    }
function onWorkbookAdded(event) {
        var workbook = event.workbook;

        workbook.addEventListener("workbookActivated", onWorkbookActivated);
        workbook.addEventListener("sheetAdded", onWorksheetAdded);
        workbook.addEventListener("sheetRemoved", onWorksheetRemoved);
        workbook.addEventListener("sheetRenamed", onWorksheetRenamed);

        addWorkbookTab(workbook.name);

        setDisplayContainer(view.workbooksContainer);
}

    function onWorkbookActivated(event) {
        selectWorkbook(event.target);
    }

    function selectWorkbook(workbook) {
        currentWorkbook = workbook;
        currentWorkbook.getWorksheets(updateSheets);
    }

function updateSheets(worksheets) {
    for (var i = 0; i < worksheets.length; i++) {

        addWorksheetTab(worksheets[i]);
    }

    selectWorksheet(worksheets[0]);

    if (currentWorksheet) {
        currentWorksheet.setCells([
            ["Test Run"]
        ], "A6");
    }
}

function addWorkSheet() {
    currentWorkbook.addWorksheet();
}


function addWorksheetTab(worksheet) {
    worksheet.addEventListener("sheetActivated", onSheetActivated);
}

function onWorksheetAdded(event) {
    addWorksheetTab(event.worksheet);
}

function onSheetActivated(event) {
    selectWorksheet(event.target);
}

function onSheetChanged(event) {
    updateCell(cell, event.data.value, event.data.formula);
}

function onSelectionChanged(event) {
    var cell = view.worksheetBody.children[event.data.row - 1].children[event.data.column];
    selectCell(cell, true);
}

function updateData() {

// A little fun with Pythagorean triples:
    setInterval(()=>{
        this.writeData();
    }, 1000)
}

function writeData() {
    var sheet1 = currentWorksheet;

    sheet1.setCells([
        ["A", "B",],
        [  Math.floor(Math.random() * 100) + 1,   Math.floor(Math.random() * 100) + 1]
    ], "A1");
}
function updateCell(cell, value, formula) {
    cell.innerText = value ? value : "aaaa";
    cell.title = formula ? formula : "";
}