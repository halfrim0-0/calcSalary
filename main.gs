// スプレッドシート
const sheet = SpreadsheetApp.getActiveSheet();
// カレンダー
let calendar;
// 時給
let salary;
// 開始日
let startDate;
// 終了日
let endDate;

function onOpen() {
    SpreadsheetApp.getActiveSpreadsheet().addMenu('給与計算', [{ name: "実行", functionName: "main" }]);
}

function main() {
    getDataFromSheet();
    const events = getEventFromCalendar();
    const totalTime = calcTotalTime(events);
    const totalSalary = calcTotalSalary(totalTime);
    setDataToSheet(totalTime, totalSalary);
}

// スプレッドシートからデータを取得する
function getDataFromSheet() {
    // カレンダーID
    const calendarID = sheet.getRange('B1').getValue();
    calendar = CalendarApp.getCalendarById(calendarID);
    salary = sheet.getRange('B2').getValue();
    startDate = sheet.getRange('B3').getValue();
    endDate = sheet.getRange('B4').getValue();
}

// カレンダーからイベントを取得する
function getEventFromCalendar() {
    // カレンダーに含まれるイベント
    return calendar.getEvents(startDate, endDate);
}

// 合計勤務時間を計算する
function calcTotalTime(events) {
    let totalTime = 0;
    for (const event of events) {
        // 開始時間
        const startTime = event.getStartTime();
        // 終了時間
        const endTime = event.getEndTime();
        // 勤務時間
        const workTime = (endTime - startTime) / 1000 / 60 / 60;
        totalTime += workTime;
    }

    return totalTime;
}

// 合計給与を計算する
function calcTotalSalary(totalTime) {
    return salary * totalTime;
}

// スプレッドシートに値を記入する
function setDataToSheet(totalTime, totalSalary) {
    sheet.getRange('B6').setValue(totalTime);
    sheet.getRange('B7').setValue(totalSalary);
}