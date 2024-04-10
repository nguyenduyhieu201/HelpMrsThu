// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using HelpMrsThu;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Excel = Microsoft.Office.Interop.Excel;


const int FullDayLeave = 2;
const int PartialDayLeave = 3;
#region Khởi tạo mảng chứa các thông tin
dynamic[,] informations = new dynamic[2000, 15];

#endregion
#region Tính số ngày đi làm theo tháng và điền vào mảng
var excelCalendarSheet = GetOpeningExcelFile("Simulate", "Calendar");
int startRow = 3;
int fromDateColumn = 7;
int toDateColumn = 8;
int hoursPerDayColumn = 9;
int accountColumn = 5;
List<RangeTimeModel> lstRangeModel = new List<RangeTimeModel >();

do
{
    try
    {
        lstRangeModel.Add(new RangeTimeModel
        {
            FromDate = excelCalendarSheet.Item3.Cells[startRow, fromDateColumn].Value,
            ToDate = excelCalendarSheet.Item3.Cells[startRow, toDateColumn].Value,
            HoursPerDay = excelCalendarSheet.Item3.Cells[startRow, hoursPerDayColumn].Value,
            Account = excelCalendarSheet.Item3.Cells[startRow, accountColumn].Value,
            Row = startRow
        });
        startRow++;
    }
    catch
    {
        break;
    }
}
while (excelCalendarSheet.Item3.Cells[startRow, fromDateColumn] != null
     && excelCalendarSheet.Item3.Cells[startRow, toDateColumn] != null
     && excelCalendarSheet.Item3.Cells[startRow, hoursPerDayColumn] != null);

int personIndexInArray = 0;
DateTime startDate = new DateTime();
DateTime endDate = new DateTime();
Dictionary<Tuple<int, int>, int> lstResults = new Dictionary<Tuple<int, int>, int>();
foreach (RangeTimeModel item in lstRangeModel)
{
    var index = GetTheIndexOfTheName(item.Account, informations);
    if (index != -10)
    {
        startDate = item.FromDate;
        endDate = item.ToDate;
        lstResults = WorkingDaysInDuration(startDate, endDate);
        foreach (var result in lstResults)
        {
            if (informations[index + 1, result.Key.Item2] == null)
            {
                informations[index + 1, result.Key.Item2] = result.Value * item.HoursPerDay / 8;
            }
            else
            {
                //var value = double.Parse(informations[index + 1, result.Key.Item2].ToString());
                //value = value + result.Value * item.HoursPerDay / 8;
                informations[index + 1, result.Key.Item2] += result.Value * item.HoursPerDay / 8;

            }

        }
    }
    else
    {
        informations[personIndexInArray, 0] = item.Account;
        startDate = item.FromDate;
        endDate = item.ToDate;
        lstResults = WorkingDaysInDuration(startDate, endDate);
        foreach (var result in lstResults)
        {
            informations[personIndexInArray + 1, result.Key.Item2] = result.Value * item.HoursPerDay / 8;
        }
        personIndexInArray = personIndexInArray + 4;
    }
}
#endregion
#region Tính số ngày nghỉ theo tháng và điền vào mảng 
var excelTMS = GetOpeningExcelFile("Simulate", "TMS");
personIndexInArray = 1;
int startRowLeaveSheet = 2;
int accountColumnLeaveSheet = 3;
int sumDaysColumn = 16;
int leaveFromColumn = 10;
int leaveToColumn = 11;
int leaveColumn = 8;
int leaveTypeColumn = 18;
List<PersonalLeaveDay> lstPersonalLeaveDay = new List<PersonalLeaveDay>();
while (excelTMS.Item3.Cells[startRowLeaveSheet, accountColumnLeaveSheet].Value != null && excelTMS.Item3.Cells[startRowLeaveSheet, sumDaysColumn].Value != null)
{
    var leaveContent = (excelTMS.Item3.Cells[startRowLeaveSheet, leaveColumn].Value.ToString().ToLower().Trim());
    if ((leaveContent.Contains("nghỉ") || leaveContent.Contains("tạm hoãn")))
    {
        lstPersonalLeaveDay.Add(new PersonalLeaveDay
        {
            Account = excelTMS.Item3.Cells[startRowLeaveSheet, accountColumnLeaveSheet].Value,
            SumsLeaveDays = excelTMS.Item3.Cells[startRowLeaveSheet, sumDaysColumn].Value,
            LeaveFrom = excelTMS.Item3.Cells[startRowLeaveSheet, leaveFromColumn].Value,
            LeaveTo = excelTMS.Item3.Cells[startRowLeaveSheet, leaveToColumn].Value,
            LeaveType = excelTMS.Item3.Cells[startRowLeaveSheet, leaveTypeColumn].Value.ToString().Contains("Buổi") 
                            ? PartialDayLeave : FullDayLeave,
        });
        startRowLeaveSheet++;
    }
    else
    {
        startRowLeaveSheet++;
    }
}

foreach(var personalLeaveDay in lstPersonalLeaveDay)
{

    int row = -10;
    for(int i = 0; i < informations.GetLength(0); i++)
    {
        
        if (informations[i, 0] is null) continue;

        if (informations[i, 0].ToString().ToLower().Equals(personalLeaveDay.Account.ToLower()))
        {
            row = i;
        }
    }
    personIndexInArray = row;
    if (personIndexInArray != -10)
    {
        startDate = personalLeaveDay.LeaveFrom;
        endDate = personalLeaveDay.LeaveTo;

        if (personalLeaveDay.LeaveType == PartialDayLeave)
        {
            if (informations[personIndexInArray + 2, personalLeaveDay.LeaveFrom.Month] == null)
            {
                informations[personIndexInArray + 2, personalLeaveDay.LeaveFrom.Month] = 0.5;
            }
            else
            {
                //var value = double.Parse(informations[personIndexInArray + 2, result.Key.Item2].ToString());
                //value = value + result.Value;
                informations[personIndexInArray + 2, personalLeaveDay.LeaveFrom.Month] += 0.5;

            }
            continue;
        }

        lstResults = LeaveDaysInDuration(startDate, endDate);
        foreach (var result in lstResults)
        {
            if (informations[personIndexInArray + 2, result.Key.Item2] == null)
            {
                informations[personIndexInArray + 2, result.Key.Item2] = result.Value;
            }
            else
            {
                //var value = double.Parse(informations[personIndexInArray + 2, result.Key.Item2].ToString());
                //value = value + result.Value;
                informations[personIndexInArray + 2, result.Key.Item2] += result.Value;

            }
        }
    }
}

static int LeaveDaysInMonth(int year, int month, DateTime startDate, DateTime endDate)
{
    var listHoliday = new List<DateTime>
    {
        new DateTime(2024,1,1),
        new DateTime(2024,2,8),
        new DateTime(2024,2,9),
        new DateTime(2024,2,10),
        new DateTime(2024,2,11),
        new DateTime(2024,2,12),
        new DateTime(2024,2,13),
        new DateTime(2024,2,14),
        new DateTime(2024,4,18),
        new DateTime(2024,4,29),
        new DateTime(2024,4,30),
        new DateTime(2024,5,1),
        new DateTime(2024,9,2),
        new DateTime(2024,9,3),
    };
    // Get the number of days in the month
    int numDays = DateTime.DaysInMonth(year, month);

    // Initialize a counter for working days
    int workingDays = 0;

    // Iterate through each day in the month
    for (int day = 1; day <= numDays; day++)
    {
        DateTime date = new DateTime(year, month, day);
        if (month == startDate.Month && day < startDate.Day)
        {
            continue;
        }

        if (month == endDate.Month && day > endDate.Day)
        {
            continue;
        }

        // Check if the day is a weekday (Monday to Friday)
        if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !listHoliday.Any(holiday => holiday.Equals(date)))
        {
            workingDays++;
        }
    }

    return workingDays;
}
//
static Dictionary<Tuple<int, int>, int> LeaveDaysInDuration(DateTime startDate, DateTime endDate)
{
    // Initialize a dictionary to store the count of working days for each month
    Dictionary<Tuple<int, int>, int> workingDaysPerMonth = new Dictionary<Tuple<int, int>, int>();

    // Iterate through each month within the duration
    while (startDate <= endDate)
    {
        int year = startDate.Year;
        int month = startDate.Month;
        // Calculate the number of working days in the current month
        int workingDays = LeaveDaysInMonth(year, month, startDate, endDate);
        // Store the count in the dictionary
        workingDaysPerMonth.Add(new Tuple<int, int>(year, month), workingDays);
        // Move to the next month
        startDate = new DateTime(year, month, 1).AddMonths(1);
    }

    return workingDaysPerMonth;
}

#endregion
#region Tính số ngày OT Theo tháng và điền vào bảng
var excelOT = GetOpeningExcelFile("Simulate", "OT");
int accountColumnInOTSheet = 3;
int OTSummaryColumn = 14;
int MonthColumn = 16;
int startRowOTSheet = 3;
List<OverTimePersonalModel> lstOverTimePersonal  = new List<OverTimePersonalModel>();
do
{
    try
    {
        lstOverTimePersonal.Add(new OverTimePersonalModel
        {
            Account = excelOT.Item3.Cells[startRowOTSheet, accountColumnInOTSheet].Value,
            OverTimeHoursSummary = excelOT.Item3.Cells[startRowOTSheet, OTSummaryColumn].Value,
            Month = (int) excelOT.Item3.Cells[startRowOTSheet, MonthColumn].Value,
        });
        startRowOTSheet++;
    }
    catch (Exception ex)
    {
        break;
    }
}
while (excelOT.Item3.Cells[startRowOTSheet, accountColumnInOTSheet] != null
     && excelOT.Item3.Cells[startRowOTSheet, OTSummaryColumn] != null
     && excelOT.Item3.Cells[startRowOTSheet, MonthColumn] != null);

foreach (var overTimePersonal in lstOverTimePersonal)
{
    int row = -10;
    for (int i = 0; i < informations.GetLength(0); i++)
    {
        if (informations[i, 0] is null) continue;
        if (informations[i, 0].ToString().ToLower().Equals(overTimePersonal.Account.ToLower()))
        {
            row = i;
        }
    }

    personIndexInArray = row;
    if (personIndexInArray != -10)
    {

        if (informations[personIndexInArray + 3, overTimePersonal.Month] == null)
        {
            informations[personIndexInArray + 3, overTimePersonal.Month] = overTimePersonal.OverTimeHoursSummary / 8;
        }
        else
        {
            informations[personIndexInArray + 3, overTimePersonal.Month] += overTimePersonal.OverTimeHoursSummary / 8;
        }
    }
}
#endregion

#region caculate the calendar effort
var lstWorkingDays = WorkingDaysInDuration(new DateTime(2024, 1, 1), new DateTime(2024, 12, 31));
int[] arrWorkingDaysPerMonth = new int[13];
// bắt đầu từ ô 1
foreach(var workingday in lstWorkingDays)
{
    arrWorkingDaysPerMonth[workingday.Key.Item2] = workingday.Value;
}

for (int i = 0; i < informations.GetLength(0);i = i + 4)
{
    if (informations[i, 0] is null) break;
    for (int month = 1; month <=12; month ++)
    {
        double calendarEffort = 0.0;
        double workdaysPerMonth = 0;
        var leavedaysPerMonth = 0;
        var otDaysPerMonth = 0;

        try
        {
            workdaysPerMonth = double.Parse( (informations[i + 1, month]).ToString());
        }
        catch (Exception ex)
        {
            workdaysPerMonth = 0;
        }

        try
        {
            leavedaysPerMonth = informations[i + 2, month];
        }
        catch (Exception ex)

        {
            leavedaysPerMonth = 0;
        }

        try
        {
            otDaysPerMonth = informations[i + 3, month];
        }
        catch (Exception ex)

        {
            otDaysPerMonth = 0;

        }

        if (workdaysPerMonth == 0)
        {
            calendarEffort = 0.0;
            informations[i, month] = calendarEffort;
            continue;
        }
        calendarEffort = (workdaysPerMonth - leavedaysPerMonth + otDaysPerMonth) / arrWorkingDaysPerMonth[month];
        informations[i, month] = calendarEffort;


    }
}


#endregion
#region test array
var excelTemp = GetOpeningExcelFile("Simulate", "Temp");
Excel.Range range = excelTemp.Item2.Range["A1"].Resize[10000, 10000];
range.Value = informations;

Console.WriteLine("");

#endregion

//Check if the name exists in array 
static int GetTheIndexOfTheName (string name, object[,] informations)
{
    int row = -10;
    for (int i = 0; i < informations.GetLength(0); i++)
    {
        if (informations[i, 0] is null) continue;
        if (informations[i, 0].ToString().ToLower().Equals(name))
        {
            row = i;
        }
    }

    return row;
}

#region CaculateManMonthAndExportToExcel
//var excelCalendar = GetOpeningExcelFile("Simulate", "Calendar");
//excelTMS = GetOpeningExcelFile("Simulate", "TMS");

//int startRow = 2;
//int accountColumn = 3;
//int sumDaysColumn = 16;
//int leaveColumn = 8;
//List<PersonalLeaveDay> lstPersonalLeaveDay = new List<PersonalLeaveDay>();
//while (excelTMS.Item3.Cells[startRow, accountColumn].Value != null && excelTMS.Item3.Cells[startRow, sumDaysColumn].Value != null)
//{
//    var leaveContent = (excelTMS.Item3.Cells[startRow, leaveColumn].Value.ToString().ToLower().Trim());
//    Console.WriteLine(leaveContent);
//    if ((leaveContent.Contains("nghỉ") || leaveContent.Contains("tạm hoãn")))
//    {
//        lstPersonalLeaveDay.Add(new PersonalLeaveDay
//        {
//            Account = excelTMS.Item3.Cells[startRow, accountColumn].Value,
//            SumsLeaveDays = excelTMS.Item3.Cells[startRow, sumDaysColumn].Value,
//        });
//        startRow++;
//    }
//    else
//    {
//        startRow++;
//    }
//}


//startRow = 3;
//int fromDateColumn = 7;
//int toDateColumn = 8;
//int hoursPerDayColumn = 9;
//accountColumn = 5;
//List<RangeTimeModel> lstRangeModel = new List<RangeTimeModel >();


//do
//{
//    try
//    {
//        lstRangeModel.Add(new RangeTimeModel
//        {
//            FromDate = excelCalendar.Item3.Cells[startRow, fromDateColumn].Value,
//            ToDate = excelCalendar.Item3.Cells[startRow, toDateColumn].Value,
//            HoursPerDay = excelCalendar.Item3.Cells[startRow, hoursPerDayColumn].Value,
//            Account = excelCalendar.Item3.Cells[startRow, accountColumn].Value,
//            Row = startRow
//        });
//        startRow++;
//    }
//    catch {
//        break;
//    }
//}
//while (excelCalendar.Item3.Cells[startRow, fromDateColumn] != null
//     && excelCalendar.Item3.Cells[startRow, toDateColumn] != null
//     && excelCalendar.Item3.Cells[startRow, hoursPerDayColumn] != null);


//// Join the lists based on the "AccountId" property using LINQ method syntax
//var lstPersonalManMonthInformation = lstRangeModel.Join(lstPersonalLeaveDay,
//                            rangeModel => rangeModel.Account,
//                            pesonalLeaveDay => pesonalLeaveDay.Account.ToLower(),
//                            (rangeModel, personalLeaveDay) => new PersonalManMonthInformation
//                            {
//                                FromDate = rangeModel.FromDate,
//                                ToDate = rangeModel.ToDate,
//                                HoursPerDay = rangeModel.HoursPerDay,
//                                Account = rangeModel.Account,
//                                Row = rangeModel.Row,
//                                SumsLeaveDays = personalLeaveDay.SumsLeaveDays
//                            });

//foreach (var range in lstPersonalManMonthInformation)
//{
//    DateTime startDate = range.FromDate;
//    DateTime endDate = range.ToDate;

//    Dictionary<Tuple<int, int>, int> result = WorkingDaysInDuration(startDate, endDate);
//    foreach (var entry in result)
//    {
//        double manMonth = (((double)(entry.Value) - range.SumsLeaveDays) * range.HoursPerDay)/(double)(entry.Value * 8);
//        excelCalendar.Item3.Cells[range.Row, 18 + entry.Key.Item2].Value = manMonth;

//    }
//}
//excelCalendar.Item1.Saved = true;

//#endregion

//#region CaculateWorkindaysPerMonth
////replace Mrs.Thu with the name or partial name of the workbook
////replace testData with the name of the worksheet
//var excelCalendar = GetOpeningExcelFile("Mrs.Thu", "testData");
////replace the start row with the value of the starting row in your excel workbook
//var startRow = 8;
//int startDateColumn = 3;
//int toDateToColumn = 4;

//// get all the range times including: fromDate, toDate and the row of this couple in excel
//List<RangeTimeModel> lstRangeTimeModel = new List<RangeTimeModel>();
//while (excelCalendar.Item3.Cells[startRow, startDateColumn].Value != null && excelCalendar.Item3.Cells[startRow, toDateToColumn].Value != null)
//{
//    lstRangeTimeModel.Add(new RangeTimeModel
//    {
//        FromDate = excelCalendar.Item3.Cells[startRow, startDateColumn].Value,
//        ToDate = excelCalendar.Item3.Cells[startRow, toDateToColumn].Value,
//        Row = startRow
//    });
//    startRow++;
//}
//// Example usage:
//foreach (var range in lstRangeTimeModel)
//{
//    DateTime startDate = range.FromDate;
//    DateTime endDate = range.ToDate;

//    Dictionary<Tuple<int, int>, int> result = WorkingDaysInDuration(startDate, endDate);
//    foreach (var entry in result)
//    {
//        excelCalendar.Item2.Cells[range.Row, 4 + entry.Key.Item1].Value = entry.Value;
//        excelCalendar.Item3.Cells[range.Row, 4 + entry.Key.Item2].Value = entry.Value;
//    }
//    excelCalendar.Item1.Saved = true;
//}
#endregion
static int WorkingDaysInMonth(int year, int month, DateTime startDate, DateTime endDate)
{
    var listHoliday = new List<DateTime>
    {
        new DateTime(2024,1,1),
        new DateTime(2024,2,8),
        new DateTime(2024,2,9),
        new DateTime(2024,2,10),
        new DateTime(2024,2,11),
        new DateTime(2024,2,12),
        new DateTime(2024,2,13),
        new DateTime(2024,2,14),
        new DateTime(2024,4,18),
        new DateTime(2024,4,29),
        new DateTime(2024,4,30),
        new DateTime(2024,5,1),
        new DateTime(2024,9,2),
        new DateTime(2024,9,3),
    };
    // Get the number of days in the month
    int numDays = DateTime.DaysInMonth(year, month);

    // Initialize a counter for working days
    int workingDays = 0;

    // Iterate through each day in the month
    for (int day = 1; day <= numDays; day++)
    {
        DateTime date = new DateTime(year, month, day);
        if (month == startDate.Month && day < startDate.Day)
        {
            continue;
        }

        if (month == endDate.Month && day > endDate.Day)
        {
            continue;
        }

        // Check if the day is a weekday (Monday to Friday)
        if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !listHoliday.Any(holiday => holiday.Equals(date)))
        {
            workingDays++;
        }
    }

    return workingDays;
}

//
static Dictionary<Tuple<int, int>, int> WorkingDaysInDuration(DateTime startDate, DateTime endDate)
{
    // Initialize a dictionary to store the count of working days for each month
    Dictionary<Tuple<int, int>, int> workingDaysPerMonth = new Dictionary<Tuple<int, int>, int>();

    // Iterate through each month within the duration
    while (startDate <= endDate)
    {
        int year = startDate.Year;
        int month = startDate.Month;
        // Calculate the number of working days in the current month
        int workingDays = WorkingDaysInMonth(year, month, startDate, endDate);
        // Store the count in the dictionary
        workingDaysPerMonth.Add(new Tuple<int, int>(year, month), workingDays);
        // Move to the next month
        startDate = new DateTime(year, month, 1).AddMonths(1);
    }

    return workingDaysPerMonth;
}

//truy cập vào file excel đang chạy
static Tuple<Excel.Workbook, Excel.Worksheet, Excel.Range> GetOpeningExcelFile(string name, string sheetname)
{
    bool wasFoundRunning = false;
    Excel.Workbook workbook = null;
    Excel.Worksheet worksheet = null;
    Excel.Range range = null;
    try
    {
        var xlApp = (Application)Marshal2.GetActiveObject("Excel.Application");
        Excel.Workbooks xlBooks = xlApp.Workbooks;
        foreach (Excel.Workbook xlbook in xlBooks)
        {
            if (xlbook.Name.ToLower().Trim().Contains(name.ToLower().Trim()))
            {
                workbook = xlbook;
                worksheet = (Excel.Worksheet)workbook.Worksheets[sheetname];
                range = worksheet.Cells[1, 1];

            }
        }
        var numBooks = xlBooks.Count;
        wasFoundRunning = true;
        xlApp.Visible = true;
        Marshal.ReleaseComObject(xlApp);

        // Call the garbage collector to release any remaining references
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
    catch (Exception e)
    {
        //Log.Error(e.Message);
        Console.WriteLine(e.Message);
        wasFoundRunning = false;
    }

    return Tuple.Create(workbook, worksheet, range);
}
//if (informations[i, 0].ToString().ToLower() == "hienpv3" || informations[i, 0].ToString().ToLower() == "hiennt25")
//{
//    Console.WriteLine("da tim thay " + informations[i, 0].ToString().ToLower());
//}
//informations[personIndexInArray, 0] = personalLeaveDay.Account;
//DateTime startDate = personalLeaveDay.LeaveFrom;
//DateTime endDate = personalLeaveDay.LeaveTo;
//Dictionary<Tuple<int, int>, int> lstResults = LeaveDaysInDuration(startDate, endDate);
//foreach (var result in lstResults)
//{
//    informations[personIndexInArray + 2, result.Key.Item2] = result.Value;
//}
//personIndexInArray = personIndexInArray + 4;
//Console.WriteLine("-------------------------------------");
//for (int i = 0; i < informations.GetLength(0); i++)
//{
//    if (informations[i, 0] is not null) Console.WriteLine("gia tri la " + informations[i,0]);

//}