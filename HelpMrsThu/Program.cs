// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using HelpMrsThu;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;


//replace Mrs.Thu with the name or partial name of the workbook
//replace testData with the name of the worksheet
var excel = GetOpeningExcelFile("Mrs.Thu", "testData");
//replace the start row with the value of the starting row in your excel workbook
var startRow = 8;

// get all the range times including: fromDate, toDate and the row of this couple in excel
List<RangeTimeModel> lstRangeTimeModel = new List<RangeTimeModel>();
while (excel.Item3.Cells[startRow, 3].Value != null && excel.Item3.Cells[startRow, 4].Value != null)
{
    lstRangeTimeModel.Add(new RangeTimeModel
    {
        FromDate = excel.Item3.Cells[startRow, 3].Value,
        ToDate = excel.Item3.Cells[startRow, 4].Value,
        Row = startRow
    });
    startRow++;
}


// Example usage:
foreach (var range in lstRangeTimeModel)
{
    DateTime startDate = range.FromDate;
    DateTime endDate = range.ToDate;

    Dictionary<Tuple<int, int>, int> result = WorkingDaysInDuration(startDate, endDate);
    foreach (var entry in result)
    {
        excel.Item2.Cells[range.Row, 4 + entry.Key.Item1].Value = entry.Value;
        excel.Item3.Cells[range.Row, 4 + entry.Key.Item2].Value = entry.Value;
    }
    excel.Item1.Saved = true;
}

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
    }
    catch (Exception e)
    {
        //Log.Error(e.Message);
        Console.WriteLine(e.Message);
        wasFoundRunning = false;
    }

    return Tuple.Create(workbook, worksheet, range);
}
