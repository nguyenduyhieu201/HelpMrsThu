using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelpMrsThu
{
    public class RangeTimeModel
    {
        public DateTime FromDate { set; get; }
        public DateTime ToDate { set; get; }
        public int Row { set; get; }
        public double HoursPerDay { set; get; }
        public string Account { set; get; }
    }

    public class PersonalLeaveDay
    {
        public DateTime LeaveFrom { set; get; }
        public DateTime LeaveTo { set; get; }
        public string Account { set; get; }
        public double SumsLeaveDays { set; get; }
        public int LeaveType { get; set; }
    }

    public class PersonalManMonthInformation
    {
        public DateTime FromDate { set; get; }
        public DateTime ToDate { set; get; }
        public int Row { set; get; }
        public double HoursPerDay { set; get; }
        public string Account { set; get; }
        public double SumsLeaveDays { set; get; }
    }

    public class OverTimePersonalModel
    {
        public string Account { set; get; }
        public double OverTimeHoursSummary { set; get; }
        public int Month { set; get; }
    }
}
