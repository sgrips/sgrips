using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Engine.Model
{

    public enum CheckTypeEnum
    {
        Exists=0,
        Compare = 1,
    }

    public  class SharedResourcesModel
    {

        public int Id { get; set; }
        public string ShellCommand { get; set; }
        public CheckTypeEnum CheckType { get; set; }
        public string ShellCommandParameters { get; set; }
        public string EmailAlert { get; set; }
        public string EmailSubject { get; set; }
        public DateTime DateTimeStart { get; set; }
        public int? RecursivePeriod { get; set; }
        public string RecursiveTime { get; set; }
        public Scheduler.DayOfWeekOperation? DayOfWeek { get; set; }

        public List<SharedResourceModel> ShareResources { get; set; }

    }

    public class SharedResourceModel
    {

        public string ServerSource { get; set; }
        public string ServerSourceShareList { get; set; }

        public string ServerTarget { get; set; }
        public string ServerTargetShareList { get; set; }
    }
}
