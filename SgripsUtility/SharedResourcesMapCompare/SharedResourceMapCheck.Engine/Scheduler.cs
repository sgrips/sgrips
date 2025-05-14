using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Engine
{
    public class Scheduler
    {

        static string _recursiveTime;
        static string _dayOfWeek;
        static string _addressService;
        static private int? _period;
        static DateTime _timeExecution;


        [Flags]
        public enum DayOfWeekOperation : byte
        {
            Monday = 1,
            Tuesday = 2,
            Wednesday = 4,
            Thursday = 8,
            Friday = 16,
            Saturday = 32,
            Sunday = 64
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="timeExecution"></param>
        /// <param name="recursiveTime"></param>
        /// <param name="dayOfWeek"></param>
        /// <returns></returns>
        public static DateTime CalculateNextTimeTask(DateTime timeLastExecution, DateTime startTimeExecution, int? recurisvePeriod, string recursiveTime, DayOfWeekOperation? dayOfWeek)
        {
            DateTime nextDateTimeExecution = timeLastExecution.Date
                                                .AddHours(startTimeExecution.Hour)
                                                .AddMinutes(startTimeExecution.Minute)
                                                .AddSeconds(startTimeExecution.Second);

            bool findNextTime = false;


            if (recurisvePeriod.HasValue)
            {
                DateTime rescursiveNextTime;
                rescursiveNextTime = DateTime.Now.AddMinutes(recurisvePeriod.Value);
                if (DateTime.Now.DayOfYear == rescursiveNextTime.DayOfYear && rescursiveNextTime > nextDateTimeExecution)
                {
                    nextDateTimeExecution = rescursiveNextTime;
                    findNextTime = true;
                }

            }


            if (!findNextTime)
            {

                //([recursiveTime]='W' OR [recursiveTime]='D' OR [recursiveTime]='M')
                if (recursiveTime == "W")
                {
                    List<DayOfWeek> dayOfWeekArr = new List<DayOfWeek>();

                    if ((dayOfWeek & DayOfWeekOperation.Monday) == DayOfWeekOperation.Monday) dayOfWeekArr.Add(DayOfWeek.Monday);
                    if ((dayOfWeek & DayOfWeekOperation.Tuesday) == DayOfWeekOperation.Tuesday) dayOfWeekArr.Add(DayOfWeek.Tuesday);
                    if ((dayOfWeek & DayOfWeekOperation.Wednesday) == DayOfWeekOperation.Wednesday) dayOfWeekArr.Add(DayOfWeek.Wednesday);
                    if ((dayOfWeek & DayOfWeekOperation.Thursday) == DayOfWeekOperation.Thursday) dayOfWeekArr.Add(DayOfWeek.Thursday);
                    if ((dayOfWeek & DayOfWeekOperation.Friday) == DayOfWeekOperation.Friday) dayOfWeekArr.Add(DayOfWeek.Friday);
                    if ((dayOfWeek & DayOfWeekOperation.Saturday) == DayOfWeekOperation.Saturday) dayOfWeekArr.Add(DayOfWeek.Saturday);
                    if ((dayOfWeek & DayOfWeekOperation.Sunday) == DayOfWeekOperation.Sunday) dayOfWeekArr.Add(DayOfWeek.Sunday);

                    _dayOfWeek = Enum.Format(typeof(DayOfWeekOperation), dayOfWeek, "G");

                    double findDelta = double.MinValue;
                    foreach (var dayWeek in dayOfWeekArr)
                    {
                        double delta = (dayWeek - nextDateTimeExecution.DayOfWeek);

                        //se maggiore di zero prendo il giorno nella corrente settimana
                        if (delta > 0)
                        {
                            findDelta = delta;
                            break;
                        }//se minore di zero cerco il primo della settimana successiva
                        else if (findDelta < delta)
                            findDelta = delta + 7;


                    }

                    nextDateTimeExecution = nextDateTimeExecution.AddDays(findDelta);

                }
                else if (recursiveTime == "D")
                    nextDateTimeExecution = nextDateTimeExecution.AddDays(1);
                else if (recursiveTime == "M")
                    nextDateTimeExecution = nextDateTimeExecution.AddMonths(1);

            }
            _timeExecution = nextDateTimeExecution;
            _recursiveTime = recursiveTime;
            _period = recurisvePeriod;

            return nextDateTimeExecution;
        }

        public override string ToString()
        {
            return string.Format("type:{0} day:{1} period:{2} \t time:{3}", _recursiveTime, _dayOfWeek, _period, _timeExecution.ToString("ddd dd-MM-yyyy HH:mm:ss"));
        }

    }
}
