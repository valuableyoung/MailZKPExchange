using MailZKPExchange.DBConnector;
using MailZKPExchange.SPR;
using SevenZip;
using System;
using System.Data;
 

namespace MailZKPExchange.Holidays
{
    public static class HolidaysRussia
    {
        enum Holidays
        {
            NewYearsDay = 1,            // Новогодние каникулы
            NewYearsHoliday = 2,        // Новогодние каникулы
            NewYearsHolidayContinued = 3, // Новогодние каникулы
            NewYearsHolidayContinued2 = 4, // Новогодние каникулы
            NewYearsHolidayContinued3 = 5, // Новогодние каникулы
            NewYearsHolidayContinued4 = 6, // Новогодние каникулы
            NewYearsHolidayContinued5 = 7, // Новогодние каникулы
            NewYearsHolidayContinued6 = 8, // Новогодние каникулы
            RussianOrthodoxChristmasDay = 9, // Рождество Христово
            InternationalWomensDay = 3,   // Международный женский день
            InternationalWorkersDay = 5,  // Праздник Весны и Труда
            VictoryDay = 9,               // День Победы
            RussiaDay = 12,                // День России
            UnitedPeopleDay = 4
        }

        public static int getHolidaysForCurrentDate()
        {
            int days = 0;
            string sql = @"select dayswait from ZKPMailParam";
            DataRow dayswait = DBExecutor.SelectRow(sql);

            days = 0;    
            for (int i = 0; i < Convert.ToInt32(dayswait[0]); i++) 
            {

                if ((DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Saturday))
                {
                    days = days + 2;
                }
                foreach (Holidays holiday in Enum.GetValues(typeof(Holidays)))
                {
                    DateTime date = new DateTime(2023, 1, (int)holiday);                    
                    if (DateTime.Now.AddDays(i).ToString("MM/dd/yyyy") == date.ToString("MM/dd/yyyy") && (DateTime.Now.AddDays(i).DayOfWeek != DayOfWeek.Saturday && DateTime.Now.AddDays(i).DayOfWeek != DayOfWeek.Sunday))
                    {
                        days++;
                    }
                }
             }
                 
            return Convert.ToInt32(dayswait[0]) + days;
        }

        public static int getHolidaysForCurrentDatefromCalendar()
        {
            int days = 0;
            string sql = @"select dayswait from ZKPMailParam";
            DataRow dayswait = DBExecutor.SelectRow(sql);

            sql = @"select TheDate  from calendar where TypeDate <> 10 and TheDate >= cast(getdate() as date)";
            DataTable holidays = DBExecutor.SelectTable(sql); 

            for (int i = 0; i < Convert.ToInt32(dayswait[0]); i++)
            {

                if ((DateTime.Now.AddDays(i).DayOfWeek == DayOfWeek.Saturday))
                {
                    days = days + 2;
                }
                foreach (DataRow dt in holidays.Rows)
                {
                     
                    if (DateTime.Now.AddDays(i).ToString("MM/dd/yyyy") == Convert.ToDateTime(dt[0]).ToString("MM/dd/yyyy") && (DateTime.Now.AddDays(i).DayOfWeek != DayOfWeek.Saturday && DateTime.Now.AddDays(i).DayOfWeek != DayOfWeek.Sunday))
                    {
                        days++;
                    }
                }
            }

            return Convert.ToInt32(dayswait[0]) + days;
        }
    }
}
