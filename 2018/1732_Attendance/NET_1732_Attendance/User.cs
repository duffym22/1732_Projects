using System;

namespace _NET_1732_Attendance
{
    class User
    {
        #region *** PROPERTIES ***
        public ulong ID { get; set; }
        public ulong Secondary_ID { get; set; }
        public string Name { get; set; }
        public bool Is_Mentor { get; set; }
        public string Status { get; set; }
        public DateTime Check_In_Time { get; set; }
        public TimeSpan User_Hours { get; set; }
        public DateTime Check_Out_Time { get; set; }
        public TimeSpan User_TotalHours { get; set; }
        public TimeSpan User_TotalMissedHours { get; set; }
        #endregion

        #region *** PUBLIC METHODS ***
        public void Calculate_Session_Hours()
        {
            TimeSpan timeSpan = Check_Out_Time - Check_In_Time;
            User_Hours += timeSpan;
            User_TotalHours += timeSpan;
        }

        public void Calculate_Missed_Hours(DateTime teamCheckoutTime)
        {
            TimeSpan timeSpan = teamCheckoutTime - Check_In_Time;
            //if timespan is negative then user accidentally checked back in after check out period
            //dont update User_TotalMissedHours (column value will be rewritten with same value)
            if (timeSpan.Ticks > 0)
            {
                User_TotalMissedHours += timeSpan;
            }
        }
        #endregion
    }
}
