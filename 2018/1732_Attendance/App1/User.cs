using System;

namespace _1732_Attendance
{
   class User
   {
      #region *** PROPERTIES ***
      public ulong ID { get; set; }
      public string Name { get; set; }
      public bool Is_Mentor { get; set; }
      public string Status { get; set; }
      public DateTime Check_In_Time { get; set; }
      public TimeSpan User_Hours { get; set; }
      public DateTime Check_Out_Time { get; set; }
      public TimeSpan User_TotalHours { get; set; }
      #endregion

      #region *** PUBLIC METHODS ***
      public void Calculate_Session_Hours()
      {
         TimeSpan timeSpan = Check_Out_Time - Check_In_Time;
         User_Hours += timeSpan;
         User_TotalHours += timeSpan;
      }
      #endregion
   }
}
