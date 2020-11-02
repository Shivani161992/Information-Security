using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;


namespace trial
{
    class Program
    {
        
        
        static void Main(string[] args)
        {
            double time = 300;
            string[] FileName_User1 = new string[1] { "ajb9b3"  };
            string[] FileName_User2 = new string[1] {  "ajdqnf"  };

            System.Data.DataSet Pvalue_ds = new System.Data.DataSet();
            System.Data.DataTable Pvalue_dt = new System.Data.DataTable("Pvalue");

            for (int i = 0; i < FileName_User1.Length; i++)
            {
                System.Data.DataRow Pvalue_dr = Pvalue_dt.NewRow();
                for (int j=0; j< FileName_User2.Length; j++)
                {
                    string path_user1 = "C:/Users/Shivani/Desktop/Info_Sec_Exce/" + FileName_User1[i] + ".xlsx";
                    string path_user2 = "C:/Users/Shivani/Desktop/Info_Sec_Exce/" + FileName_User2[j] + ".xlsx";

                    System.Data.DataSet ds_user1_data = new System.Data.DataSet();
                    ds_user1_data = Getexceluser1(path_user1); //user one data with duration and date and time filter 
                    System.Data.DataSet ds_user2_data = new System.Data.DataSet();
                    ds_user2_data = Getexceluser2(path_user2); //user two data with duration and date and time filter 

                    System.Data.DataSet week1_DS = new System.Data.DataSet();
                    week1_DS = GetDataManipulation_Week1(ds_user1_data, ds_user2_data, time); //calculate avarage on user 1 and user 2 for week 1
                    System.Data.DataSet week2_DS = new System.Data.DataSet();
                    week2_DS = GetDataManipulation_Week2(ds_user1_data, ds_user2_data, time); // calculate averag of user 1 and user 2 for week 2

                    //System.Data.DataSet week1_DS_WZ = new System.Data.DataSet();
                    //week1_DS_WZ = Getdata_w_week1(week1_DS); //without zero
                    //System.Data.DataSet week2_DS_WZ = new System.Data.DataSet();
                    //week2_DS_WZ = Getdatawithoutzero_week2(week2_DS); //without zero

                    double r1a2a = Getcorrelation_r1a2a(week1_DS, week2_DS);
                    double r1a2b = Getcorrelation_r1a2b(week1_DS, week2_DS);
                    double r2a2b = Getcorrelation_r2a2b(week2_DS, week2_DS);
                    double n = week1_DS.Tables[0].Rows.Count;
                    double z = GetCorrelation(r1a2a, r1a2b, r2a2b, n);
                    double phi = Getphi(z);

                    Console.WriteLine("Value of r1a2a is: {0}", r1a2a);
                    Console.WriteLine("Value of r1a2b is: {0}", r1a2b);
                    Console.WriteLine("Value of r2a2b is: {0}", r2a2b);
                    Console.WriteLine("Value of phi is: {0}", phi);

                   
                    Pvalue_dt.Columns.Add(new System.Data.DataColumn(FileName_User2[j], typeof(string)));

                    

                    Pvalue_dr[FileName_User2[j]] = phi;

                  
                    

                    // Pvalue_dt.Rows[i][j] = phi;

                    //Pvalue_ds.Tables[0].Rows[i]
                    //Pvalue_ds.Tables(0).Rows(4).Item(0) = "Updated Company Name";
                    // Pvalue_dt.Rows[i][j] = phi;
                    Console.WriteLine("Value of phi is: {0}", phi);


                }
                Pvalue_dt.Rows.Add(Pvalue_dr);
            }
             Pvalue_ds.Tables.Add(Pvalue_dt);



        }
       
        public static DataSet Getexceluser1(string path_user1)
        {

            Excel.Application xlApp_user1 = new Excel.Application();
            // Excel.Workbook xlWorkbook_user1 = xlApp_user1.Workbooks.Open(@"C:/Users/Priyal/Desktop/InfoSec_Project/ajb9b3.xlsx");
            Excel.Workbook xlWorkbook_user1 = xlApp_user1.Workbooks.Open(path_user1);
            Excel._Worksheet xlWorksheet_user1 = xlWorkbook_user1.Sheets[1];
            Excel.Range xlRange_user1 = xlWorksheet_user1.UsedRange;
            int rowCount_user1 = xlRange_user1.Rows.Count;
            int colCount_user1 = xlRange_user1.Columns.Count;

            double value_user1 = xlWorksheet_user1.Cells[2, 1].Value2;
            System.Data.DataSet ds_user1 = new System.Data.DataSet();
            System.Data.DataTable dt_user1 = new System.Data.DataTable("MyTable_User1");
           
            dt_user1.Columns.Add(new System.Data.DataColumn("doctets", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("Real First Packet", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("Duration", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("o/d", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("EtoH", typeof(DateTime)));


            for (int j_user1 = 2; j_user1 <= rowCount_user1; j_user1++) //change 500 to rowCount_user1 
            {
                if (xlWorksheet_user1.Cells[j_user1, 10].Value2 != 0)
                {
                    var epochcheck_user1 = xlWorksheet_user1.Cells[j_user1, 6].Value2;
                    System.DateTime dtDateTimecheck_user1 = new System.DateTime(1970, 1, 1, 0, 0, 0, 0);
                    dtDateTimecheck_user1 = Convert.ToDateTime(dtDateTimecheck_user1.AddMilliseconds(epochcheck_user1).ToLocalTime());

                    DateTime Startdate_check_user1 = new DateTime(2013, 02, 04, 8, 00, 00); //change time

                    DateTime enddate_check_user1 = new DateTime(2013, 02, 15, 17, 00, 00); //change time

                    if ((dtDateTimecheck_user1.Date >= Startdate_check_user1.Date))
                    {

                        if ((dtDateTimecheck_user1.Date <= enddate_check_user1.Date))
                        {
                            if ((dtDateTimecheck_user1.TimeOfDay >= Startdate_check_user1.TimeOfDay))
                            {
                                if ((dtDateTimecheck_user1.TimeOfDay <= enddate_check_user1.TimeOfDay))
                                {

                                    System.Data.DataRow dr_user1 = dt_user1.NewRow();
                                   
                                    dr_user1["doctets"] = xlWorksheet_user1.Cells[j_user1, 4].Value2;
                                  
                                    dr_user1["Real First Packet"] = xlWorksheet_user1.Cells[j_user1, 6].Value2;
                                   
                                    dr_user1["Duration"] = xlWorksheet_user1.Cells[j_user1, 10].Value2;

                                    double octets_user1 = xlWorksheet_user1.Cells[j_user1, 4].Value2;
                                    double duration_user1 = xlWorksheet_user1.Cells[j_user1, 10].Value2;
                                    dr_user1["o/d"] = Math.Round((Math.Round(octets_user1, 2) / duration_user1), 4);
                                    var epoch_user1 = xlWorksheet_user1.Cells[j_user1, 6].Value2;
                                    dr_user1["EtoH"] = dtDateTimecheck_user1;


                                    dt_user1.Rows.Add(dr_user1);
                                }
                            }
                        }
                    }
                }

            }
            ds_user1.Tables.Add(dt_user1);

            return ds_user1;
            //var name = ds_user1.Tables[0].Rows[0][11].ToString();

        }
        public static DataSet Getexceluser2(string path_user2)
        {

            Excel.Application xlApp_user1 = new Excel.Application();
            // Excel.Workbook xlWorkbook_user1 = xlApp_user1.Workbooks.Open(@"C:/Users/Priyal/Desktop/InfoSec_Project/ajdqnf.xlsx");
            Excel.Workbook xlWorkbook_user1 = xlApp_user1.Workbooks.Open(path_user2);
            Excel._Worksheet xlWorksheet_user1 = xlWorkbook_user1.Sheets[1];
            Excel.Range xlRange_user1 = xlWorksheet_user1.UsedRange;
            int rowCount_user1 = xlRange_user1.Rows.Count;
            int colCount_user1 = xlRange_user1.Columns.Count;

            double value_user1 = xlWorksheet_user1.Cells[2, 1].Value2;
            System.Data.DataSet ds_user1 = new System.Data.DataSet();
            System.Data.DataTable dt_user1 = new System.Data.DataTable("MyTable_User1");
            //for (int i_user1 = 1; i_user1 <= colCount_user1; i_user1++)
            //{
            //    dt_user1.Columns.Add(new System.Data.DataColumn(xlWorksheet_user1.Cells[1, i_user1].Value2, typeof(string)));

            //}
            dt_user1.Columns.Add(new System.Data.DataColumn("doctets", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("Real First Packet", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("Duration", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("o/d", typeof(float)));
            dt_user1.Columns.Add(new System.Data.DataColumn("EtoH", typeof(DateTime)));


            for (int j_user1 = 2; j_user1 <= rowCount_user1; j_user1++) //change 500 to rowCount_user1 
            {
                if (xlWorksheet_user1.Cells[j_user1, 10].Value2 != 0)
                {
                    var epochcheck_user1 = xlWorksheet_user1.Cells[j_user1, 6].Value2;
                    System.DateTime dtDateTimecheck_user1 = new System.DateTime(1970, 1, 1, 0, 0, 0, 0);
                    dtDateTimecheck_user1 = Convert.ToDateTime(dtDateTimecheck_user1.AddMilliseconds(epochcheck_user1).ToLocalTime());

                    DateTime Startdate_check_user1 = new DateTime(2013, 02, 04, 8, 00, 00); //change time

                    DateTime enddate_check_user1 = new DateTime(2013, 02, 15, 17, 00, 00); //change time

                    if ((dtDateTimecheck_user1.Date >= Startdate_check_user1.Date))
                    {

                        if ((dtDateTimecheck_user1.Date <= enddate_check_user1.Date))
                        {
                            if ((dtDateTimecheck_user1.TimeOfDay >= Startdate_check_user1.TimeOfDay))
                            {
                                if ((dtDateTimecheck_user1.TimeOfDay <= enddate_check_user1.TimeOfDay))
                                {

                                    System.Data.DataRow dr_user1 = dt_user1.NewRow();
                                    //dr_user1["unix_secs"] = xlWorksheet_user1.Cells[j_user1, 1].Value2;
                                    //dr_user1["sysuptime"] = xlWorksheet_user1.Cells[j_user1, 2].Value2;
                                    //dr_user1["dpkts"] = xlWorksheet_user1.Cells[j_user1, 3].Value2;
                                    dr_user1["doctets"] = xlWorksheet_user1.Cells[j_user1, 4].Value2;
                                    //dr_user1["doctets/dpkts"] = xlWorksheet_user1.Cells[j_user1, 5].Value2;
                                    dr_user1["Real First Packet"] = xlWorksheet_user1.Cells[j_user1, 6].Value2;
                                    //dr_user1["Real End Packet"] = xlWorksheet_user1.Cells[j_user1, 7].Value2;
                                    //dr_user1["first"] = xlWorksheet_user1.Cells[j_user1, 8].Value2;
                                    //dr_user1["last"] = xlWorksheet_user1.Cells[j_user1, 9].Value2;
                                    dr_user1["Duration"] = xlWorksheet_user1.Cells[j_user1, 10].Value2;

                                    double octets_user1 = xlWorksheet_user1.Cells[j_user1, 4].Value2;
                                    double duration_user1 = xlWorksheet_user1.Cells[j_user1, 10].Value2;
                                    dr_user1["o/d"] = Math.Round((Math.Round(octets_user1, 2) / duration_user1), 4);
                                    var epoch_user1 = xlWorksheet_user1.Cells[j_user1, 6].Value2;
                                    //System.DateTime dtDateTime_user1 = new System.DateTime(1970, 1, 1, 0, 0, 0, 0);
                                    //dtDateTime_user1 = Convert.ToDateTime(dtDateTime_user1.AddMilliseconds(epoch_user1).ToLocalTime());
                                    //System.DateTime dtDateTime= new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds(epoch).ToShortDateString();
                                    dr_user1["EtoH"] = dtDateTimecheck_user1;


                                    dt_user1.Rows.Add(dr_user1);
                                }
                            }
                        }
                    }
                }

            }
            ds_user1.Tables.Add(dt_user1);

            return ds_user1;

            // var name = ds_user1.Tables[0].Rows[0][11].ToString();
        }

        public static DataSet GetDataManipulation_Week1(DataSet ds_user1, DataSet ds_user2, double time)
        {
            


            System.Data.DataSet Week1_TimeWindowds = new System.Data.DataSet();
            System.Data.DataTable Week_1TimeWindowdt = new System.Data.DataTable("TimeWindow_User1_week1");
            var ID = 0;
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("ID", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("User1_time_week1", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("Average_user1_week1", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("User2_time_week1", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("Average_user2_week1", typeof(string)));
            DateTime Startdate_week1 = new DateTime(2013, 02, 04, 8, 00, 00); //change time

            DateTime enddate_week1 = new DateTime(2013, 02, 08, 17, 00, 00); // change time



            

            var endtime_weel1 = enddate_week1.ToString("HH:mm:ss tt");
            while (Startdate_week1.Date <= enddate_week1.Date)
            {
                var Odate = Startdate_week1.ToString("MM-dd-yyyy");
                var startday = Startdate_week1.ToString("ddd");
                DateTime StartIncre_week1 = new DateTime();
                var starttime_week1 = Startdate_week1.ToString("HH:mm:ss tt");
                StartIncre_week1 = Startdate_week1;

                if (startday != "Sat")
                {
                    if (startday != "Sun")
                    {
                        while (StartIncre_week1.TimeOfDay < enddate_week1.TimeOfDay)
                        {
                            double seconds = +(time-1);

                            StartIncre_week1 = StartIncre_week1.AddSeconds(seconds);
                            var start_timeInc = StartIncre_week1.ToString("HH:mm:ss tt");

                            System.Data.DataRow week_1TimeWindowdr = Week_1TimeWindowdt.NewRow();
                            week_1TimeWindowdr["User1_time_week1"] = Odate + " " + startday + " " + "(" + starttime_week1 + "-" + start_timeInc + ")";
                            week_1TimeWindowdr["User2_time_week1"] = Odate + " " + startday + " " + "(" + starttime_week1 + "-" + start_timeInc + ")";


                            int counter = 0;
                            double average = 0;
                            int count = 0;

                            for (int a = 0; a < ds_user1.Tables[0].Rows.Count; a++)
                            {
                                DateTime namevalue = Convert.ToDateTime(ds_user1.Tables[0].Rows[a][4].ToString()); //gettime 
                                double getod = Convert.ToDouble(ds_user1.Tables[0].Rows[a][3].ToString());

                                if (namevalue.Date == Startdate_week1.Date)
                                {
                                    DateTime ST_week1 = Convert.ToDateTime(starttime_week1);
                                    DateTime STI_week1 = Convert.ToDateTime(start_timeInc);

                                    if (namevalue.TimeOfDay >= ST_week1.TimeOfDay)
                                    {
                                        if (namevalue.TimeOfDay <= STI_week1.TimeOfDay) // check endtime for 8:05:00
                                        {
                                            counter = counter + 1;
                                            average = average + getod;

                                            count = 1;
                                        }
                                    }
                                }
                            }
                            if (count == 1)
                            {

                                double average_Value_Week1 = Math.Round(average, 4) / counter;
                                week_1TimeWindowdr["Average_user1_week1"] = Math.Round(average_Value_Week1, 4);
                            }
                            else
                            {
                                double average_Value_Week1 = 0;
                                week_1TimeWindowdr["Average_user1_week1"] = Math.Round(average_Value_Week1, 0);
                            }


                            int counter2 = 0;
                            double average2 = 0;
                            int count2 = 0;

                            for (int a = 0; a < ds_user2.Tables[0].Rows.Count; a++)
                            {
                                DateTime namevalue = Convert.ToDateTime(ds_user2.Tables[0].Rows[a][4].ToString()); //gettime 
                                double getod = Convert.ToDouble(ds_user2.Tables[0].Rows[a][3].ToString());

                                if (namevalue.Date == Startdate_week1.Date)
                                {
                                    DateTime ST_week1 = Convert.ToDateTime(starttime_week1);
                                    DateTime STI_week1 = Convert.ToDateTime(start_timeInc);

                                    if (namevalue.TimeOfDay >= ST_week1.TimeOfDay)
                                    {
                                        if (namevalue.TimeOfDay <= STI_week1.TimeOfDay) // check endtime for 8:05:00
                                        {
                                            counter2 = counter2 + 1;
                                            average2 = average2 + getod;

                                            count2 = 1;
                                        }
                                    }
                                }
                            }
                            if (count2 == 1)
                            {

                                double average_Value_Week1_user2 = Math.Round(average2, 4) / counter2;
                                week_1TimeWindowdr["Average_user2_week1"] = Math.Round(average_Value_Week1_user2, 4);
                            }
                            else
                            {
                                double average_Value_Week1_user2 = 0;
                                week_1TimeWindowdr["Average_user2_week1"] = Math.Round(average_Value_Week1_user2, 0);
                            }


                            double incre = +1;
                            StartIncre_week1 = StartIncre_week1.AddSeconds(incre);
                            start_timeInc = StartIncre_week1.ToString("HH:mm:ss tt");
                            starttime_week1 = start_timeInc;

                            ID = ID + 1;
                            week_1TimeWindowdr["ID"] = ID;
                            Week_1TimeWindowdt.Rows.Add(week_1TimeWindowdr);

                            // calculate average




                        }
                    }
                }
                double day = +1;
                Startdate_week1 = Startdate_week1.AddDays(day);
            }
            Week1_TimeWindowds.Tables.Add(Week_1TimeWindowdt);
            return Week1_TimeWindowds;
        }
        public static DataSet GetDataManipulation_Week2(DataSet ds_user1, DataSet ds_user2, double time)
        {
            System.Data.DataSet Week1_TimeWindowds = new System.Data.DataSet();
            System.Data.DataTable Week_1TimeWindowdt = new System.Data.DataTable("TimeWindow_User1_week1");
            var ID = 0;
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("ID", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("User1_time_week2", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("Average_user1_week2", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("User2_time_week2", typeof(string)));
            Week_1TimeWindowdt.Columns.Add(new System.Data.DataColumn("Average_user2_week2", typeof(string)));
            DateTime Startdate_week1 = new DateTime(2013, 02, 11, 8, 00, 00); //change time

            DateTime enddate_week1 = new DateTime(2013, 02, 15, 17, 00, 00); // change time



           

            var endtime_weel1 = enddate_week1.ToString("HH:mm:ss tt");
            while (Startdate_week1.Date <= enddate_week1.Date)
            {
                var Odate = Startdate_week1.ToString("MM-dd-yyyy");
                var startday = Startdate_week1.ToString("ddd");
                DateTime StartIncre_week1 = new DateTime();
                var starttime_week1 = Startdate_week1.ToString("HH:mm:ss tt");
                StartIncre_week1 = Startdate_week1;

                if (startday != "Sat")
                {
                    if (startday != "Sun")
                    {
                        while (StartIncre_week1.TimeOfDay < enddate_week1.TimeOfDay)
                        {
                            double seconds = +(time - 1);

                            StartIncre_week1 = StartIncre_week1.AddSeconds(seconds);
                            var start_timeInc = StartIncre_week1.ToString("HH:mm:ss tt");

                            System.Data.DataRow week_1TimeWindowdr = Week_1TimeWindowdt.NewRow();
                            week_1TimeWindowdr["User1_time_week2"] = Odate + " " + startday + " " + "(" + starttime_week1 + "-" + start_timeInc + ")";
                            week_1TimeWindowdr["User2_time_week2"] = Odate + " " + startday + " " + "(" + starttime_week1 + "-" + start_timeInc + ")";

                            int counter = 0;
                            double average = 0;
                            int count = 0;

                            for (int a = 0; a < ds_user1.Tables[0].Rows.Count; a++)
                            {
                                DateTime namevalue = Convert.ToDateTime(ds_user1.Tables[0].Rows[a][4].ToString()); //gettime 
                                double getod = Convert.ToDouble(ds_user1.Tables[0].Rows[a][3].ToString());

                                if (namevalue.Date == Startdate_week1.Date)
                                {
                                    DateTime ST_week1 = Convert.ToDateTime(starttime_week1);
                                    DateTime STI_week1 = Convert.ToDateTime(start_timeInc);

                                    if (namevalue.TimeOfDay >= ST_week1.TimeOfDay)
                                    {
                                        if (namevalue.TimeOfDay <= STI_week1.TimeOfDay) // check endtime for 8:05:00
                                        {
                                            counter = counter + 1;
                                            average = average + getod;

                                            count = 1;
                                        }
                                    }
                                }
                            }
                            if (count == 1)
                            {

                                double average_Value_Week1 = Math.Round(average, 4) / counter;
                                week_1TimeWindowdr["Average_user1_week2"] = Math.Round(average_Value_Week1, 4);
                            }
                            else
                            {
                                double average_Value_Week1 = 0;
                                week_1TimeWindowdr["Average_user1_week2"] = Math.Round(average_Value_Week1, 0);
                            }

                            int counter2 = 0;
                            double average2 = 0;
                            int count2 = 0;
                            for (int a = 0; a < ds_user2.Tables[0].Rows.Count; a++)
                            {
                                DateTime namevalue = Convert.ToDateTime(ds_user2.Tables[0].Rows[a][4].ToString()); //gettime 
                                double getod = Convert.ToDouble(ds_user2.Tables[0].Rows[a][3].ToString());

                                if (namevalue.Date == Startdate_week1.Date)
                                {
                                    DateTime ST_week1 = Convert.ToDateTime(starttime_week1);
                                    DateTime STI_week1 = Convert.ToDateTime(start_timeInc);

                                    if (namevalue.TimeOfDay >= ST_week1.TimeOfDay)
                                    {
                                        if (namevalue.TimeOfDay <= STI_week1.TimeOfDay) // check endtime for 8:05:00
                                        {
                                            counter2 = counter2 + 1;
                                            average2 = average2 + getod;

                                            count2 = 1;
                                        }
                                    }
                                }
                            }
                            if (count2 == 1)
                            {

                                double average_Value_Week1_user2 = Math.Round(average2, 4) / counter2;
                                week_1TimeWindowdr["Average_user2_week2"] = Math.Round(average_Value_Week1_user2, 4);
                            }
                            else
                            {
                                double average_Value_Week1_user2 = 0;
                                week_1TimeWindowdr["Average_user2_week2"] = Math.Round(average_Value_Week1_user2, 0);
                            }







                            double incre = +1;
                            StartIncre_week1 = StartIncre_week1.AddSeconds(incre);
                            start_timeInc = StartIncre_week1.ToString("HH:mm:ss tt");
                            starttime_week1 = start_timeInc;

                            ID = ID + 1;
                            week_1TimeWindowdr["ID"] = ID;
                            Week_1TimeWindowdt.Rows.Add(week_1TimeWindowdr);

                            // calculate average




                        }
                    }
                }
                double day = +1;
                Startdate_week1 = Startdate_week1.AddDays(day);
            }
            Week1_TimeWindowds.Tables.Add(Week_1TimeWindowdt);
            return Week1_TimeWindowds;
        }

       //public static DataSet Getdata_w_week1(DataSet ds_week1) // with zero not included
       // {
       //     System.Data.DataSet ds_w_week1 = new System.Data.DataSet();
       //     System.Data.DataTable dt_w_week1 = new System.Data.DataTable("TimeWindow_User1_week1");
            
       //     dt_w_week1.Columns.Add(new System.Data.DataColumn("ID", typeof(string)));
       //     dt_w_week1.Columns.Add(new System.Data.DataColumn("User1_time_week1", typeof(string)));
       //     dt_w_week1.Columns.Add(new System.Data.DataColumn("Average_user1_week1", typeof(string)));
       //     dt_w_week1.Columns.Add(new System.Data.DataColumn("User2_time_week1", typeof(string)));
       //     dt_w_week1.Columns.Add(new System.Data.DataColumn("Average_user2_week1", typeof(string)));

       //     for(int i=0; i< ds_week1.Tables[0].Rows.Count; i++)
       //     {
       //         double average_user1 = Convert.ToDouble(ds_week1.Tables[0].Rows[i][2].ToString());
       //         double average_user2 = Convert.ToDouble(ds_week1.Tables[0].Rows[i][4].ToString());
                
       //         if((average_user1 !=0) && (average_user2 != 0))
       //         {
                    
                    
       //             dt_w_week1.ImportRow(ds_week1.Tables[0].Rows[i]);
                    
       //         }
       //         else if ((average_user1 != 0) && (average_user2 == 0))
       //         {
                    
       //                 dt_w_week1.ImportRow(ds_week1.Tables[0].Rows[i]);
                    
       //         }
       //         else if ((average_user2 != 0) && (average_user1 == 0))
       //         {
                   
       //                 dt_w_week1.ImportRow(ds_week1.Tables[0].Rows[i]);
                    
       //         }


       //     }
       //     ds_w_week1.Tables.Add(dt_w_week1);
       //     return ds_w_week1;
           
       // }
       // public static DataSet Getdatawithoutzero_week2(DataSet ds_week2) // with zero not included
       // {

       //     System.Data.DataSet ds_w_week2 = new System.Data.DataSet();
       //     System.Data.DataTable dt_w_week2 = new System.Data.DataTable("TimeWindow_User1_week1");
            
       //     dt_w_week2.Columns.Add(new System.Data.DataColumn("ID", typeof(string)));
       //     dt_w_week2.Columns.Add(new System.Data.DataColumn("User1_time_week1", typeof(string)));
       //     dt_w_week2.Columns.Add(new System.Data.DataColumn("Average_user1_week1", typeof(string)));
       //     dt_w_week2.Columns.Add(new System.Data.DataColumn("User2_time_week1", typeof(string)));
       //     dt_w_week2.Columns.Add(new System.Data.DataColumn("Average_user2_week1", typeof(string)));

       //     for (int i = 0; i < ds_week2.Tables[0].Rows.Count; i++)
       //     {
       //         double average_user1 = Convert.ToDouble(ds_week2.Tables[0].Rows[i][2].ToString());
       //         double average_user2 = Convert.ToDouble(ds_week2.Tables[0].Rows[i][4].ToString());
       //         if ((average_user1 != 0) && (average_user2 != 0))
       //         {


       //             dt_w_week2.ImportRow(ds_week2.Tables[0].Rows[i]);
                   
       //         }
       //         else if ((average_user1 != 0) && (average_user2 == 0))
       //         {

       //             dt_w_week2.ImportRow(ds_week2.Tables[0].Rows[i]);
                   
       //         }
       //         else if ((average_user2 != 0) && (average_user1 == 0))
       //         {
       //             dt_w_week2.ImportRow(ds_week2.Tables[0].Rows[i]);
                    
       //         }


       //     }
       //     ds_w_week2.Tables.Add(dt_w_week2);
       //     return ds_w_week2;

       // }

        public static double Getcorrelation_r1a2a(DataSet ds_user1_week1, DataSet ds_user1_week2)
        {
            //average for user 1 week 1
            double UW1_average = 0;
            for (int i = 0; i < ds_user1_week1.Tables[0].Rows.Count; i++)
            {
                //double number = ds_user1_week1.Tables[0].Rows.Count;
                double avg= Convert.ToDouble(ds_user1_week1.Tables[0].Rows[i][2].ToString());
                UW1_average = (avg+ UW1_average);
            }

            // calculate average for user1 week 2
            double U1W2_average = 0;
            for (int i = 0; i < ds_user1_week2.Tables[0].Rows.Count; i++)
            {
                
                double avg = Convert.ToDouble(ds_user1_week2.Tables[0].Rows[i][2].ToString());
                U1W2_average = (avg + U1W2_average);
            }


            double number_week1 = ds_user1_week1.Tables[0].Rows.Count;
            double number_week2 = ds_user1_week2.Tables[0].Rows.Count;

            double avg_user1week1 = UW1_average / number_week1;
            double avg_user1week2 = U1W2_average / number_week2;

           // Console.WriteLine(avg_user1week1);
           // Console.WriteLine(avg_user1week2);

          double numerator = 0;
            double sum_x = 0;
            double sum_y = 0;
            for ( int i = 0; i < number_week1; i++)
            {
                double x = Convert.ToDouble(ds_user1_week1.Tables[0].Rows[i][2].ToString());
                double y = Convert.ToDouble(ds_user1_week2.Tables[0].Rows[i][2].ToString());

                double x_avguw1 = (x - avg_user1week1);
                double y_avguw2 = (y - avg_user1week2);
                double multi_num = x_avguw1 * y_avguw2;
                numerator = numerator + multi_num;

                double square_x = x_avguw1 * x_avguw1;
                sum_x = sum_x + square_x;

                double square_y = y_avguw2 * y_avguw2;
                sum_y = sum_y + square_y;


            }

            double multi_square = sum_x * sum_y;

            double square_rt = Math.Sqrt(multi_square);

            double r1a2a = numerator / square_rt;


            return r1a2a;
        }
        public static double Getcorrelation_r1a2b(DataSet ds_user1_week1, DataSet ds_user2_week2)
        {
            //average for user 1 week 1
            double UW1_average = 0;
            for (int i = 0; i < ds_user1_week1.Tables[0].Rows.Count; i++)
            {
                //double number = ds_user1_week1.Tables[0].Rows.Count;
                double avg = Convert.ToDouble(ds_user1_week1.Tables[0].Rows[i][2].ToString());
                UW1_average = (avg + UW1_average);
            }

            // calculate average for user2 week 2
            double U2W2_average = 0;
            for (int i = 0; i < ds_user2_week2.Tables[0].Rows.Count; i++)
            {

                double avg = Convert.ToDouble(ds_user2_week2.Tables[0].Rows[i][4].ToString());
                U2W2_average = (avg + U2W2_average);
            }

            double number_week1_user1 = ds_user1_week1.Tables[0].Rows.Count;
            double number_week2_user2 = ds_user2_week2.Tables[0].Rows.Count;

            double avg_user1week1 = UW1_average / number_week1_user1;
            double avg_user1week2 = U2W2_average / number_week2_user2;

            // Console.WriteLine(avg_user1week1);
            // Console.WriteLine(avg_user1week2);

            double numerator = 0;
            double sum_x = 0;
            double sum_y = 0;



            for (int i = 0; i < number_week1_user1; i++)
            {
                double x = Convert.ToDouble(ds_user1_week1.Tables[0].Rows[i][2].ToString());
                double y = Convert.ToDouble(ds_user2_week2.Tables[0].Rows[i][4].ToString());

                double x_avguw1 = (x - avg_user1week1);
                double y_avguw2 = (y - avg_user1week2);
                double multi_num = x_avguw1 * y_avguw2;
                numerator = numerator + multi_num;

                double square_x = x_avguw1 * x_avguw1;
                sum_x = sum_x + square_x;

                double square_y = y_avguw2 * y_avguw2;
                sum_y = sum_y + square_y;


            }




            double multi_square = sum_x * sum_y;

            double square_rt = Math.Sqrt(multi_square);

            double r1a2a = numerator / square_rt;
            return r1a2a;
        }
        public static double Getcorrelation_r2a2b(DataSet ds_user1_week2, DataSet ds_user2_week2)
        {
            //average for user 1 week 1
            double UW1_average = 0;
            for (int i = 0; i < ds_user1_week2.Tables[0].Rows.Count; i++)
            {
                //double number = ds_user1_week1.Tables[0].Rows.Count;
                double avg = Convert.ToDouble(ds_user1_week2.Tables[0].Rows[i][2].ToString());
                UW1_average = (avg + UW1_average);
            }
            // calculate average for user2 week 2
            double U2W2_average = 0;
            for (int i = 0; i < ds_user2_week2.Tables[0].Rows.Count; i++)
            {

                double avg = Convert.ToDouble(ds_user2_week2.Tables[0].Rows[i][4].ToString());
                U2W2_average = (avg + U2W2_average);
            }

            double number_week1_user1 = ds_user1_week2.Tables[0].Rows.Count;
            double number_week2_user2 = ds_user2_week2.Tables[0].Rows.Count;

            double avg_user1week1 = UW1_average / number_week1_user1;
            double avg_user1week2 = U2W2_average / number_week2_user2;

            // Console.WriteLine(avg_user1week1);
            // Console.WriteLine(avg_user1week2);

            double numerator = 0;
            double sum_x = 0;
            double sum_y = 0;
            for (int i = 0; i < number_week1_user1; i++)
            {
                double x = Convert.ToDouble(ds_user1_week2.Tables[0].Rows[i][2].ToString());
                double y = Convert.ToDouble(ds_user2_week2.Tables[0].Rows[i][4].ToString());

                double x_avguw1 = (x - avg_user1week1);
                double y_avguw2 = (y - avg_user1week2);
                double multi_num = x_avguw1 * y_avguw2;
                numerator = numerator + multi_num;

                double square_x = x_avguw1 * x_avguw1;
                sum_x = sum_x + square_x;

                double square_y = y_avguw2 * y_avguw2;
                sum_y = sum_y + square_y;


            }

            double multi_square = sum_x * sum_y;

            double square_rt = Math.Sqrt(multi_square);

            double r1a2a = numerator / square_rt;
            return r1a2a;
        }

        public static double GetCorrelation(double r1a2a, double r1a2b, double r2a2b, double n)
        {
            if(r2a2b==1)
            {
                r2a2b = 0.99;
            }
            else if(r1a2a==1)
            {
                r1a2a = 0.99;
            }
            else if(r1a2b == 1)
            {
                r1a2b = 0.99;
            }
            double rm_sqr;  
            double r1a2a_sqr= r1a2a * r1a2a;
            double r1a2b_sqr = r1a2b * r1a2b;
            rm_sqr = (r1a2a_sqr + r1a2b_sqr) / 2;

            double f;
            double rm_sqr_1 = (1- rm_sqr);
            double r2a2b_1 = (1- r2a2b);
            double deno_f=( 2* rm_sqr_1);
            f = r2a2b_1 / deno_f;

            double h;
            double num_mul = f * rm_sqr;
            double num_mul_1 = (1- num_mul);
            h = num_mul_1 / rm_sqr_1;

            double Z1a2b;
            double Z1a2b_log_num = (1 + r1a2b);
            double Z1a2b_log_deno = (1 - r1a2b);
            double Z1a2b_log_value = (Z1a2b_log_num / Z1a2b_log_deno);
            double Z1a2b_log = Math.Log(Z1a2b_log_value);
            Z1a2b = ((1 * Z1a2b_log)/2);

            double Z1a2a;
            double Z1a2a_log_num = (1 + r1a2a);
            double Z1a2a_log_deno = (1 - r1a2a);
            double Z1a2a_log_value = (Z1a2a_log_num / Z1a2a_log_deno);
            double Z1a2a_log = Math.Log(Z1a2a_log_value);
            Z1a2a = ((1 * Z1a2a_log) /2);

            double Z;
            double Z1a2a_Z1a2b = (Z1a2a- Z1a2b);
            double Z_deno = (2 * r2a2b_1 * h);
            double Z_num = n - 3;
            double Z_num_sqr = Math.Sqrt(Z_num);
            double Z_div = (Z_num_sqr / Z_deno);
            Z = (Z1a2a_Z1a2b * Z_div);


            return Z;

        }
        public static double Getphi(double z)
        {
            double p = 0.3275911;
            double a1 = 0.254829592;
            double a2 = -0.284496736;
            double a3 = 1.421413741;
            double a4 = -1.453152027;
            double a5 = 1.061405429;

            int sign;
            if (z < 0.0)
            {
                sign = -1;
            }
                
            else
            {
                sign = 1;
            }
            double x = Math.Abs(z) / Math.Sqrt(2.0);
            double t = 1.0 / (1.0 + p * x);
            double erf = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x);
            double phi= 0.5 * (1.0 + sign * erf);
            return phi;
        }

    }
}
