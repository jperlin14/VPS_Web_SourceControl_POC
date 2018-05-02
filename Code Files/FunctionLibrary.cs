using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Threading;

//namespace Vps

    namespace VPS_Web
    {       
        // RandomGenerator class contains methods which generate random strings, number or dates
        public class RandomGenerator  
            {  
            /* Generate a random number between two integer numbers.
               It takes two integer numbers and generate random number between them.
               Parameters:-
                min:- minimum integer number
                max:- maximum integer number
              
               Return Type:-Random integer number which is between minimum and maximum number.
            */        
            public static int RandomNumber(int min, int max)  
                {  
                    Random random = new Random();  
                    return random.Next(min, max);  
                }  
      
            /* Generate a random string with a given size and type.
               It takes size of the string, type of the string like:-Program, Episode or Schedule and generate random string .
               Parameters:-
                size:- Require size of the string
                lowerCase:- If user want random string in lower case then pass true otherwise false
                Type:- Pass type of any string. This string append with return string.
              
               Return Type:-Random string of given size and type.
            */                   
            public static string RandomString(int size, bool lowerCase, string type)  
            {  
                StringBuilder builder = new StringBuilder();  
                Random random = new Random();  
                char ch;  
                type = type.ToUpper();
                            
                builder.Append(type);
                builder.Append(" ");            
                
                for (int i = 0; i < size; i++)  
                {  
                    ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));  
                    builder.Append(ch);  
                }  
                            
                
                if (lowerCase)  
                    return builder.ToString().ToLower();  
            
            return builder.ToString();  
            }  
      
            /* Generate a random password.
               Parameters:-NA
              
               Return Type:-Random password.
            */                   
            public static string RandomPassword()  
            {  
                StringBuilder builder = new StringBuilder();  
                builder.Append(RandomString(4, true,"ANY"));  
                builder.Append(RandomNumber(1000, 9999));  
                builder.Append(RandomString(2, false,"ANY"));  
                return builder.ToString();  
            } 
            
           /* Generate a Date by adding or subtracting days from todays date.
              Parameters:-
               days:-integer number used for addition or subtraction days from todays date and generate date.
              
               Return Type:-Date in string format.
            */                   
            public static string GenerateDate(int days)
            {
                var dateAndTime = DateTime.Now;
                var date = dateAndTime.Date.AddDays(days);
                return date.ToString("MM/dd/yyyy");
            }
            
           /*Generate Current Date and Time
             Parameter:-
             format:-Date and time format which user needs
             Return Type:-
             string:-Return data and time format for given format
           */
            public static string CurrentDataTime(string format)
            {
                var dateAndTime = DateTime.Now;
                return DateTime.Now.ToString(format);
            }      
        } 
        
        // RandomUtil class contains methods which generate random names for reports
        public class ReportUtil
        {
            public static string reportRandomName;
            
           /* Generate random name for excel report and save report.
              Parameters:-
               path:-path of excel report.
               appName:-Report Type.
                
               Return Type:-NA.
            */                   
            
            public static void SaveReportExcel(string path, string appName)  
                {
                    //string path = @"c:\reports"; // or whatever  
                    if (!System.IO.Directory.Exists(path))  
                    {  
                       System.IO.Directory.CreateDirectory(path);
                    }
                    
                    Thread.Sleep(40000);
                    var excelProcesses = System.Diagnostics.Process.GetProcesses().Where(p => p.ProcessName == "EXCEL").ToList();
                                                                                 
                    if (excelProcesses.FirstOrDefault() != null)
                    {
                        Console.WriteLine(excelProcesses.FirstOrDefault().ProcessName.ToString());
                        
                        try                    
                        {                        
                            //Convert reference to process to a reference to an Excel Application Object
                            Microsoft.Office.Interop.Excel.Application oExcelApp = 
                            (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                                                                                    
                            string saveAsFileName = path + "\\" + appName + "_" + RandomGenerator.CurrentDataTime("MMddyyyy_HHmm") + ".xlsx";
                            Console.WriteLine(saveAsFileName);
                            // Console.ReadKey();
                            oExcelApp.DisplayAlerts = false;
                            oExcelApp.ActiveWorkbook.SaveAs(saveAsFileName);
                            oExcelApp.Quit();
                        }
                        catch (Exception e) 
                        {
                            Console.WriteLine("Report Excel file is not open in given time.");
                        } 
                    }
                        
                //Closing Excel processes
                System.Diagnostics.Process[] process=System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (System.Diagnostics.Process p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch { }
                    }
                }
            }
                
           /* Generate random name for  reports which open in Report Viewer.
              Parameters:-
               filepath:-path of excel report.
                
               Return Type:-Random Name for Report.
            */             
            
            public static string GenerateReportRandomReportName(string filepath)      
            {
                int idx = filepath.ToString().LastIndexOf('\\');
                string reportPath = filepath.ToString().Substring(0,idx);
                Console.WriteLine(reportPath);
                if (!System.IO.Directory.Exists(reportPath))
                {
                    try 
                    {
                        // Try to create the directory.
                        System.IO.Directory.CreateDirectory(reportPath);
                    }
                    catch (Exception e) 
                    {
                        Console.WriteLine("Given Path does not found");
                    }    
                } 

                try 
                {
                    string[] tokens = filepath.Split('.');
                    Console.WriteLine(tokens[0].ToString());
                    Console.WriteLine(tokens[1].ToString());

                    filepath = tokens[0] + "_" + RandomGenerator.CurrentDataTime("MMddyyyy_HHmm") + "." + tokens[1];
                       
                                                          
                }
                catch (Exception e) 
                {
                    filepath = null;
                    Console.WriteLine("Given file path is not correct.");
                }    
                return filepath; 
                
            }
        }          
    }

