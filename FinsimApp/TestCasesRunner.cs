using AutoIt;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Winium;
using System.Diagnostics;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using excel = Microsoft.Office.Interop.Excel;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using System.Reflection;
using OpenQA.Selenium.Remote;
using System.Net;
using System.Security.Policy;
using System.Collections.Generic;
using System.Configuration;
using NLog;
using OpenQA.Selenium.Support.UI;

namespace FinsimApp
{
    class TestCasesRunner
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        bool exp = false;
        int i = 0, j = 0;
        int proclen = 0;
        string key = string.Empty;
        string executetype = string.Empty;
        string lastexecutetype = string.Empty;
        string execute = string.Empty;
        string tstname = string.Empty;
        string format = string.Empty;
        string otpinsert = " ";
        bool pinentr = false;
        string excelpath = string.Empty;
        string excelFileName = string.Empty;
        string dirpath = string.Empty;
        string atm = string.Empty;
        string fip = string.Empty;
        string fport = string.Empty;
        string uname = string.Empty;
        string lanflag = string.Empty;
        string cardpool = string.Empty;
        string card = string.Empty;
        string cardno = string.Empty;
        string pin = string.Empty;
        string trantype = string.Empty;
        string url = string.Empty;
        WiniumDriver driver = null;
        string driverPath = string.Empty;
        bool stat = false;
        string expoutput = string.Empty;

        DBValidations db = new DBValidations();
        WorkDistributer wd = new WorkDistributer();
        public void TestRunner()
        {
            
            
            DesktopOptions options = new DesktopOptions();
            

            excel.Application x1 = new excel.Application();
            dirpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            excelFileName = ConfigurationManager.AppSettings["ExcelFileName"];
            excelpath = dirpath + "\\"+excelFileName;
            logger.Info("Excel Path : "+excelpath);
            excel.Workbook wb = x1.Workbooks.Open(excelpath);
            //Adding TestCase Sheet
            excel.Workbook wb1 = x1.Workbooks.Add();
            excel.Worksheet testsheet = wb1.ActiveSheet;
            //Loading Data
            excel._Worksheet sheet1 = wb.Sheets[1];
            excel._Worksheet sheet2 = wb.Sheets[2];
            excel._Worksheet sheet3 = wb.Sheets[3];
            excel._Worksheet sheet4 = wb.Sheets[4];
            excel.Range rng1 = sheet1.UsedRange;
            excel.Range rng2 = sheet2.UsedRange;
            excel.Range rng3 = sheet3.UsedRange;
            excel.Range rng4 = sheet4.UsedRange;
            int trnno = rng3.Count;
            //Configuration sheet data
            logger.Info("Fetching Data For Configuration");
            atm = rng1.Cells[1][2].value2;
            logger.Info("ATM NAME  " + atm);
            fip = rng1.Cells[2][2].value2;
            logger.Info("FINSIM IP   " + fip);
            fport = rng1.Cells[3][2].value2;
            logger.Info(" FINSIM PORT   " + fport);
            uname = rng1.Cells[9][2].value2;
            logger.Info("USER NAME   " + uname);
            lanflag = rng1.Cells[10][2].value2;
            logger.Info("   " + lanflag);
            logger.Info("Fetching Completed");
            db.GenerateConData(rng1);

            // Checking Process running Already or not
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Contains("ATM Simulator - ") || clsProcess.ProcessName.Contains("ATClient"))
                {
                    proclen = 1;

                }
            }
            //Card pool data
            cardpool = rng2.Cells[1][2].value2;
            logger.Info(" CARD POOL  "+cardpool);
            card = rng2.Cells[2][2].value2;
            logger.Info( " CARD NAME  " + card);
            cardno = rng2.Cells[3][2].value2;
            logger.Info("CARD NO  " + cardno);
            pin = rng2.Cells[4][2].value2;
            logger.Info(" PIN  " + pin);
            trantype= rng2.Cells[5][2].value2;
            logger.Info(" TRAN-TYPE  " + trantype);
            
            int serviceport = WorkDistributer.FreeTcpPort();
            driverPath = "WiniumDesktopDriver";
            
            WiniumDriverService service = WiniumDriverService.CreateDesktopService(driverPath);
            service.HostName = "localhost";
            service.Port=serviceport;
            service.Start();

            options.ApplicationPath = "ATClient.exe";
            options.Arguments = "no-sandbox";
            options.DebugConnectToRunningApp = false;
            options.LaunchDelay = 2;

            url = "http://127.0.0.1:"+serviceport;
            driver = new WiniumDriver(new Uri(url), options);

            
            Actions ac = new Actions(driver);
            //starting report module
            string dateforrpt = DateTime.Now.ToString("yyyyMMdd_h:mm:ss");
            string datefortest = DateTime.Now.ToString("yyyyMMdd");
            string rptpath = dirpath + "\\Reports\\" + dateforrpt; // + "\\report.html";
            logger.Info(rptpath);
            //TestCase File Location 
            string testfilename = ConfigurationManager.AppSettings["TestCaseFileName"];
            string testfilelocation =dirpath+"\\Reports\\"+testfilename+ datefortest+ ".xlsx";
            testsheet.Cells[2][1]="SNO";
            testsheet.Cells[3][1]= "TESTCASE NAME";
            testsheet.Cells[4][1] = "EXPECTED OUTPUT";
            testsheet.Cells[5][1] = "ACTUAL OUTPUT";
            testsheet.Cells[6][1] = "DATE OF EXECUTION";
            testsheet.Cells[7][1] = "STAN";
            testsheet.Cells[8][1] = "STATUS";
            testsheet.Range["B1:H1"].Interior.Color = System.Drawing.Color.LightGreen;
            testsheet.Range["B1:H1"].Font.Bold = true;
            testsheet.Range["B1:H1"].Font.Color = System.Drawing.Color.White;
            testsheet.Range["B1:H1"].Font.Size = 12;
            testsheet.Range["B1:H1"].RowHeight=30;
            testsheet.Range["B1:H1"].HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;
            testsheet.Range["B1:H1"].VerticalAlignment = excel.XlVAlign.xlVAlignCenter;
            var htmlreport = new ExtentHtmlReporter(rptpath);
            htmlreport.Config.Encoding = "utf-8";
            htmlreport.Config.DocumentTitle = "Finsim ATM Test Report";
            htmlreport.Config.ReportName = "report.html";
            
            //report setup
            var extent = new ExtentReports();
            extent.AttachReporter(htmlreport);
            extent.AddSystemInfo("ATM Name",atm);
            extent.AddSystemInfo("Card Pool", cardpool);
            extent.AddSystemInfo("Card Name", card);
            extent.AddSystemInfo("Card Pin", pin);
            extent.AddSystemInfo("User Name", uname);

            var test = extent.CreateTest("Pre-requisite Test");
           
            logger.Info("Total Finsim Instances are Running : "+proclen);
            // Opening Finsim For Transactions

            //checking if atclient is already running if running open new instance
            if (proclen > 0)
            {
                AutoItX.WinActivate("You Are Already Running FINsim");
                driver.FindElementByName("Yes").Click();
                Thread.Sleep(500);
                // Is running
                driver.FindElementByName("New Server").Click();
                Thread.Sleep(1000);
                AutoItX.WinActivate("PaySim Classic - Connect To Server");
                logger.Info("Connecting to IP : " + fip);
                wd.VerifyTextonTextBox(driver.FindElementById("1904"), fip);
                logger.Info("Connecting to Port : " + fport);
                wd.VerifyTextonTextBox(driver.FindElementById("1908"), fport);
                driver.FindElementByName("Connect").Click();
                test.Pass("Connect Sucessfully to Port " + fport);

                driver.FindElementById("1821").Click();
                wd.VerifyTextonTextBox(driver.FindElementById("1821"), uname);
                Thread.Sleep(1000);
                driver.FindElementById("1818").Click();
                wd.VerifyTextonTextBox(driver.FindElementById("1818"), atm);
                Thread.Sleep(2000);
                wd.Click(driver, driver.FindElementByName(atm));
                driver.FindElementByName("Force").Click();
                test.Pass("ATM Open Sucessfully ");
                wd.VerifyCheckBoxSelected(driver.FindElementById("25484"), true);
                test.Pass("EMV CheckBox Enabled");
            }
            else
            {
                driver.FindElementByName("New Server").Click();
                Thread.Sleep(1000);
                AutoItX.WinActivate("PaySim Classic - Connect To Server");
                logger.Info("Connecting to IP : "+fip);
                wd.VerifyTextonTextBox(driver.FindElementById("1904"), fip);
                logger.Info("Connecting to Port : " + fport);
                wd.VerifyTextonTextBox(driver.FindElementById("1908"), fport);
                driver.FindElementByName("Connect").Click();
                test.Pass("Connect Sucessfully to Port " + fport);

                driver.FindElementById("1821").Click();
                wd.VerifyTextonTextBox(driver.FindElementById("1821"), uname);
                Thread.Sleep(1000);
                driver.FindElementById("1818").Click();
                wd.VerifyTextonTextBox(driver.FindElementById("1818"), atm);
                Thread.Sleep(2000);
                wd.Click(driver, driver.FindElementByName(atm));
                driver.FindElementByName("Force").Click();
                test.Pass("ATM Open Sucessfully ");
                
                wd.VerifyCheckBoxSelected(driver.FindElementById("25484"), true);
                test.Pass("EMV CheckBox Enabled");
            }

           
            Thread.Sleep(2000);



            // performing  Transactions using selected Atm
            wd.SelectFromDropDown("ATM Simulator - " + atm, "[CLASS:ComboBox; INSTANCE:1]", cardpool);
            test.Pass("Selected Cardpool Sucessfully " + cardpool);
            Thread.Sleep(3000);
            IWebElement Card = driver.FindElementByName(card);

            // Find the last real row
            int lastUsedRow = wd.ExcelRows(rng3);
            logger.Info("Total no of times loop runing : "+(lastUsedRow));
            logger.Info("Total no of TestCases : " + (lastUsedRow-1));
            test.Info("Total number of TestCases " + (lastUsedRow - 1));

            int lastUsedColumn = wd.ExcelColumns(rng3);
            logger.Info("Total number of Columns"+lastUsedColumn);

            // starting transaction 
            j = 5;

            for (i = 2; i <= lastUsedRow; i++)
            {
                
                logger.Info("Running outer loop iteration"+i);
                try {
                    exp = false;
                    // execution type if Y then run 
                    executetype = rng3.Cells[3][i].value2;
                    format = rng3.Cells[2][i].value2;

                    lastexecutetype = rng3.Cells[3][i - 1].value2;
                    //checking transaction Flow 
                    if (trantype == "ON-US")
                    {
                        execute = rng3.Cells[4][i].value2;
                        tstname = rng3.Cells[1][i].value2;
                        
                    }
                    else
                    {
                        execute = rng4.Cells[2][i].value2;
                    }

                    //executing testcase if Execute = Y
                    if (execute != null && execute.Equals("Y"))
                    {
                        string dateofexecution= DateTime.Now.ToString("yyyy/MM/dd");
                        testsheet.Cells[2][i] = i - 1;
                        testsheet.Cells[3][i] = tstname;
                        testsheet.Cells[6][i] = dateofexecution;
                        var test1 = extent.CreateTest(tstname + " TestCase " + (i - 1));
                        logger.Info("Current Transaction Format " + format);
                        //Main Transactiion Execution

                        if (lastexecutetype != "C")
                        {
                            if (format.Contains("CARD BASED"))
                            {
                                logger.Info("Going to Insert Card : ");
                                ac.DoubleClick(Card).Build().Perform();
                                logger.Info("Card Inserted Sucessfully : ");
                                test1.Info("Card Inserted Sucessfully : ");
                                Thread.Sleep(1000);
                                AutoItX.WinActivate("ATM Simulator - " + atm);
                                Thread.Sleep(2000);
                                // language selection 
                                if (!lanflag.Equals("N"))
                                {
                                    logger.Info("Selecting Language");
                                    wd.Click(driver, driver.FindElementByName(lanflag));
                                }

                                // Pin entry
                                Thread.Sleep(2000);
                                char[] p = pin.ToArray();
                                foreach (char pi in p)
                                {
                                    Thread.Sleep(100);
                                    wd.SendKeys(pi.ToString());
                                }
                                pinentr = true;
                                logger.Info("Pin Enter Sucessfully : "+pin);
                                test1.Info("Pin Enter Sucessfully " + pin);

                            }
                            else
                            {
                                logger.Info(DateTime.Now.ToString("HH:mm:ss:ffffff") + "   " + "Sending A for biomatric Transaction");
                                wd.SendKeys("A");
                            }


                        }



                        Thread.Sleep(3000);
                        // inserting FDKs from Excel sheet
                        //inner for starts
                       
                        for (j = 6; j <= lastUsedColumn; j++)
                        {
                            expoutput = rng3.Cells[5][i].value2;
                            testsheet.Cells[4][i] = expoutput;
                            logger.Info("Stat Variable value "+ stat);
                            stat = false;
                            //if (stat == true)
                            //{
                            //    break;
                            //}
                            //logger.Info("Key Insertion Iteration " + i); 
                            logger.Info("Key Iteration " + j);
                            try {

                                if (pinentr== true) {
                                    logger.Info("Waiting for main screen");
                                    wd.WaitForMainScreen(driver);
                                    Thread.Sleep(1000);
                                    pinentr = false;
                                }
                                if (trantype == "ON-US")
                                {
                                    key = rng3.Cells[j][i].value2;
                                }
                                else
                                {
                                    key = rng4.Cells[j][i].value2;
                                }

                                
                                // sending keys
                                try
                                {
                                    if (key != null)
                                    {
                                        logger.Info("key is not Empty "+ key);
                                        logger.Info("key Iteration "+j);

                                        if (format.Contains("BIO OTP") && j == 8)
                                        {
                                            logger.Info("Update OTP in DB");
                                            db.UPDATEOTP(cardno);
                                            Thread.Sleep(300);
                                            otpinsert = "YES";
                                        }

                                        int m = 0;
                                        int l = key.Length;
                                        if (l > 1)
                                        {
                                            char[] p = key.ToArray();
                                            for (m = 0; m < l - 1; m++)
                                            {
                                                char pi = p[m];
                                                Thread.Sleep(100);
                                                wd.SendKeys(pi.ToString());
                                            }

                                            int len = key.Length - 1;
                                            logger.Info("printing len " + len);
                                            string k = key.Substring(len, 1);
                                            logger.Info("Clicking key with numbers " + k);
                                            wd.Click(driver, driver.FindElementByName(k));
                                            if (otpinsert == "YES")
                                            {
                                                otpinsert = " ";
                                                Thread.Sleep(2000);
                                            }
                                           
                                        }
                                        else
                                        {
                                            Thread.Sleep(500);
                                                logger.Info("Clicking Key : " + key);
                                                //Thread.Sleep(1000);
                                                wd.Click(driver, driver.FindElementByName(key));
                                                //wd.SendKeys(key);
                                            



                                        }

                                    }
                                    if(key == null)
                                    {
                                        stat = false;
                                        logger.Info("Enter in with out key else");
                                        try {
                                            bool display = wd.WaitToDisplayElement(driver, "D");
                                            logger.Info("display boolean " + display);
                                            if (display == true && executetype != "C")
                                            {
                                                logger.Info("Canceling Transaction");
                                                driver.FindElementByName("Cancel").Click();
                                                //driver.FindElementById("25469").Click();
                                                //wd.Click(driver, driver.FindElementByName("Cancel"));
                                                logger.Info("Exit from Loop");
                                                stat = true;
                                            }
                                            else
                                            {
                                                logger.Info("Exit from loop Element Not Displayed");
                                                stat = true;
                                            }
                                            logger.Info("Exit tryblock in frequency check");
                                        }
                                        catch
                                        {
                                            logger.Info("          ");
                                            stat = true;
                                        }

                                       

                                    }

                                }catch
                                {
                                    stat = false;
                                    logger.Info("Unable to send FDKs Taking ScreeenShot");
                                    WorkDistributer.TakeScreenshot(driver, tstname);
                                    logger.Info("Screen Shot Captured");
                                    exp = true;
                                    Thread.Sleep(3000);
                                    driver.FindElement(By.Name("Options")).Click();
                                    driver.FindElement(By.Name("Reset TCP/IP Comms")).Click();
                                    
                                    Thread.Sleep(8000);
                                    stat = true;
                                }

                            }
                            catch (Exception e) {
                                logger.Info("Exception Occur in keys : "+ e.Message);
                            }

                            if (stat == true)
                            {
                                logger.Info("Breaking Loop");
                                break;
                            }

                            
                        }


                        // inner For END
                        logger.Info("Exp variaable value "+ exp);
                        if (exp == false)
                        {
                            logger.Info("Entered in Exp if");
                            if (executetype != "C")
                            {
                                wd.WaitForTransactionToFinish(driver);
                            }


                            wd.OpenAndCloseJournal(driver,"open");
                            string stan = wd.ParseJournal(driver);
                            testsheet.Cells[7][i] = stan;
                            logger.Info("STAN : " + stan);
                            test1.Info("STAN : " + stan);
                            Thread.Sleep(1000);
                            wd.OpenAndCloseJournal(driver,"close");
                            if (executetype == "C")
                            {
                                wd.Click(driver, driver.FindElementByName(executetype));
                            }
                            logger.Info("Starting Database Processing");
                             db.GenerateConnectionString();
                            db.ExecuteQuery(trantype, tstname, test1, stan,cardno,expoutput, testsheet,i);



                        }
                        else
                        {
                            test1.Fail("Test Fails Because ATM Screen Stuck");
                        }


                            Thread.Sleep(2000);
                        


                        if (i == lastUsedRow)
                        {
                            logger.Info("Test Cases Finished : ");
                        }
                    }


                   

                }
                catch (Exception e) {
                    logger.Info("Exception occur while performing Transaction : "+ e.Message);
                }
                
            }

            // outer for loop End

            logger.Info("Activity Finished : ");

            wd.SaveSheet(wb1,testfilelocation);

            wb.Close();
            wb1.Close();
            extent.Flush();
            
            driver.Dispose();
          
            

        }

   


    }
}
