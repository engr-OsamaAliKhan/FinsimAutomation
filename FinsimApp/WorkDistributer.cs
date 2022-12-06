using AutoIt;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Winium;
using System.Net.Sockets;
using System.Net;
using AventStack.ExtentReports;
using System.Diagnostics;
using NLog;
using System.IO;

namespace FinsimApp
{
    class WorkDistributer
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        string LastSTANRRN = string.Empty;


        public void TearDown(WiniumDriver driver,string DeviceName, ExtentReports extent)
        {

            AutoItX.WinClose("ATM Simulator - " + DeviceName, "");
            if (driver != null)
                driver.Close();

            extent.Flush();
           
            Thread.Sleep(1000);
        }


        //Sending Keys
        public void SendKeys(string Keys)
        {
            AutoItX.Send(Keys);
            Thread.Sleep(100);

        }
        //Open and Close Journal

        public void OpenAndCloseJournal(WiniumDriver driver,string command)
        {
            bool jrnl = false;
            logger.Info("Opening Journal");
             driver.FindElementByName("Windows").Click();
            
            if (command.ToLower() == "open")
            {
                try
                {

                    if (jrnl)

                    {
                        driver.FindElementByName("Windows").Click();


                    }
                    else
                    {
                        SendKeys("j");

                    }
                }
                catch
                {
                    logger.Info("Element not found in given time while opening journal: 1");
                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3.00);
                }

                Thread.Sleep(1000);

            }
            try
            {
                if (driver.FindElementById("3801").Displayed==true) {
                    jrnl = true;
                }
                var element = driver.FindElementById("3801");
                

            }
            catch
            {
                logger.Info("Element not found in given time while opening journal: 2");
            }
            if (command.ToLower() == "close")
            {
                try
                {
                    if (jrnl)
                    //if (driver.FindElementByName("Journal").Selected == true)
                    {
                        SendKeys("j");

                    }
                    else
                    {
                        driver.FindElement(By.Name("Windows")).Click();


                    }
                }
                catch
                {
                    logger.Info("Element not found in given time while opening journal: 3");
                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10.00);
                }
            }

        }

        //Parsing Journal
        public string ParseJournal(WiniumDriver driver)
        {
            String stan = null;
            //¤-------------------------
            //-----------------------------------
            // ¤-------------------------
            //¤-------------------------
            string[] journaltext = driver.FindElementById("3801").Text.Split(new string[] { "-----------------------------------", "¤-------------------------", "-------------------------", "¤" }, StringSplitOptions.RemoveEmptyEntries);
            string currentTrxn = journaltext[journaltext.Length - 1];

            if (currentTrxn.Contains("STATUS"))
            {
                string[] currentTrxnParts = currentTrxn.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var v in currentTrxnParts)
                {
                    if (v.Contains("STAN") && v != LastSTANRRN)
                    {

                        LastSTANRRN = v;
                        string[] stanRRN = LastSTANRRN.Split(new string[] { " ", ":" }, StringSplitOptions.RemoveEmptyEntries);
                        String rrn = stanRRN[stanRRN.Length - 1];
                        
                        stan = stanRRN[stanRRN.Length - 3];

                    }
                    else if (v == LastSTANRRN)
                    {
                        logger.Info("Transaction Failed", "No New Transaction found for this Transaction");
                    }

                }
            }
            return stan;
        }



        //Clear text from Textbox
        public void ClearTextByElement(IWebElement element)
        {
            SendKeys("{END}");
            int existingCount = element.Text.Length;
            for (int i = 0; i < existingCount; i++)
            { //element.Clear();
                TextClear();
            }
        }
        public void TextClear()
        {

            AutoItX.Send("{BACKSPACE}");
            Thread.Sleep(20);

        }
        public void VerifyTextonTextBox(IWebElement element, string Text)
        {


            if (element.Text != Text || element.Text == null || element.Text == "")
            {
                ClearTextByElement(element);
                element.SendKeys(Text);

            }
        }
        //Clicking on element

        public void Click(IWebDriver driver, IWebElement element)
        {
            //driver.Manage().Timeouts().ImplicitWait();
          // driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10.00);




            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40.00));
            wait.IgnoreExceptionTypes(new[] { typeof(StaleElementReferenceException), typeof(NoSuchElementException) });

            try
            {
                wait.Until(driver2 =>

               element.Displayed == true);
                //element.Click();
                Actions Ac = new Actions(driver);
                Ac.Click(element).Build().Perform();

            }
            catch {
                logger.Info("Element not found in given time : ");
            }
            
            Thread.Sleep(3000);
          

        }
        // Selecting from Dropdown
        public void SelectFromDropDown(string ApplicationName, string ComboBox, string Value)
        {
            AutoItX.WinActivate(ApplicationName, "");

            AutoItX.ControlCommand(ApplicationName, "", ComboBox, "SelectString", Value);

        }

        //Waithing for transaction finish
        public void WaitForTransactionToFinish(IWebDriver driver)
        {
            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120.00));
            wait.IgnoreExceptionTypes(new[] { typeof(StaleElementReferenceException), typeof(NoSuchElementException) });

            try
            {
                string name1 = driver.FindElement(By.Id("25466")).GetAttribute("Name");
                Console.WriteLine(name1);
                wait.Until(driver2 =>
                driver.FindElement(By.Id("25466")).GetAttribute("Name") == "State : 000"
                    );
                logger.Info("State 000 found");
            }
            catch(Exception e) {
                logger.Info("EXCEPTION OCCUR   "+e.Message);
                logger.Info("Transaction take more then 1 min to finish");
            }

        }

        public void WaitForMainScreen(IWebDriver driver)
        {
            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
            wait.IgnoreExceptionTypes(new[] { typeof(StaleElementReferenceException), typeof(NoSuchElementException) });

            try
            {
                wait.Until(driver2 =>
                driver.FindElement(By.Id("25466")).GetAttribute("Name") == "State : 930"
                    );
                logger.Info("Main screen Reached");
            }
            catch (Exception e)
            {
                logger.Info("EXCEPTION OCCUR   " + e.Message);
            }

        }


        public void WaitForIdleState(IWebDriver driver)
        {
            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50.00));
            wait.IgnoreExceptionTypes(new[] { typeof(StaleElementReferenceException), typeof(NoSuchElementException) });

            try
            {
                wait.Until(driver2 =>
                driver.FindElement(By.Id("25466")).GetAttribute("Name") == "Screen : 010"
                    );
            }
            catch (Exception e)
            {
                logger.Info("EXCEPTION OCCUR   " + e.Message);
            }

        }





        public int ExcelRows(excel.Range rng3)
        {
            int lastUsedRow = rng3.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            return lastUsedRow;
        }

        public int ExcelColumns(excel.Range rng3)
        {
            int lastUsedColumn = rng3.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            return lastUsedColumn;
        }

        public static void TakeScreenshot(WiniumDriver driver,string testname)
        {
            string path1 = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug", "");
            string path = path1 + "Screenshot\\" + "ScreenShot_"+ testname +DateTime.Now.ToString("yyyyMMddHHmmss")+ ".png";
            ITakesScreenshot ts =(ITakesScreenshot)driver;
            Screenshot sct=ts.GetScreenshot();
            sct.SaveAsFile(path, ScreenshotImageFormat.Png);
            
        }

        public static int FreeTcpPort()
        {
            TcpListener l = new TcpListener(IPAddress.Loopback, 0);
            l.Start();
            int port = ((IPEndPoint)l.LocalEndpoint).Port;
            l.Stop();
            return port;
        }

        public void VerifyCheckBoxSelected(IWebElement element, bool select)
        {

            bool checkboxselect = element.Selected;

            if (checkboxselect != select)
                element.Click();

        }

        public bool WaitToDisplayElement(WiniumDriver driver, string key1)
        {
            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10.00));
            wait.IgnoreExceptionTypes(new[] { typeof(StaleElementReferenceException), typeof(NoSuchElementException) });
            
            try
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                //driver.Manage().Timeouts().ImplicitWait(TimeSpan.FromSeconds(5));
                Console.WriteLine("Start Waiting : ");
                //wait.Until(driver2 =>
                //driver.FindElementByName(key1).Displayed == true
                //    );
                Thread.Sleep(5000);
                if (driver.FindElementByName(key1).Displayed == true)
                {
                    logger.Info("Element Displayed");
                    return true;
                }
                else
                {
                    logger.Info("Element Not Found : ");
                    return false;
                }
                //return true;
            }
            catch (NoSuchElementException e)
            {
                logger.Info("Element Not Found : ");
                return false;
            }


        }

        public void SaveSheet(excel.Workbook wb1,string LocationAsURL)
        {
            if (DoesFileExistInDirectory(LocationAsURL))
            {
                wb1.Save();
            }
            else
            {
                wb1.SaveAs(LocationAsURL);
            }
        }

        private bool DoesFileExistInDirectory(string LocationAsURL)
        {
            return File.Exists(LocationAsURL);
        }







    }
}
