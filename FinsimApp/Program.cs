using AutoIt;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Winium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Configuration;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Winium.Elements.Desktop;
using Winium.Elements.Desktop.Extensions;
using excel = Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using System.Reflection;

namespace FinsimApp
{
    class Program
    {
        static void Main(string[] args)
        {
            TestCasesRunner tr = new TestCasesRunner();
            tr.TestRunner();
        }
    }
}
