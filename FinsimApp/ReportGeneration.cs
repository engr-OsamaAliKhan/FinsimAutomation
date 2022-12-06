using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinsimApp
{
    class ReportGeneration
    {
        ExtentHtmlReporter html;
        ExtentReports extent;
        ExtentTest test;
        public void StartReport(string rptpath)
        {
            html = new ExtentHtmlReporter(rptpath);
            extent = new ExtentReports();
            extent.AttachReporter(html);

        }

        public void AddInfoToReport(string title,string value)
        {
            extent.AddSystemInfo(title, value);
        }

        public void CreateTests(string testname)
        {
           test =extent.CreateTest(testname);
        }

        public void TestPass(string test_case_id, string test_details)
        {
            
            test.Log(Status.Pass, "Test Case ID: " + test_case_id + " <br>" + test_details);

        }

        public void TestFailwithException(string test_case_id, string test_fail_details)
        {
            try
            {
                if (string.IsNullOrEmpty(test_case_id))
                    test_case_id = "ex";

                test.Log(Status.Fail, "Test Case ID: " + test_case_id + " <br>" + test_fail_details);
            }
            catch { }

        }
        public void TestFail(string test_case_id, string test_fail_details)
        {
            try
            {
                if (string.IsNullOrEmpty(test_case_id))
                    test_case_id = "ex";

                test.Log(Status.Fail, "Test Case ID: " + test_case_id + " <br>" + test_fail_details);
            }
            catch { }

        }
    }
}
