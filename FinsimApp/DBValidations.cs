using System;
using System.Configuration;
using Oracle.ManagedDataAccess.Client;
using AventStack.ExtentReports;
using System.Threading;
using excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;
using NLog;

namespace FinsimApp
{
    class DBValidations
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

       // String hp = String.Empty;
        string columnval = string.Empty;
        string columnname = string.Empty;
        string dbcon = string.Empty;
        int z = 0;
        OracleConnection conn;
        string dbip = string.Empty;
        string dBPort = string.Empty;
        string dataSource = string.Empty;
        string dbUserId = string.Empty;
        string dbPassword = string.Empty;
         
       public void GenerateConData(excel.Range rng1)
        {

            dbip = rng1.Cells[4][2].value2;
            dBPort = rng1.Cells[5][2].value2;
            dataSource = rng1.Cells[6][2].value2;
            dbUserId = rng1.Cells[7][2].value2;
            dbPassword = rng1.Cells[8][2].value2;

        }
        

        public void GenerateConnectionString()
        {
            logger.Info("DB credentials : " + dbip + " " + dBPort + " " + dataSource + " " + dbUserId + " " + dbPassword);
            dbcon = "Data Source = (DESCRIPTION = " + "(ADDRESS = (PROTOCOL = TCP)(HOST = " + dbip + ")(PORT = " + dBPort + "))" +
                                 "(CONNECT_DATA = " +
                                 "(SERVER = DEDICATED)" +
                                 "(SERVICE_NAME = " + dataSource + " )" +
                                 ")" +
                                 ");User Id = " + dbUserId + "; Password = " + dbPassword + ";";
            logger.Info("Connection String : "+dbcon);

            conn = new OracleConnection(dbcon);
            


        }
      
        
        // OracleConnection conn = new OracleConnection(ConfigurationManager.ConnectionStrings["con"].ConnectionString);
        public void OpenConnection()
        {
     
            try
            {
                 conn.Open();
            }
            catch (Exception e) {
                logger.Info("Issue while Connecting Database : "+e);
            }
            
        }

        public void CloseConnection()
        {
            conn.Close();
        }

        //public void QSelect(string trantype, string tstname, string stan) {
        //    if (trantype == "ON-US")
        //    {
        //        if (tstname.Contains("Balance"))
        //        {
        //            hp = $"select tlog.*, errcd.ERROR_NAME as TRAN_STATUS from TBLTRANSACTIONLOG tlog left join TBLCFGERRORCODES errcd on tlog.RESP_CODE = errcd.ERROR_CODE where tlog.SYS_TRACE_AUDIT_NO = '{stan}'";
        //        }
        //        else
        //        {
        //            hp = $"select tlog.*,cust.PRODUCTID,cust.CUSTOMER_ID,clim.REMAINING_LIMIT,errcd.ERROR_NAME as TRAN_STATUS from TBLTRANSACTIONLOG tlog left join TBLCUSTCHANNELACCT cust on tlog.PAN = substr(trim(cust.RELATIONSHIP_ID), 1, 16) and tlog.ACQUIRING_CHANNEL_ID = cust.CHANNEL_ID left join TBLREMAININGCARDLIMIT clim on cust.RELATIONSHIP_ID = clim.RELATIONSHIP_ID left join TBLCFGLIMIT lim on lim.TRAN_CODE = tlog.TRAN_CODE and lim.CHANNEL_ID = tlog.ACQUIRING_CHANNEL_ID and cust.PRODUCTID = lim.GROUP_ID left join TBLCFGERRORCODES errcd on tlog.RESP_CODE = errcd.ERROR_CODE where SYS_TRACE_AUDIT_NO = '{ stan }' and lim.CHANNEL_ID = tlog.ACQUIRING_CHANNEL_ID and lim.AMOUNT_ID = clim.AMOUNT_ID";
        //        }
        //    }
        //    else
        //    {
        //        hp = $"select tlog.*,errcd.ERROR_NAME as TRAN_STATUS from TBLTRANSACTIONLOG tlog left join TBLCFGERRORCODES errcd on tlog.RESP_CODE = errcd.ERROR_CODE where SYS_TRACE_AUDIT_NO = '{stan}'";
        //    }


        //}

        public void ExecuteQuery(string trantype,string tstname,ExtentTest test1,string stan,string cardno,string expoutput,excel.Worksheet testsheet,int i) {
            try {
                List<string> chkFields = FieldsToCheck();
                chkFields.ForEach(p => logger.Info(p));
                OpenConnection();
                //QSelect(trantype, tstname, stan);
                //OracleCommand cmd1 = new OracleCommand(hp, conn);
                //OracleDataReader rd1 = cmd1.ExecuteReader();
                OracleCommand cmd1 = conn.CreateCommand();
                if (trantype == "ON-US") {
                    logger.Info("Running For ON-US transaction");
                    cmd1.CommandText = "PKGAUTOMATION.spGetATMTranDetailON_US";
                    cmd1.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd1.Parameters.Add("inStan", OracleDbType.Varchar2).Value = stan;
                    cmd1.Parameters.Add("inTranType", OracleDbType.Varchar2).Value = tstname;
                    cmd1.Parameters.Add("outCursor", OracleDbType.RefCursor).Direction = System.Data.ParameterDirection.Output;
                }
                else
                {
                    cmd1.CommandText = "PKGAUTOMATION.spGetATMTranDetailOFF_US";
                    cmd1.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd1.Parameters.Add("inStan", OracleDbType.Varchar2).Value = stan;
                    cmd1.Parameters.Add("outCursor", OracleDbType.RefCursor).Direction = System.Data.ParameterDirection.Output;
                }
                 
                OracleDataReader rd1 = cmd1.ExecuteReader();
                while (rd1.Read())
                {

                    for (z = 0; z < rd1.FieldCount; z++)
                    {
                        columnname = rd1.GetName(z);
                        if (rd1.GetValue(z) is null || rd1.GetValue(z).ToString() == "")
                        {
                            columnval = "[       ]";
                        }
                        else
                        {
                            columnval = rd1.GetString(z);
                        }

                        if (z % 2 != 0)
                        {
                            Console.Write(columnname + " : " + columnval + "            ");

                        }
                        else
                        {
                            logger.Info(columnname + " : " + columnval);
                        }
                       // columnname == "RESP_CODE"
                        if (chkFields.Contains(columnname) )
                        {
                            if (columnval != "[       ]") {
                                test1.Info(columnname + " : " + columnval);
                            }
                            else
                            {
                                test1.Info("<h7 style='color: red;'>" + columnname + " : " + columnval+"</h7>");
                            }

                            
                        }
                  
                        if (columnname == "TRAN_STATUS")
                        {
                            if (columnval.Equals(expoutput))
                            {
                                testsheet.Cells[5][i] = columnval;
                                testsheet.Cells[8][i] = "PASS";
                                test1.Pass("Test case Pass Reason : " + columnval);
                            }
                            else if(columnval.Equals("ERR_FREQ_EXCEEDED"))
                            {
                                testsheet.Cells[5][i] = columnval;
                                testsheet.Cells[8][i] = "FAIL";
                                test1.Fail("Test case Fail Reason : " + columnval);
                                UPDATEFREQUENCY(cardno);
                            }
                            else
                            {
                                testsheet.Cells[5][i] = columnval;
                                testsheet.Cells[8][i] = "FAIL";
                                test1.Fail("Test case Fail Reason : " + columnval);
                            }
                        }


                    }



                }
                CloseConnection();
                Thread.Sleep(1000);
            }
            catch (Exception e) {
                logger.Info("EXCEPTION OCCUR : "+e.Message);
            }
          
        }

        public void UPDATEOTP(string cardno)
        {
            
            OpenConnection();
            OracleCommand cmd1 = conn.CreateCommand();
            cmd1.CommandText = "PKGAUTOMATION.spUpdateOTP";
            cmd1.CommandType = System.Data.CommandType.StoredProcedure;
            cmd1.Parameters.Add("inCardNum", OracleDbType.Varchar2).Value = cardno;
            cmd1.Parameters.Add("inChannelID", OracleDbType.Varchar2).Value = "0001";
            //hp = $"update TBLOTP set OTP='FB13D6CC7ED7B16E77C4B0F42256D284' where RELATIONSHIP_ID like'{cardno}%'";
            //OracleCommand cmd1 = new OracleCommand(hp, conn);
            OracleDataReader rd1 = cmd1.ExecuteReader();
            logger.Info("OTP Updated : ");
            //hp = $"commit";
            //OracleCommand cmd2 = new OracleCommand(hp, conn);
            //OracleDataReader rd2 = cmd1.ExecuteReader();
            CloseConnection();
        }
        public void UPDATEFREQUENCY(string cardno)
        {

           
            OracleCommand cmd1 = conn.CreateCommand();
            cmd1.CommandText = "PKGAUTOMATION.spUpdateRemainingFrequency";
            cmd1.CommandType = System.Data.CommandType.StoredProcedure;
            cmd1.Parameters.Add("inCardNum", OracleDbType.Varchar2).Value = cardno;
            OracleDataReader rd1 = cmd1.ExecuteReader();
            logger.Info("Frequency Updated : ");
           
        }

        public List<string> FieldsToCheck()
        {
          string dirpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            string filepath = dirpath + "\\FilesToCheckFields\\Withdrawl.txt";
            List<string> lines = System.IO.File.ReadLines(filepath).ToList();
            return lines;
        
        }





    }
}
