using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;
using ClosedXML.Excel;

namespace EBO_Statement_Detail_Update
{
    class EBO_Statement_Detail_Update
    {
        double DaysToLookBack = 0;      // make sure this value is positive

        string connectionString;
        string connectionString_RO;

        DataTable dt_allData = new DataTable();
        DataTable dt_ExportData = new DataTable("ExportedData");

        DataTable dt_UpdateData = new DataTable("DateToUpdate");

        const string ChangePerson = "0102";
        const string StatementGroup = "STMT-G";

        string connectionString_SQL = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=amc-mysql.americollect.com;DATABASE=EBO_Specialist_Tool;UID=EBO_Special_Tool;PWD=N3Hu7yFGGO97p51F; OPTION=3;Connect Timeout=50;";

        public EBO_Statement_Detail_Update(bool inTestMode,  string inCredGrp = "")
        {
            AMC_Functions.DetermineDSNFile oDSN = new AMC_Functions.DetermineDSNFile();


            connectionString = inTestMode ? oDSN.getDSNFile("jerrodr", "Training DB5", true) : oDSN.getDSNFile("jerrodr", "DB5", true);
            connectionString_RO = inTestMode ? oDSN.getDSNFile("jerrodr", "Training DB5", false) : oDSN.getDSNFile("jerrodr", "DB5", false);

            using (OdbcConnection con = new OdbcConnection(connectionString_RO))
            {
                con.Open();

                QueryData(inCredGrp, con, inTestMode);
                AddInDTHeaders();
                EvaluateData(con, inTestMode);
                ExportDataIfApplicable(inCredGrp);
            }

            InsertIntoDataBase(inCredGrp);

            
        }

        /// <summary>
        /// Query for all of the data
        /// </summary>
        /// <param name="inCredGrp"></param>
        private void QueryData(string inCredGrp, OdbcConnection inCN, bool inTestMode)
        {
            string selectSQL;

            DateTime queryDate = DateTime.Today.AddDays(DaysToLookBack * -1);

            if (inTestMode)
            {
                if (inCredGrp == "")
                {
                    selectSQL = $@"SELECT amanumber, amdnumber, ambalance, amcnumber, acctwin_A.wdtext[3] as 'TotalCharges', acctwin_A.wdtext[4] as 'TotalAdjust', acctwin_A.wdtext[5] as 'TotalIns', acctwin_A.wdtext[6] as 'TotalPat',
                            '' as 'TotalPayments', 
                            '' as 'TotalInsPayments', 
                            '' as 'TotalAdjusts', acctwin_A.wdtext[1] as 'ADATA:AA1', acctwin_A.wdtext[2] as 'ADATA:AA2', acctwin_A.wdtext[3] as 'ADATA:AA3', acctwin_A.wdtext[4] as 'ADATA:AA4', acctwin_A.wdtext[5] as 'ADATA:AA5', acctwin_A.wdtext[6] as 'ADATA:AA6', acctwin_A.wdtext[7] as 'ADATA:AA7', acctwin_A.wdtext[8] as 'ADATA:AA8', acctwin_A.wdtext[9] as 'ADATA:AA9', acctwin_A.wdtext[10] as 'ADATA:AA10', acctwin_A.wdtext[11] as 'ADATA:AA11', acctwin_A.wdtext[12] as 'ADATA:AA12', acctwin_A.wdtext[13] as 'ADATA:AA13', acctwin_A.wdtext[14] as 'ADATA:AA14', acctwin_A.wdtext[15] as 'ADATA:AA15', acctwin_A.wdtext[16] as 'ADATA:AA16'
                                FROM PUB.acctmstr qA
                                LEFT JOIN PUB.windata acctwin_A on acctwin_A.wdtype = 'A' and acctwin_A.wdcode = 'A' and acctwin_A.wdnumber = amanumber and acctwin_A.wdagency = amagency
                                JOIN PUB.credgrpd on PUB.credgrpd.gdcnumber = amcnumber
                                WHERE ambalance > 0 and acctwin_A.wdtext[3] != '' and gdgnumber = '{StatementGroup}' WITH (NOLOCK)";

                    Console.WriteLine("Querying for all data to update for Statement Data..");
                }
                else
                {
                    // otherwise query for the specific cred group
                    selectSQL = $@"SELECT amanumber, amdnumber, ambalance, amcnumber, acctwin_A.wdtext[3] as 'TotalCharges', acctwin_A.wdtext[4] as 'TotalAdjust', acctwin_A.wdtext[5] as 'TotalIns', acctwin_A.wdtext[6] as 'TotalPat',
                            '' as 'TotalPayments', 
                            '' as 'TotalInsPayments', 
                            '' as 'TotalAdjusts', acctwin_A.wdtext[1] as 'ADATA:AA1', acctwin_A.wdtext[2] as 'ADATA:AA2', acctwin_A.wdtext[3] as 'ADATA:AA3', acctwin_A.wdtext[4] as 'ADATA:AA4', acctwin_A.wdtext[5] as 'ADATA:AA5', acctwin_A.wdtext[6] as 'ADATA:AA6', acctwin_A.wdtext[7] as 'ADATA:AA7', acctwin_A.wdtext[8] as 'ADATA:AA8', acctwin_A.wdtext[9] as 'ADATA:AA9', acctwin_A.wdtext[10] as 'ADATA:AA10', acctwin_A.wdtext[11] as 'ADATA:AA11', acctwin_A.wdtext[12] as 'ADATA:AA12', acctwin_A.wdtext[13] as 'ADATA:AA13', acctwin_A.wdtext[14] as 'ADATA:AA14', acctwin_A.wdtext[15] as 'ADATA:AA15', acctwin_A.wdtext[16] as 'ADATA:AA16'  
                                FROM PUB.acctmstr qA
                                LEFT JOIN PUB.windata acctwin_A on acctwin_A.wdtype = 'A' and acctwin_A.wdcode = 'A' and acctwin_A.wdnumber = amanumber and acctwin_A.wdagency = amagency
                                JOIN PUB.credgrpd on PUB.credgrpd.gdcnumber = amcnumber
                            WHERE ambalance > 0 and gdgnumber = '{inCredGrp}' WITH (NOLOCK)";

                    Console.WriteLine($"Querying for all data for {inCredGrp} to update for Statement Data for..");
                }

                using (OdbcCommand SelectCMD = new OdbcCommand(selectSQL, inCN))
                using (OdbcDataAdapter adapter = new OdbcDataAdapter(SelectCMD))
                {
                    adapter.Fill(dt_allData);
                }

                Console.WriteLine("Querying for the transactions for LIVE currently to evaluate for testing.");

                // now, since we're testing, need to change over to the live connection and get the transactions for today, so we can get an accurate export of the data that needs to be updated
                AMC_Functions.DetermineDSNFile oDSN = new AMC_Functions.DetermineDSNFile();

                string connectionString = oDSN.getDSNFile("jerrodr", "DB5", false);

                //(select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmrcptcode NOT IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C') as 'TotalPAyments'
                //(select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmrcptcode IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C') as 'TotalInsPayments'
                //(select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'A') as 'TotalAdjusts'
                using (OdbcConnection ConLive = new OdbcConnection(connectionString))
                {
                    string select_TotalPay = $@"SELECT sum(baamount) as 'TotalPayments' 
                                            from PUB.tranmstr 
                                            JOIN PUB.balances on PUB.balances.baserial = tmtserial 
                                            JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber 
                                            WHERE amdnumber = ? and tmrcptcode NOT IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C'";

                    string select_TotalIns = $@"SELECT sum(baamount) as 'TotalInsPayments' 
                                            from PUB.tranmstr 
                                            JOIN PUB.balances on PUB.balances.baserial = tmtserial 
                                            JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber 
                                            WHERE amdnumber = ? and tmrcptcode IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C'";

                    string select_TotalAdj = $@"SELECT sum(baamount) as 'TotalAdjusts' 
                                            from PUB.tranmstr 
                                            JOIN PUB.balances on PUB.balances.baserial = tmtserial 
                                            JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber 
                                            WHERE amdnumber = ? and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'A'";

                    ConLive.Open();

                    foreach (DataRow dr in dt_allData.Rows)
                    {
                        // total payments
                        using (OdbcCommand SelectCMD = new OdbcCommand(select_TotalPay, ConLive))
                        {
                            SelectCMD.Parameters.Add("@dbtr", OdbcType.VarChar).Value = dr["amdnumber"].ToString();

                            using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                            {
                                while (Reader.Read())
                                {
                                    dr["TotalPayments"] = Reader["TotalPayments"].ToString();
                                }
                            }
                        }

                        // total insurnace payments
                        using (OdbcCommand SelectCMD = new OdbcCommand(select_TotalIns, ConLive))
                        {
                            SelectCMD.Parameters.Add("@dbtr", OdbcType.VarChar).Value = dr["amdnumber"].ToString();

                            using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                            {
                                while (Reader.Read())
                                {
                                    dr["TotalInsPayments"] = Reader["TotalInsPayments"].ToString();
                                }
                            }
                        }

                        // total adjustments
                        using (OdbcCommand SelectCMD = new OdbcCommand(select_TotalAdj, ConLive))
                        {
                            SelectCMD.Parameters.Add("@dbtr", OdbcType.VarChar).Value = dr["amdnumber"].ToString();

                            using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                            {
                                while (Reader.Read())
                                {
                                    dr["TotalAdjusts"] = Reader["TotalAdjusts"].ToString();
                                }
                            }
                        }

                    }
                }
            }
            else 
            {
                if (inCredGrp == "")
                {
                    selectSQL = $@"SELECT amanumber, ambalance, amcnumber, acctwin_A.wdtext[3] as 'TotalCharges', acctwin_A.wdtext[4] as 'TotalAdjust', acctwin_A.wdtext[5] as 'TotalIns', acctwin_A.wdtext[6] as 'TotalPat',
                            (select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmrcptcode NOT IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C') as 'TotalPayments', 
                            (select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmrcptcode IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C') as 'TotalInsPayments', 
                            (select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'A') as 'TotalAdjusts', acctwin_A.wdtext[1] as 'ADATA:AA1', acctwin_A.wdtext[2] as 'ADATA:AA2', acctwin_A.wdtext[3] as 'ADATA:AA3', acctwin_A.wdtext[4] as 'ADATA:AA4', acctwin_A.wdtext[5] as 'ADATA:AA5', acctwin_A.wdtext[6] as 'ADATA:AA6', acctwin_A.wdtext[7] as 'ADATA:AA7', acctwin_A.wdtext[8] as 'ADATA:AA8', acctwin_A.wdtext[9] as 'ADATA:AA9', acctwin_A.wdtext[10] as 'ADATA:AA10', acctwin_A.wdtext[11] as 'ADATA:AA11', acctwin_A.wdtext[12] as 'ADATA:AA12', acctwin_A.wdtext[13] as 'ADATA:AA13', acctwin_A.wdtext[14] as 'ADATA:AA14', acctwin_A.wdtext[15] as 'ADATA:AA15', acctwin_A.wdtext[16] as 'ADATA:AA16'
                                FROM PUB.acctmstr qA
                                LEFT JOIN PUB.windata acctwin_A on acctwin_A.wdtype = 'A' and acctwin_A.wdcode = 'A' and acctwin_A.wdnumber = amanumber and acctwin_A.wdagency = amagency
                                JOIN PUB.credgrpd on PUB.credgrpd.gdcnumber = amcnumber
                                WHERE ambalance > 0 and acctwin_A.wdtext[3] != '' and gdgnumber = '{StatementGroup}' WITH (NOLOCK)";

                    Console.WriteLine("Querying for all data to update for Statement Data..");
                }
                else
                {
                    // otherwise query for the specific cred group
                    selectSQL = $@"SELECT amanumber, ambalance, amcnumber, acctwin_A.wdtext[3] as 'TotalCharges', acctwin_A.wdtext[4] as 'TotalAdjust', acctwin_A.wdtext[5] as 'TotalIns', acctwin_A.wdtext[6] as 'TotalPat',
                            (select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmrcptcode NOT IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C') as 'TotalPayments', 
                            (select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmrcptcode IN ('I', 'Y') and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'C') as 'TotalInsPayments', 
                            (select sum(baamount) from PUB.tranmstr JOIN PUB.balances on PUB.balances.baserial = tmtserial JOIN PUB.acctmstr on PUB.acctmstr.amanumber = tmanumber WHERE amdnumber = qA.amdnumber and tmtrandate >= '{queryDate.ToString("yyyy-MM-dd")}' and tmtrancode = 'A') as 'TotalAdjusts', acctwin_A.wdtext[1] as 'ADATA:AA1', acctwin_A.wdtext[2] as 'ADATA:AA2', acctwin_A.wdtext[3] as 'ADATA:AA3', acctwin_A.wdtext[4] as 'ADATA:AA4', acctwin_A.wdtext[5] as 'ADATA:AA5', acctwin_A.wdtext[6] as 'ADATA:AA6', acctwin_A.wdtext[7] as 'ADATA:AA7', acctwin_A.wdtext[8] as 'ADATA:AA8', acctwin_A.wdtext[9] as 'ADATA:AA9', acctwin_A.wdtext[10] as 'ADATA:AA10', acctwin_A.wdtext[11] as 'ADATA:AA11', acctwin_A.wdtext[12] as 'ADATA:AA12', acctwin_A.wdtext[13] as 'ADATA:AA13', acctwin_A.wdtext[14] as 'ADATA:AA14', acctwin_A.wdtext[15] as 'ADATA:AA15', acctwin_A.wdtext[16] as 'ADATA:AA16'  
                                FROM PUB.acctmstr qA
                                LEFT JOIN PUB.windata acctwin_A on acctwin_A.wdtype = 'A' and acctwin_A.wdcode = 'A' and acctwin_A.wdnumber = amanumber and acctwin_A.wdagency = amagency
                                JOIN PUB.credgrpd on PUB.credgrpd.gdcnumber = amcnumber
                            WHERE ambalance > 0 and gdgnumber = '{inCredGrp}' WITH (NOLOCK)";

                    Console.WriteLine($"Querying for all data for {inCredGrp} to update for Statement Data for..");
                }

                using (OdbcCommand SelectCMD = new OdbcCommand(selectSQL, inCN))
                using (OdbcDataAdapter adapter = new OdbcDataAdapter(SelectCMD))
                {
                    adapter.Fill(dt_allData);
                }
            }
         }

        /// <summary>
        ///  need to add in the headers for the exported data table
        /// </summary>
        private void AddInDTHeaders()
        {
            dt_ExportData.Columns.Add("Account Number");
            dt_ExportData.Columns.Add("Current Balance");
            dt_ExportData.Columns.Add("Creditor Code");
            dt_ExportData.Columns.Add("Total Charges");
            dt_ExportData.Columns.Add("Total Adjustments");
            dt_ExportData.Columns.Add("Total Patient Payments");
            dt_ExportData.Columns.Add("Total Insurance Payments");
            dt_ExportData.Columns.Add("OFf By");


            dt_UpdateData.Columns.Add("amanumber");
            dt_UpdateData.Columns.Add("ambalance");
            dt_UpdateData.Columns.Add("amcnumber");
            dt_UpdateData.Columns.Add("TotalCharges (ADATA:AA3");
            dt_UpdateData.Columns.Add("TotalAdjust (ADATA:AA4)");
            dt_UpdateData.Columns.Add("TotalIns (ADATA:AA5)");
            dt_UpdateData.Columns.Add("TotalPat (ADATA:AA6)");
            dt_UpdateData.Columns.Add("TotalPayments");
            dt_UpdateData.Columns.Add("TotalInsPayments");
            dt_UpdateData.Columns.Add("TotalAdjusts");
            dt_UpdateData.Columns.Add("ADATA:AA1");
            dt_UpdateData.Columns.Add("ADATA:AA2");
            dt_UpdateData.Columns.Add("ADATA:AA3");
            dt_UpdateData.Columns.Add("NEW ADATA:AA4");
            dt_UpdateData.Columns.Add("NEW ADATA:AA5");
            dt_UpdateData.Columns.Add("NEW ADATA:AA6");
            dt_UpdateData.Columns.Add("ADATA:AA7");
            dt_UpdateData.Columns.Add("ADATA:AA8");
            dt_UpdateData.Columns.Add("ADATA:AA9");
            dt_UpdateData.Columns.Add("ADATA:AA10");
            dt_UpdateData.Columns.Add("ADATA:AA11");
            dt_UpdateData.Columns.Add("ADATA:AA12");
            dt_UpdateData.Columns.Add("ADATA:AA13");
            dt_UpdateData.Columns.Add("ADATA:AA14");
            dt_UpdateData.Columns.Add("ADATA:AA15");
            dt_UpdateData.Columns.Add("ADATA:AA16");
        }

        /// <summary>
        /// Check the data that came from the query and see which ones already match, and which can be updated/need to be manually evaluated
        /// </summary>
        private void EvaluateData(OdbcConnection inCN, bool inTestMode)
        {
            foreach (DataRow dr in dt_allData.Rows)
            {
                // make variables for each, just in case they aren't parseable can treat them as zeros
                decimal CurrentBalance = dr["ambalance"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["ambalance"].ToString().Trim().Replace("$", ""));
                decimal TotalCharges = dr["TotalCharges"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalCharges"].ToString().Trim().Replace("$", ""));
                decimal TotalPatPayments = dr["TotalPat"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalPat"].ToString().Trim().Replace("$", ""));
                decimal TotalInsPayments = dr["TotalIns"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalIns"].ToString().Trim().Replace("$", ""));
                decimal TotalAdjustments = dr["TotalAdjust"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalAdjust"].ToString().Trim().Replace("$", ""));

                decimal TotalAdjInPeriod = dr["TotalAdjusts"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalAdjusts"].ToString().Trim().Replace("$", ""));
                decimal TotalInsPayInPeriod = dr["TotalInsPayments"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalInsPayments"].ToString().Trim().Replace("$", ""));
                decimal TotalPatPayInPeriod = dr["TotalPayments"].ToString().Trim() == string.Empty ? 0 : Convert.ToDecimal(dr["TotalPayments"].ToString().Trim().Replace("$", ""));

                // check if they already match, if so, then we can ignore it
                if (Math.Round(CurrentBalance,2) != Math.Round((TotalCharges - TotalPatPayments - TotalInsPayments - TotalAdjustments),2))
                {
                    // check if adding the totals for the last time frame added together end up making them match
                    if (Math.Round(CurrentBalance, 2) == Math.Round(((TotalCharges - TotalPatPayments - TotalInsPayments - TotalAdjustments) + TotalAdjInPeriod + TotalInsPayInPeriod + TotalPatPayInPeriod), 2))
                    {
                        // update the data, need to make an array of the new data
                        string[] newWindowData = new string[16];

                        newWindowData[0] = dr["ADATA:AA1"].ToString(); newWindowData[1] = dr["ADATA:AA2"].ToString(); newWindowData[2] = dr["ADATA:AA3"].ToString(); newWindowData[3] = Decimal.Add(TotalAdjustments, (TotalAdjInPeriod * -1)).ToString();
                        newWindowData[4] = Decimal.Add(TotalInsPayments, (TotalInsPayInPeriod * -1)).ToString(); newWindowData[5] = Decimal.Add(TotalPatPayments, (TotalPatPayInPeriod * -1)).ToString(); newWindowData[6] = dr["ADATA:AA7"].ToString(); newWindowData[7] = dr["ADATA:AA8"].ToString();
                        newWindowData[8] = dr["ADATA:AA9"].ToString(); newWindowData[9] = dr["ADATA:AA10"].ToString(); newWindowData[10] = dr["ADATA:AA11"].ToString(); newWindowData[11] = dr["ADATA:AA12"].ToString();
                        newWindowData[12] = dr["ADATA:AA13"].ToString(); newWindowData[13] = dr["ADATA:AA14"].ToString(); newWindowData[14] = dr["ADATA:AA15"].ToString(); newWindowData[15] = dr["ADATA:AA16"].ToString();

                        if (inTestMode)
                        {
                            string newADATA_AA4 = Decimal.Add(TotalAdjustments, (TotalAdjInPeriod * -1)).ToString();
                            string newADATA_AA5 = Decimal.Add(TotalInsPayments, (TotalInsPayInPeriod * -1)).ToString();
                            string newADATA_AA6 = Decimal.Add(TotalPatPayments, (TotalPatPayInPeriod * -1)).ToString();

                            // test mode, so want to export to a table instead so can monitor the details
                            dt_UpdateData.Rows.Add(dr["amanumber"].ToString(), dr["ambalance"].ToString(), dr["amcnumber"].ToString(), dr["TotalCharges"].ToString().Trim(), dr["TotalAdjust"].ToString().Trim(), dr["TotalIns"].ToString().Trim(), dr["TotalPat"].ToString().Trim(), TotalPatPayInPeriod, TotalInsPayInPeriod, TotalAdjInPeriod, dr["ADATA:AA1"].ToString(), dr["ADATA:AA2"].ToString(), dr["ADATA:AA3"].ToString(), newADATA_AA4, newADATA_AA5, newADATA_AA6, dr["ADATA:AA7"].ToString(), dr["ADATA:AA8"].ToString(), dr["ADATA:AA9"].ToString(), dr["ADATA:AA10"].ToString(), dr["ADATA:AA11"].ToString(), dr["ADATA:AA12"].ToString(), dr["ADATA:AA13"].ToString(), dr["ADATA:AA14"].ToString(), dr["ADATA:AA15"].ToString(), dr["ADATA:AA16"].ToString());
                        }
                        else
                        {
                            AMC_Functions.UpdateBHWindows oBHUpdate = new AMC_Functions.UpdateBHWindows(dr["amanumber"].ToString(), "AW", "A", inCN, newWindowData, ChangePerson);
                        }

                        
                    }
                    else
                    {
                        // export to manual, since it needs to be reviewed
                        dt_ExportData.Rows.Add(dr["amanumber"].ToString(), dr["ambalance"].ToString(), dr["amcnumber"].ToString(), dr["TotalCharges"].ToString(), dr["TotalAdjust"].ToString(), dr["TotalPat"].ToString(), dr["TotalIns"].ToString(), Math.Round((TotalCharges - TotalPatPayments - TotalInsPayments - TotalAdjustments), 2));
                    }
                }
                // no else, don't need anything if they do match
            }
        }

        /// <summary>
        /// if any data is in the export data table, then export it, otherwise leave it
        /// </summary>
        private void ExportDataIfApplicable(string inCredGrp)
        {
            Console.WriteLine("Exporting data if applicable.");

            if (dt_ExportData.Rows.Count >= 1)
            {
                string fileName;

                if (inCredGrp == string.Empty)
                {
                    fileName = DateTime.Now.ToString("yyyyMMdd_HH.mm.ss") + "_" + ChangePerson + "_ALL_Statement_ADATA_Mismatch.xlsx";
                }
                else
                {
                    fileName = DateTime.Now.ToString("yyyyMMdd_HH.mm.ss") + "_" + ChangePerson + "_" + inCredGrp + "_Statement_ADATA_Mismatch.xlsx";
                }

                XLWorkbook excelExport = new XLWorkbook();
                excelExport.Worksheets.Add(dt_ExportData);
                excelExport.SaveAs(@"G:\Instructions\Visual_Studio\Jerrod\Exports\EBO Statement Data Manuals\" + fileName);
            }

            if (dt_UpdateData.Rows.Count >= 1)
            {
                string fileName;

                if (inCredGrp == string.Empty)
                {
                    fileName = DateTime.Now.ToString("yyyyMMdd_HH.mm.ss") + "_" + ChangePerson + "_ALL_Statement_ADATA_Updates.xlsx";
                }
                else
                {
                    fileName = DateTime.Now.ToString("yyyyMMdd_HH.mm.ss") + "_" + ChangePerson + "_" + inCredGrp + "_Statement_ADATA_Updates.xlsx";
                }

                XLWorkbook excelExport = new XLWorkbook();
                excelExport.Worksheets.Add(dt_ExportData);
                excelExport.SaveAs(@"G:\Instructions\Visual_Studio\Jerrod\Exports\EBO Statement Data Manuals\" + fileName);

                AMC_Functions.GeneralFunctions oGenFun = new AMC_Functions.GeneralFunctions();

                oGenFun.SendEmail("jerrodr@americollect.com", "", "Statement Updater - EBO - Complete", "Please review the data stored here: " + fileName, "macro@americollect.com", "", false, true);
            }


        }

        /// <summary>
        /// Insert the data into a database, so it can be easily accessible from the statement/final cancel process
        /// </summary>
        /// <param name="inCredGrp"></param>
        private void InsertIntoDataBase(string inCredGrp)
        {
            string CreditorGroup = inCredGrp == "" ? "All_Cred" : inCredGrp;

            

            string selectSQL = "SELECT * from `Final_Cancel_Statements` WHERE `CreditorGroup` = ?";
            string updateSQL = "UPDATE `Final_Cancel_Statements` SET `TotalManuals` = ?, `DateTimeRan` = ? WHERE `CreditorGroup` = ? ";
            string insertSQL = "INSERT INTO `Final_Cancel_Statements` (`CreditorGroup`, `TotalManuals`, `DateTimeRan`) VALUES (?, ?, ?)";

            using (OdbcConnection con_SQL = new OdbcConnection(connectionString_SQL))
            {
                con_SQL.Open();

                using (OdbcCommand SelectCMD = new OdbcCommand(selectSQL, con_SQL))
                {
                    SelectCMD.Parameters.Add("@credGrp", OdbcType.VarChar).Value = CreditorGroup;

                    using (OdbcDataReader Reader = SelectCMD.ExecuteReader())
                    {
                        if (Reader.HasRows)
                        {
                            using (OdbcCommand UpdateCMD = new OdbcCommand(updateSQL, con_SQL))
                            {
                                UpdateCMD.Parameters.Add("@TotalCount", OdbcType.BigInt).Value = dt_ExportData.Rows.Count;
                                UpdateCMD.Parameters.Add("@DateTimeRan", OdbcType.DateTime).Value = DateTime.Now;
                                UpdateCMD.Parameters.Add("@CredGrp", OdbcType.VarChar).Value = CreditorGroup;

                                UpdateCMD.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            using (OdbcCommand InsertCMD = new OdbcCommand(insertSQL, con_SQL))
                            {
                                InsertCMD.Parameters.Add("@CredGrp", OdbcType.VarChar).Value = CreditorGroup;
                                InsertCMD.Parameters.Add("@TotalCount", OdbcType.BigInt).Value = dt_ExportData.Rows.Count;
                                InsertCMD.Parameters.Add("@DateTimeRan", OdbcType.DateTime).Value = DateTime.Now;

                                InsertCMD.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }
    }
}
