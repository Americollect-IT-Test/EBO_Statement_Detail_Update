using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EBO_Statement_Detail_Update
{
    class Program
    {
        static void Main(string[] args)
        {
            String UserName = Environment.GetEnvironmentVariable("USERNAME").ToUpper();

            AMC_Functions.GeneralFunctions oGenFun = new AMC_Functions.GeneralFunctions();

            bool TestMode = true;

            DateTime StartTime = DateTime.Now;

            DateTime EndTime;


            if (UserName == "MACRO" || UserName == "ADMINISTRATOR")
            {
                try
                {
                    // run all clients, so query for all
                    EBO_Statement_Detail_Update oStatementUpdate = new EBO_Statement_Detail_Update(TestMode);

                    // after completion, store the current time, to pass through for the write log function
                    EndTime = DateTime.Now;

                    // write the log file, as long as it gets through
                    oGenFun.WriteLogFile(true, StartTime, EndTime, false, "jerrodr");
                }
                catch (Exception Ex)
                {
                    // after completion, store the current time, to pass through for the write log function
                    EndTime = DateTime.Now;

                    oGenFun.WriteLogFile(false, StartTime, EndTime, true, "jerrodr", Ex);
                }
                
            }
            else
            {
                // called by a person, so need to get the arguments if any, and pass those through, otherwise query for all
                if (args.Length >= 1)
                {
                    EBO_Statement_Detail_Update oStatementUpdate = new EBO_Statement_Detail_Update(TestMode, args[0]);
                }
                else
                {
                    EBO_Statement_Detail_Update oStatementUpdate = new EBO_Statement_Detail_Update(TestMode);
                }
            }
        }
    }
}
