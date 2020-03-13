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

            bool TestMode = false;

            DateTime StartTime = DateTime.Now;

            DateTime EndTime;

            

            if (UserName == "MACRO" || UserName == "ADMINISTRATOR")
            {
                // run all clients, so query for all
                EBO_Statement_Detail_Update oStatementUpdate = new EBO_Statement_Detail_Update(TestMode);
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
