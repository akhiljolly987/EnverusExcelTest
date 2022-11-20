# EnverusExcelTest
using System.Configuration;
using System.Data.OleDb;
using System.Linq;
using Dapper;

namespace Assignment.Rigcount
{
    class ExcelDataAccess
    {
        public static string RigcountFileConnection()
        {
            var fileName = ConfigurationManager.AppSettings["TestDataSheetPath"];          
            var con = string.Format(@"https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl",Worldwide Rig Counts - Current & Historical Data);
            return con;
        }

        public static UserData GetRigcount(string Year)
        {
            using (var connection = new OleDbConnection(RigcountFileConnection()))
            {
                connection.Open();
                var query = string.Format("select * from WORLDWIDE RIG ROUND where Year between '2021' AND '2022.",Year);
                var value = connection.Query<UserData>(query).FirstOrDefault();
                connection.Close();
                return value;
            }
        }
    }
}
