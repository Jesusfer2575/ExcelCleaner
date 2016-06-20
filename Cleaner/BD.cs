using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cleaner
{
    class BD
    {
        private string sConnectionString = "Server=whhl7mb769.database.windows.net,1433;Password=F1gu1M3x.Wf.2016;User ID=FM_User@whhl7mb769;database=dobleceroDB;Integrated security=SSPI; Trusted_Connection=False;Encrypt=True;";
        private SqlConnection objCon;
        public BD() {
        }

        public void OpenConnection() {
            this.objCon = new SqlConnection(this.sConnectionString);
            this.objCon.Open();
        }

        public void CloseConnection() {
            this.objCon.Dispose();
            this.objCon.Close();
        }
    }
}
