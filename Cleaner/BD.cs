using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cleaner
{
    class BD
    {
        private string sConnectionString = "Server=whhl7mb769.database.windows.net,1433;Password=F1gu1M3x.Wf.2016;User ID=FM_User@whhl7mb769;database=dobleceroDB2;Integrated security=SSPI; Trusted_Connection=False;Encrypt=True;";
        private SqlConnection objCon;

        public BD() {
        }

        /// <summary>
        /// This method only open a connection for the object SqlConnection
        /// </summary>
        public void OpenConnection() {
            this.objCon = new SqlConnection(this.sConnectionString);
            this.objCon.Open();
        }

        /// <summary>
        /// This method make an insert for the database with the format of the excel file
        /// </summary>
        /// <param name="query"></param>
        /// <param name="categoria"></param>
        /// <param name="subcategoria"></param>
        /// <param name="nombre"></param>
        /// <param name="codigo"></param>
        /// <param name="descripcion"></param>
        /// <param name="medidas"></param>
        /// <param name="material"></param>
        /// <param name="color"></param>
        /// <param name="precio"></param>
        /// <param name="precio_publico"></param>
        /// <param name="idcat"></param>
        /// <param name="idsubcat"></param>
        /// <param name="stoday"></param>
        /// <returns></returns>
        public int Fill(string query, string categoria, string subcategoria, string nombre, string codigo, string descripcion, string medidas, string material, string color, string precio, string precio_publico, string idcat, string idsubcat, string stoday) {
            //string query = "insert into Articulos(Nombre,Descripcion,Codigo,Medidas,Material,Precio,PrecioPublico,IdCategoria,IdSubCategoria,FechaAlta) values(@nom,@desc,@cod,@med,@mat,@p,@pp,@idcat,@idsubcat,@fecha)";
            SqlCommand command = new SqlCommand(query, objCon);
            command.Parameters.AddWithValue("@nom", nombre);
            command.Parameters.AddWithValue("@desc", descripcion);
            command.Parameters.AddWithValue("@cod", codigo);
            command.Parameters.AddWithValue("@med", medidas);
            command.Parameters.AddWithValue("@mat", material);
            command.Parameters.AddWithValue("@p", precio);
            command.Parameters.AddWithValue("@pp",precio_publico);
            command.Parameters.AddWithValue("@idcat", idcat);
            command.Parameters.AddWithValue("@idsubcat", idsubcat);
            command.Parameters.AddWithValue("@fecha", stoday);

            int rows = command.ExecuteNonQuery();
            return rows;
        }

        /// <summary>
        /// This method is for make a request for the database like a select
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public SqlDataReader GetDataSet(string query) {
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            cmd.CommandText = query;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = objCon;

            reader = cmd.ExecuteReader();
            return reader;
        }

        public string GetData(string query,string categoria)
        {

            SqlCommand command = new SqlCommand(query,objCon);
            
            command.Parameters.AddWithValue("@nom_categoria",categoria);
            int result = (Int32)command.ExecuteScalar();

            return result.ToString();
        }

        /// <summary>
        /// This method release the resources and close the connection
        /// </summary>
        public void CloseConnection() {
            this.objCon.Dispose();
            this.objCon.Close();
        }
    }
}
