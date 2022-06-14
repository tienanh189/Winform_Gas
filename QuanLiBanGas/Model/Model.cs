using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace QuanLiBanGas.Model
{
    public class Model
    {
        private static Model instance; 
        private string strCon = @"Data Source=ADMIN\SQLEXPRESS01;Initial Catalog=QLBanGa;Integrated Security=True";
       
        public static Model Instance
        {
            get { if (instance == null) instance = new Model(); return Model.instance; }
            private set { Model.instance = value; }
        }
        private Model() { }
        public DataTable GetTable(string str,object[] para = null) 
        {
            DataTable dataTable = new DataTable();

            using(SqlConnection sqlConnection = new SqlConnection(strCon))
            {
                sqlConnection.Open();

                SqlCommand command = new SqlCommand(str, sqlConnection);

                if (para != null)
                {
                    string[] listP = str.Split(' ');
                    int i = 0;
                    foreach (var item in listP)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, para[i]);
                        }
                    }
                }
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(dataTable);
                sqlConnection.Close();
            }
            return dataTable;
        }//Dùng khi muốn in bảng
        public int GetResIUD(string str, object[] para = null)
        {
            int res = 0;

            using (SqlConnection sqlConnection = new SqlConnection(strCon))
            {
                sqlConnection.Open();

                SqlCommand command = new SqlCommand(str, sqlConnection);

                if (para != null)
                {
                    string[] listP = str.Split(' ');
                    int i = 0;
                    foreach (var item in listP)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, para[i]);
                        }
                    }
                }
                res = command.ExecuteNonQuery();
                sqlConnection.Close();
               
            }
            return res;
        }//Dùng khi Insert, Update, Delete
        public object GetScalar(string str, object[] para = null)
        {
            object res = 0;

            using (SqlConnection sqlConnection = new SqlConnection(strCon))
            {
                sqlConnection.Open();

                SqlCommand command = new SqlCommand(str, sqlConnection);

                if (para != null)
                {
                    string[] listP = str.Split(' ');
                    int i = 0;
                    foreach (var item in listP)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, para[i]);
                        }
                    }
                }

                res = command.ExecuteScalar();
                sqlConnection.Close();
                
            }
            return res;
        }//Dùng khi in ra 1 cột duy nhất


    }
}
