using Google.Cloud.Firestore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
            InsertData();
        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
        }
        public async Task InsertData()
        {
            string str = "";
            string constring = ConfigurationSettings.AppSettings["ConnectionString"].ToString();
            string firestore = ConfigurationSettings.AppSettings["ProjectId"].ToString();
            DataTable dt = new DataTable();
            FirestoreDb db = FirestoreDb.Create(firestore);
            SqlCommand sqlCommand = new SqlCommand($"Select * from [dbo].[Source] WHERE SENT_TO_FIREBASE = 0");
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCommand);
            using (var sql1 = new SqlConnection(constring))
            {
                try
                {
                    sql1.Open();
                    sqlCommand.Connection = sql1;
                    adapter.Fill(dt);
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sql1.Close();
                    sqlCommand = null;
                }
            }
            if (dt.Rows.Count >= 1)
            {
                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        Dictionary<string, object> data = new Dictionary<string, object>
                        {
                            { "Code", row["Code"] },
                            { "Quantity", row["Quantity"] },
                            { "lineNo", row["lineNo"] },
                            { "location", row["location"] },
                        };

                        DocumentReference docRef = db.Collection("Source").Document();
                        await docRef.SetAsync(data);
                        //updatelist.Add(row["Id"].ToString());
                        str = str + " " + row["Id"].ToString() + ",";
                    }
                    catch
                    {

                    }
                }
                if (str.EndsWith(","))
                {
                    str = str.TrimEnd(',');
                }
                UpdateData(str);
            }
        }
        private void UpdateData(string sb)
        {
            string constring = ConfigurationSettings.AppSettings["ConnectionString"].ToString();
            SqlCommand sqlCommand = new SqlCommand($"UPDATE [dbo].[Source] SET SENT_TO_FIREBASE = 1 WHERE ID IN ({sb})");
            SqlDataAdapter adapter = new SqlDataAdapter(sqlCommand);
            using (var sql1 = new SqlConnection(constring))
            {
                try
                {
                    sql1.Open();
                    sqlCommand.Connection = sql1;
                    sqlCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sql1.Close();
                    sqlCommand = null;
                }
            }

        }
    }
}
