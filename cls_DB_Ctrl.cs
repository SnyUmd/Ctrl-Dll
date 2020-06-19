//#define Evacuation

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.SqlClient;//GetConnectionString()のために必要

namespace Ctrl_Dll
{
    public class cls_DB_Ctrl
    {
        /*
        //******************************************************
        //サーバ接続先情報セット
        //******************************************************
        public string M_DB_Open(string str_server_name, string str_ip_add, string str_user_name, string str_pass)
        {
            return String.Format(@"Data Source=({0}/{1});" +
                                 @"Integrated Security=False;" +
                                 @"User ID=({2});" +
                                 @"Password=({3})", str_server_name, str_ip_add, str_user_name, str_pass);
        }
        */

        /*
        //*************************************************************************************************
        //サーバ接続先情報文字列のセット
        //*************************************************************************************************
        public string M_GetConnectionString(string str_server_name, string str_ip_add, string str_user_name, string str_pass)
        {
            var builder = new SqlConnectionStringBuilder()
            {
                DataSource = String.Format("({0}/{1})",str_server_name, str_ip_add),
                IntegratedSecurity = false,
                UserID = String.Format("({0})", str_user_name),
                Password = String.Format("({0})", str_pass)
            };

            return builder.ToString();
        }
        */



        //*************************************************************************************************
        //データベースへ接続
        //*************************************************************************************************
        public void mDBConnect(ref SqlConnection sc,
                                string str_server_name,
                                string str_db_dir)
        {
            // データベース接続の準備
            sc = new SqlConnection();
            sc.ConnectionString = 
                string.Format(@"Data Source={0};
                              AttachDbFilename = {1};
                              Integrated Security=True;
                              Connect Timeout=30", str_server_name, str_db_dir);
            /*
            sc.ConnectionString =
                string.Format(@"Data Source={0};
                              Initial Catalog={1};
                              Integrated Security=True;
                              Connect Timeout=30;
                              Encrypt=False;
                              TrustServerCertificate=True;
                              ApplicationIntent=ReadWrite;
                              MultiSubnetFailover=False", str_server_name, str_db_dir);
            */

            // データベースの接続開始
            sc.Open();

            MessageBox.Show("オープン完了");
        }

        //*************************************************************************************************
        //データベース切断
        //参照渡しにする事で、scがnullになるのを防げた
        //*************************************************************************************************
        public void mDB_DisConnect(ref SqlConnection sc)
        {
            sc.Close();

            MessageBox.Show("クローズ完了");
        }

        //*************************************************************************************************
        //レコード数取得
        //*************************************************************************************************
        public int mDB_RecordCount(ref SqlConnection sc, string str_db_name, string str_table_name)
        {
            //SqlCommand cmd = new SqlCommand(@"SELECT COUNT(*) FROM [@DB_NAME].[dbo].[@TABLE_NAME]");
            //cmd.Parameters.Add(new SqlParameter("@DB_NAME", str_db_name));
            //cmd.Parameters.Add(new SqlParameter("@TABLE_NAME", str_table_name));
            int Count;
            

            using (var command = sc.CreateCommand())
            {
                //SQL文
                //command.CommandText = @"SELECT COUNT(*) FROM [部品管理_0_1].[dbo].[member]";
                //command.CommandText = @"SELECT COUNT(*) FROM [" + str_db_name + "].[dbo].[" + str_table_name + "]";
                command.CommandText = String.Format(@"SELECT COUNT(*) FROM [{0}].[dbo].[{1}]" , str_db_name, str_table_name);
                //SQL実行
                Count = (int)command.ExecuteScalar();
            }
            
            return Count;
        }

        //*************************************************************************************************
        //列数の取得
        //*************************************************************************************************
        public int mDB_ColCount(ref SqlConnection sc, string str_db_name, string str_db_table_name)
        {
            int Count;
            using (var command = sc.CreateCommand())
            {
                //列の取得のクエリ文
                command.CommandText = 
                    String.Format(@"SELECT COUNT(*) FROM {0}.SYS.COLUMNS 
                                  WHERE OBJECT_ID = OBJECT_ID('{0}.dbo.{1}')", str_db_name, str_db_table_name);
                //列数を取得
                Count = (int)command.ExecuteScalar();
            }
            //列数を返す
            return Count;
        }






        //*************************************************************************************************
        //SQL文の実行
        //*************************************************************************************************
        public void mDB_SQL_Set(ref SqlConnection sc, string str_sql)
        {
            try
            {
                using (var command = sc.CreateCommand())
                {
                    //SQL文
                    command.CommandText = str_sql;
                    //SQL実行
                    command.ExecuteNonQuery();
                    MessageBox.Show("実行完了");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("実行失敗");
                return;
            }
        }

        //================================================================================================================
        //退避
        //================================================================================================================
#if Evacuation
        //*************************************************************************************************
        //データベースへ接続
        //参照渡しにする事で、クローズの際にscがnullになるのを防げた
        //*************************************************************************************************
        public void mDBConnect(ref SqlConnection sc, string str_db)
        {
            // データベース接続の準備
            sc = new SqlConnection(str_db);
            // データベースの接続開始
            sc.Open();

            MessageBox.Show("オープン完了");
        }
#endif
    }
}
