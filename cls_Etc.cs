using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Windows.Forms;

using System.Drawing; //参照の追加の事

using System.Net; //IPアドレス取得の為に

namespace Ctrl_Dll
{
    public class cls_Etc
    {
        string NewLine = "\r\n";
        string Comma = ",";

        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //お試し区域


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆Windows起動時刻の取得◆◆◆◆◆◆◆
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public  string Test_State_Up_Time()
        {
            System.Diagnostics.EventLog[] logs = System.Diagnostics.EventLog.GetEventLogs();
            StringBuilder sb = new StringBuilder();
            foreach (System.Diagnostics.EventLog log in logs)
            {
                if (log.Log == "System")
                {
                    int cnt = 0;
                    for (int i = log.Entries.Count - 1; i > 0; i--)
                    {

                        if (log.Entries[i].Source == "Microsoft-Windows-Kernel-General")
                        {

                            if (log.Entries[i].InstanceId == 12)
                            {
                                sb.AppendLine(log.Entries[i].TimeGenerated.ToString("[yyyy/MM/dd HH:mm:ss]") + "起動");
                                cnt++;
                            }
                            else if (log.Entries[i].InstanceId == 13)
                            {
                                sb.AppendLine(log.Entries[i].TimeGenerated.ToString("[yyyy/MM/dd HH:mm:ss]") + "終了");
                                cnt++;
                            }
                        }
                        if (cnt > 10)
                        {
                            break;
                        }
                    }
                }
            }
            return sb.ToString();
        }


        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★




        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆ウェイト関数 ◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void wait(int A)
        {
            System.Threading.Thread.Sleep(A);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆現在時刻の取得 ◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public DateTime Current_Time()
        {
            return System.DateTime.Now;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆現在時刻の取得 ◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Time_Value_To_int(DateTime T, int Sts)
        {
            int Return_int;

            switch (Sts)
            {
                case 1:
                    Return_int = int.Parse(T.ToString("yyyy"));
                    break;
                case 2:
                    Return_int = int.Parse(T.ToString("MM"));
                    break;
                case 3:
                    Return_int = int.Parse(T.ToString("dd"));
                    break;
                case 4:
                    Return_int = int.Parse(T.ToString("HH"));
                    break;
                case 5:
                    Return_int = int.Parse(T.ToString("mm"));
                    break;
                default:
                    MessageBox.Show("プログラムを確認してください。ステータスは1～5までです。" );
                    Return_int = -1;
                    break;

            }
            return Return_int;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆Windows最終起動時刻の取得◆◆◆◆◆
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public DateTime State_Up_Time()
        {
            System.Diagnostics.EventLog[] logs = System.Diagnostics.EventLog.GetEventLogs();
            StringBuilder sb = new StringBuilder();
            foreach (System.Diagnostics.EventLog log in logs)
            {
                if (log.Log == "System")
                {
                    for (int i = log.Entries.Count - 1; i > 0; i--)
                    {
                        if (log.Entries[i].Source == "Microsoft-Windows-Kernel-General" && log.Entries[i].InstanceId == 12)
                        {
                            return log.Entries[i].TimeGenerated;
                        }
                    }
                }
            }
            MessageBox.Show("起動時間の取得に失敗しました。");
            return System.DateTime.Now;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //◆◆◆◆◆◆Windows最終終了時刻の取得◆◆◆◆◆
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public DateTime Shutdown_Time()
        {
            System.Diagnostics.EventLog[] logs = System.Diagnostics.EventLog.GetEventLogs();
            StringBuilder sb = new StringBuilder();
            foreach (System.Diagnostics.EventLog log in logs)
            {
                if (log.Log == "System")
                {
                    for (int i = log.Entries.Count - 1; i > 0; i--)
                    {
                        if (log.Entries[i].Source == "Microsoft-Windows-Kernel-General" && log.Entries[i].InstanceId == 13)
                        {
                            return log.Entries[i].TimeGenerated;
                        }
                    }
                }
            }
            MessageBox.Show("終了時間の取得に失敗しました。");
            return System.DateTime.Now;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //郵便番号リストから、指定郵便番号の住所を取得
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Address_Code_Serch(string Adc_List, string Serch_Code)
        {
            int Point1;
            int Point2;

            string Address_Str1 = "";
            string Address_Str2 = "";
            string Address_Str3 = "";
            string Ad_Code;

            string Return_Str = "";

            //郵便番号部を取得
            Ad_Code = ",\"" + Serch_Code + "\",";


            //コードを検索
            Point1 = Adc_List.IndexOf(Ad_Code);

            if (Point1 >= 0)
            {
                //◆住所1先頭位置を取得
                for (int i = 0; i <= 4; i++)
                {
                    Point1 = Adc_List.IndexOf(",", Point1) + 1;
                    //ダブルクオーテーションをスルーするため。
                    if (i == 4) Point1 += 1;
                }
                //住所1の末尾位置を取得
                Point2 = Adc_List.IndexOf(",", Point1);
                //ダブルクオーテーションをスルーするため
                Point2 -= 1;
                //アドレス1を取得
                Address_Str1 = Adc_List.Substring(Point1, Point2 - Point1);

                //◆住所2を取得
                Point1 = Point2 + 3;
                Point2 = Adc_List.IndexOf(",", Point1);
                Point2 -= 1;
                Address_Str2 = Adc_List.Substring(Point1,Point2 - Point1);

                //◆住所3を取得
                Point1 = Point2 + 3;
                Point2 = Adc_List.IndexOf(",", Point1);
                Point2 -= 1;
                Address_Str3 = Adc_List.Substring(Point1, Point2 - Point1);

                Return_Str = Address_Str1 + Address_Str2 + Address_Str3;

            }
            return Return_Str;
        }



        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// タスクトレイのアイコンにツールチップをバルーンで表示する。
        /// </summary>
        /// <param name="NotifyIcon_Ctrl"></param>
        /// <param name="Tytle"></param>
        /// <param name="Mess"></param>
        /// <param name="View_Time"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Baloon_Show_Task_Tray(NotifyIcon NotifyIcon_Ctrl, string Tytle, string Mess, int View_Time)
        {
            //バルーンヒントの設定
            //バルーンヒントのタイトル
            NotifyIcon_Ctrl.BalloonTipTitle = Tytle;
            //バルーンヒントに表示するメッセージ
            NotifyIcon_Ctrl.BalloonTipText = Mess;
            //バルーンヒントに表示するアイコン
            NotifyIcon_Ctrl.BalloonTipIcon = ToolTipIcon.Info;
            //バルーンヒントを表示する
            //表示する時間をミリ秒で指定する
            NotifyIcon_Ctrl.ShowBalloonTip(View_Time);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ツールチップをコントロール上に表示する
        /// </summary>
        /// <param name="ToolT"></param>
        /// <param name="Ctrl"></param>
        /// <param name="Tytle"></param>
        /// <param name="Mess"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //public void Baloon_Show_Ctrl(ToolTip View_ToolTip, Control Target_Ctrl, string Tytle, string Mess)
        public void Baloon_Show_Ctrl(Control Target_Ctrl, string Tytle, string Mess)
        {
            //ToolTipオブジェクトを作成する
            //this.componentsがない時は、this.componentsを省略する
            //System.Windows.Forms.ToolTip toolTip1 =
            //    new System.Windows.Forms.ToolTip(this.components);
            System.Windows.Forms.ToolTip toolTip1 =
                new System.Windows.Forms.ToolTip();

            //ツールチップをバルーンウィンドウとして表示する
            //View_ToolTip.IsBalloon = true;
            toolTip1.IsBalloon = true;

            //
            toolTip1.BackColor = Color.Yellow;

            //ツールチップのタイトル
            //View_ToolTip.ToolTipTitle = Tytle;
            toolTip1.ToolTipTitle = Tytle;
            //ツールチップに表示するアイコン
            //View_ToolTip.ToolTipIcon = ToolTipIcon.Info;
            toolTip1.ToolTipIcon = ToolTipIcon.Info;
            //ツールチップを表示するコントロールと、表示するメッセージ
            //View_ToolTip.SetToolTip(Target_Ctrl, Mess);
            toolTip1.SetToolTip(Target_Ctrl, Mess);
        }

        //*********************************************************************************
        //IPアドレスの取得
        //*********************************************************************************
        public string mGet_IPAddress()
        {
            /*// ホスト名を取得する
            string hostname = Dns.GetHostName();

            // ホスト名からIPアドレスを取得する
            IPAddress[] adrList = Dns.GetHostAddresses(hostname);

            return adrList.a.ToString();*/

            string ipaddress = "";
            //ホスト名を取得
            IPHostEntry ipentry = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress ip in ipentry.AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    ipaddress = ip.ToString();
                    break;
                }
            }
            return ipaddress;
        }

        //*********************************************************************************
        /// <summary>
        /// 指定URLのテキストを取得
        /// </summary>
        /// <param name="strURL"></param>
        /// <returns></returns>
        //*********************************************************************************
        public string mRead_URL_Text(string strURL)
        {
            string strRTN;
            try
            {
                WebClient wc = new WebClient();

                Stream st = wc.OpenRead(strURL);
                StreamReader sr = new StreamReader(st);
                strRTN = sr.ReadToEnd();
                return strRTN;
            }
            catch(Exception)
            {
                return "err";
            }
        }
    }
}
