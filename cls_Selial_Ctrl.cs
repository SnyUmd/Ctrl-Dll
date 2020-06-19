using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Windows.Forms;


namespace Ctrl_Dll
{
    public class cls_Selial_Ctrl
    {
        string stNewLine = "\r\n";
        
        //************************************************************************
        /// <summary>
        /// 利用可能なシリアルポート名リストを取得
        /// </summary>
        /// <returns>ポート名のリストを返す</returns>
        //************************************************************************
        public List<string> mGetPortNames()
        {
            List<string> lst_stReturn = new List<string>();

            lst_stReturn.AddRange(SerialPort.GetPortNames());
            return lst_stReturn;
        }


        //************************************************************************
        /// <summary>
        /// 利用可能なシリアルポートをコンボボックスにセット
        /// </summary>
        /// <param name="cmb">コンボボックス(参照渡し)</param>
        /// <returns></returns>
        //************************************************************************
        public bool mSerialPortSet(ref ComboBox cmb)
        {
            List<string> lst_st = new List<string>();

            lst_st.AddRange(SerialPort.GetPortNames());

            if (lst_st.Count > 0)
            {
                foreach (string strP in lst_st)
                    cmb.Items.Add(strP);
                cmb.SelectedIndex = 0;
                return true;
            }
            else
                return false;
        }




        //************************************************************************
        /// <summary>
        /// シリアルポートをオープンする
        /// </summary>
        /// <param name="port">シリアルポート</param>
        /// <param name="st_port_name">シリアルポート名</param>
        /// <returns>オープン実行結果を返す</returns>
        //************************************************************************
        public bool mSerialOpen(SerialPort port, string st_port_name, int i_baudrate)
        {
            //bool blReturn = false;

            switch (i_baudrate)
            {
                case 110:
                case 300:
                case 600:
                case 1200:
                case 2400:
                case 4800:
                case 9600:
                case 14400:
                case 19200:
                case 38400:
                case 57600:
                case 115200:
                case 230400:
                case 460800:
                case 921600:
                    port.BaudRate = i_baudrate;
                    break;
                default:
                    MessageBox.Show("Programエラー" + stNewLine + 
                                    "ボーレートの値(" + i_baudrate + ")が正しくありません。");
                    return false;
                    //break;
            }

            port.PortName = st_port_name;
            try
            {
                port.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ポートのオープンに失敗しました。" + stNewLine + 
                                "ポート名を確認してください。<" + st_port_name + ">");
                return false;
            }
            
            return true;
        }


        //************************************************************************
        /// <summary>
        /// シリアルポート データ送信
        /// </summary>
        /// <param name="port">シリアルポート</param>
        /// <param name="stSend">送信テキスト</param>
        /// <returns>送信実行結果を返す</returns>
        //************************************************************************
        public bool m_bl_SerialSend(SerialPort port, string stSend)
        {
            //! シリアルポートをオープンしていない場合、処理を行わない.
            if (port.IsOpen == false)
            {
                MessageBox.Show("シリアルポートが開かれていません。");
                return false;
            }
            try
            {
                port.Write(stSend);
            }
            catch (Exception ex)
            {
                MessageBox.Show("送信に失敗しました。");
                return false;
            }
            return true;
        }


        //************************************************************************
        /// <summary>
        /// シリアルポート データ受信
        /// </summary>
        /// <param name="port">ポート</param>
        /// <returns>テキストデータを返す</returns>
        //************************************************************************
        public string m_bl_SerialReceived(SerialPort port)
        {
            string str_Return  = "";
            //! シリアルポートをオープンしていない場合、処理を行わない.
            if (port.IsOpen == false)
            {
                MessageBox.Show("シリアルポートが開かれていません。");
                return "err";
            }

            try
            {
                str_Return = port.ReadTo("\r");
            }
            catch (Exception ex)
            {
                MessageBox.Show("受信に失敗しました");
                return "err";
            }

            return str_Return;
        }

        //************************************************************************
        public string mSerialReadLine(SerialPort serial_p)
        {
            return serial_p.ReadLine();
        }

    }
}



//************************************************************************
//************************************************************************
//************************************************************************
