using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ctrl_Dll
{

    public class cls_M_D_umd
    {
        string strNL = "\r\n";

        public string strURL_Headder = @"https://www.4sync.com/web/directDownload/";
        public int[] ary_iENC = { 2, 0, 1, 7, 0, 9, 0, 2 };


        //*****************************************************************************************
        public string mEnctyption(string str)
        {
            string strRTN = "";
            string strTarget = str;
            int iLoop;
            int iBackSlash_cnt = str.Length - str.Replace(@"/", "").Length;//"\"の出現回数を格納

            
            if (iBackSlash_cnt < 5)
                return "err";//____________________________


            strTarget = strTarget.Replace(strURL_Headder, "");//ヘッダーを削除
            
            iLoop = strTarget.Length;


            int iENC_Num = 0;
            int iBuf;
            string strBuf;
            char cBuf;
            bool blAddition = true;
            for (int i = 0; i < iLoop; i++)
            {
                if (iENC_Num == ary_iENC.Count())
                {
                    iENC_Num = 0;
                    blAddition = !blAddition;
                }

                strBuf = strTarget.Substring(i, 1);
                cBuf = Convert.ToChar(strBuf);

                if (blAddition)
                    //iBuf = int.Parse(strBuf) + ary_iENC[iENC_Num];
                    iBuf = (int)cBuf + ary_iENC[iENC_Num];
                else
                    //iBuf = int.Parse(strBuf) - ary_iENC[iENC_Num];
                    iBuf = (int)cBuf - ary_iENC[iENC_Num];

                //strRTN = iBuf.ToString() + strRTN; 
                strRTN = (char)iBuf + strRTN;

                if (i < iLoop - 1)
                {
                    iENC_Num++;
                    //strRTN = "," + strRTN;
                }
            }
            return strRTN;
        }

        //*****************************************************************************************
        public string mRestoration(string str)
        {
            string strRTN = "";
            string strTarget = str;
            int iLoop;

            int iENC_Num = 0;
            int iBuf;
            string strBuf;
            bool blAddition = false;
            char cBuf;

            strTarget = mHTML_SpecialCharacter_Change(strTarget);

            iLoop = strTarget.Length;
            for (int i = 0; i < iLoop; i++)
            {
                if (iENC_Num == ary_iENC.Count())
                {
                    iENC_Num = 0;
                    blAddition = !blAddition;
                }

                strBuf = strTarget.Substring(strTarget.Length - 1 - i, 1);
                cBuf = Convert.ToChar(strBuf);

                if (blAddition)
                    //iBuf = int.Parse(strBuf) + ary_iENC[iENC_Num];
                    iBuf = (int)cBuf + ary_iENC[iENC_Num];
                else
                    iBuf = (int)cBuf - ary_iENC[iENC_Num];
                strRTN += (char)iBuf;

                if (i < iLoop - 1)
                    iENC_Num++;
            }

            return strURL_Headder + strRTN;
        }

        //*****************************************************************************************
        public string mHTML_SpecialCharacter_Change(string str)
        {
            string strRTN = str;
            char cDoubleQ = '"';
            char cSpace = Convert.ToChar(32);

            strRTN = strRTN.Replace("&gt;", ">");
            strRTN = strRTN.Replace("&lt;", "<");
            strRTN = strRTN.Replace("&quot;", cDoubleQ.ToString());//ダブルクォーテーション
            strRTN = strRTN.Replace("&amp;", "&");
            strRTN = strRTN.Replace("&#39;", "'");
            strRTN = strRTN.Replace("&nbsp;", cSpace.ToString());//空白文字

            return strRTN;
        }
    }
}
