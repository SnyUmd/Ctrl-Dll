//#define debug_m_ArrayMake


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

//ひらがな判定に必要(Regex)
using System.Text.RegularExpressions;

//◆◆◆◆◆◆漢字をひらがなに変換するのに必要◆◆◆◆◆◆◆◆◆◆◆◆
using Microsoft.VisualBasic;//【.NET】
using ExcelApplication = Microsoft.Office.Interop.Excel.Application; //【.com】
//◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆


namespace Ctrl_Dll
{
    public class cls_Text_Cont
    {
        string New_Line = "\r\n";


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ひらがな判定
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool IsHiragana(string str)
        {
            return Regex.IsMatch(str, @"^\p{IsHiragana}*$");
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ひらがな以外を削除する
        /// </summary>
        /// <param name="Target_Str"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Only_Hiragana(string Target_Str)
        {
            string Str = "";
            string Retrun_Str = "";

            for (int i = 0; i < Target_Str.Length; i++)
            {
                Str = Target_Str.Substring(i, 1);
                if (Regex.IsMatch(Str, @"^\p{IsHiragana}*$"))
                    Retrun_Str += Str;
            }

            return Retrun_Str;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 漢字を【カタカナ】【ひらがな】に変換する
        /// </summary>
        /// <param name="Target_Str">変換する文字列</param>
        /// <param name="Change_Type_Katakana1_Hirangana2">1：カタカナ / 2：ひらがな</param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Change_Kana(string Target_Str, int Change_Type_Katakana1_Hirangana2)
        {
            string Return_Kana_Str = "";

            //Excelオブジェクト作成
            ExcelApplication objExcel = new ExcelApplication();
            //指定文字をカタカナに変換
            Return_Kana_Str = objExcel.GetPhonetic(Target_Str);

            if (Change_Type_Katakana1_Hirangana2 == 2)
                //カタカナをひらがなに変換
                Return_Kana_Str = Microsoft.VisualBasic.Strings.StrConv(Return_Kana_Str, Microsoft.VisualBasic.VbStrConv.Hiragana, 0x411);

            objExcel = null;
            return Return_Kana_Str;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字列を指定した文字で包む
        /// </summary>
        /// <param name="Str">対象文字列</param>
        /// <param name="Start_Str">先頭に追加する文字</param>
        /// <param name="End_Str">末尾に追加する文字</param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Str_Wrap(string Str, string Start_Str, string End_Str)
        {
            string Return_Str = Str;

            Return_Str = Return_Str.Insert(0, Start_Str);
            Return_Str = Return_Str.Insert(Return_Str.Length, End_Str);

            return Return_Str;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字列の先頭と末尾を削除する。
        /// </summary>
        /// <param name="Str">対象文字列</param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Str_Wrap_Rome(string Str)
        {
            string Retunf_Str = Str;
            Retunf_Str = Retunf_Str.Remove(0, 1);
            Retunf_Str = Retunf_Str.Remove(Retunf_Str.Length - 1, 1);

            return Retunf_Str;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字列の中の指定文字をすべて指定文字に変換
        /// </summary>
        /// <param name="Target_Str"></param>
        /// <param name="Target_Word"></param>
        /// <param name="Change_Word"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string StrA_Change_Str_B(string Target_Str, string Target_Word, string Change_Word)
        {
            int num = Character_Figure(Target_Str, Target_Word);

            string Return_Str = "";

            Return_Str = Target_Str.Replace(Target_Word, Change_Word);
            return Return_Str;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字列を指定文字から指定文字までで区切り、リスト化(Include:含める)
        /// </summary>
        /// <param name="Target_Str">操作文字列</param>
        /// <param name="Delimit_Wards_Start">スタート目印文字</param>
        /// <param name="Delimit_Wards_End">エンド目印文字</param>
        /// <param name="Delimit_Wards_Start_Include">true:スタート目印文字を含めて返す</param>
        /// <param name="Delimit_Wards_End_Include">true:エンド目印文字を含めて返す</param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public List<string> Str_Delimit_Set_List
            (string Target_Str, string Delimit_Wards_Start, string Delimit_Wards_End, bool Delimit_Wards_Start_Include, bool Delimit_Wards_End_Include)
        {
            List<string> Return_List = new List<string>();
            int num;
            int Start_Point = 0;
            int End_Point = 0;
            string Write_Str = "";
            //"の初めの位置を取得
            num = Character_Figure(Target_Str, Delimit_Wards_Start);
            //全員のデータを抽出
            Start_Point = Target_Str.IndexOf(Delimit_Wards_Start, Start_Point);//追加

            for (int i = 0; i < num; i++)
            {
                //Start_Point = Target_Str.IndexOf(Delimit_Wards_Start, Start_Point) + Delimit_Wards_Start.Length;

                End_Point = Target_Str.IndexOf(Delimit_Wards_End, Start_Point);
                if (End_Point < 0)
                    break;

                //スタート区切り文字を含まない場合
                //スタートポイントをスタート区切り文字分のシフトを行う。
                if (Delimit_Wards_Start_Include)
                    Start_Point = Start_Point - Delimit_Wards_Start.Length;

                //エンド区切り文字を含む場合
                //エンドポイントをエンド区切り文字分のシフト
                if (Delimit_Wards_End_Include)
                    End_Point = End_Point + Delimit_Wards_End.Length;

                //リストに書き込む文字をセット
                Write_Str = Target_Str.Substring(Start_Point, End_Point - Start_Point);
                Return_List.Add(Write_Str);

                //スタート区切り文字を含んだ時
                //次のスタート文字検索のために、スタート区切り文字文をポイントシフト
                if (Delimit_Wards_Start_Include)
                {
                    Start_Point = Start_Point + Delimit_Wards_Start.Length;
                }
                Start_Point = End_Point + 1;
            }

            if (Return_List.Count > 0 && Return_List[0].ToString() == "")
                Return_List.RemoveAt(0);

            return Return_List;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字列中に指定文字がいくつあるかを取得
        /// </summary>
        /// <param name="STR"></param>
        /// <param name="Serch"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Character_Figure(string STR, string Serch)
        {

            if (STR == null || STR == "") return -1;

            string Serch_Remove_STR;
            int Serch_Count = 0;

            Serch_Remove_STR = STR.Replace(Serch, "");
            Serch_Count = (STR.Length - Serch_Remove_STR.Length) / Serch.Length;

            return Serch_Count;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定文字の後に指定文字を加える
        /// </summary>
        /// <param name="str"></param>
        /// <param name="Serch_Char"></param>
        /// <param name="Add_Char"></param>
        /// <param name="All_Char"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Character_Add_Back
            (string str, string Serch_Char, string Add_Char, bool All_Char)
        {
            int Add_Point;
            int Serch_Char_Len;

            string Return_str;

            Add_Point = str.IndexOf(Serch_Char, 0);
            Serch_Char_Len = Serch_Char.Length;
            Add_Point += Serch_Char_Len;

            if (All_Char)
            {
                Return_str = str.Insert(Add_Point, Add_Char);
                while (str.IndexOf(Serch_Char, Add_Point) >= 0)
                {
                    Add_Point = str.IndexOf(Serch_Char, Add_Point) + Serch_Char_Len;
                    Return_str = str.Insert(Add_Point, Add_Char);
                }
            }
            else
            {
                Return_str = str.Insert(Add_Point, Add_Char);
            }

            return Return_str;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定文字の前に指定文字を加える
        /// </summary>
        /// <param name="str"></param>
        /// <param name="Serch_Char"></param>
        /// <param name="Add_Char"></param>
        /// <param name="All_Char"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Character_Add_Front
            (string str, string Serch_Char, string Add_Char, bool All_Char)
        {
            int Add_Point;
            int Serch_Char_Len;

            string Return_str;

            Add_Point = str.IndexOf(Serch_Char, 0);
            Serch_Char_Len = Serch_Char.Length;
            //Add_Point += Serch_Char_Len;

            if (All_Char)
            {
                Return_str = str.Insert(Add_Point, Add_Char);
                while (str.IndexOf(Serch_Char, Add_Point + Serch_Char_Len) >= 0)
                {
                    Add_Point = str.IndexOf(Serch_Char, Add_Point + Serch_Char_Len);
                    Return_str = str.Insert(Add_Point, Add_Char);
                }
            }
            else
            {
                Return_str = str.Insert(Add_Point, Add_Char);
            }

            return Return_str;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定文字から指定文字までにある文字列を取得
        /// </summary>
        /// <param name="STR"></param>
        /// <param name="Start_Str"></param>
        /// <param name="End_Str"></param>
        /// <param name="Start_Str_Include">true:スタート検索文字を含む</param>
        /// <param name="End_Str_Include">true:エンド検索文字を含む</param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Str_Extraction
            (string STR, string Start_Str, string End_Str, bool Start_Str_Include, bool End_Str_Include)
        {
            int Start_Point;
            int End_Point;

            Start_Point = STR.IndexOf(Start_Str, 0) + Start_Str.Length;
            End_Point = STR.IndexOf(End_Str, Start_Point);


            if (Start_Str_Include)
                Start_Point -= Start_Str.Length;
            if (End_Str_Include)
                End_Point += End_Str.Length;

            if (Start_Point >= 0 && End_Point > 0)
                return STR.Substring(Start_Point, End_Point - Start_Point);
            else if ((End_Point - Start_Point) <= 0)
                //MessageBox.Show("文字列配置が異常です。" + New_Line + "プログラムを確認してください。");
                return "ERR";
            //return STR.Substring(Str_Point1 + Str1.Length, Str_Point2 - Str_Point1);
            else
                return STR.Substring(Start_Point, End_Point - Start_Point);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定位置からn回目に出てくる指定文字の位置を検索
        /// </summary>
        /// <param name="Str"></param>
        /// <param name="Serch_Str"></param>
        /// <param name="Serch_Start_Point"></param>
        /// <param name="Num"></param>
        /// <param name="Rear_Point"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Str_Serch_Number_Of_Times
            (string Str, string Serch_Str, int Serch_Start_Point, int Num, bool Rear_Point)
        {
            int Point1 = Serch_Start_Point;
            //指定分の改行コード後の位置を取得
            for (int i = 0; i < Num; i++)
            {
                switch (i)
                {
                    //初期は文字検索のスタートポイント位置に検索文字分プラスしない。
                    case 0:
                        Point1 = Str.IndexOf(Serch_Str, Point1);
                        break;
                    default:
                        Point1 = Str.IndexOf(Serch_Str, Point1 + Serch_Str.Length);
                        break;
                }
            }
            if (Rear_Point)
            {
                return Point1 + Serch_Str.Length;
            }
            else
                return Point1;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定位置から見て、初めの指定文字から指定文字までにある文字列を取得
        /// </summary>
        /// <param name="Position"></param>
        /// <param name="STR"></param>
        /// <param name="Start_Str"></param>
        /// <param name="End_Str"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Str_Extraction_Specified_Positiont
            (int Position, string STR, string Start_Str, string End_Str)
        {
            int Start_Point;
            int End_Point;

            Start_Point = STR.IndexOf(Start_Str, Position) + Start_Str.Length;
            End_Point = STR.IndexOf(End_Str, Start_Point);

            if ((End_Point - Start_Point) <= 0)
            {
                //MessageBox.Show("文字列配置が異常です。" + New_Line + "プログラムを確認してください。");
                return "ERR";
                //return STR.Substring(Str_Point1 + Str1.Length, Str_Point2 - Str_Point1);
            }

            else
            {
                return STR.Substring(Start_Point, End_Point - Start_Point);
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定位置から指定文字手前までの文字列を取得
        /// </summary>
        /// <param name="STR"></param>
        /// <param name="Secified_Char"></param>
        /// <param name="Start_Point"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Secified_Char_Read(string STR, string Secified_Char, int Start_Point)
        {
            int Secified_Char_Point;
            int Secified_Char_Len;

            string Return_STR;

            ///Start_Point = 0;
            Secified_Char_Len = Secified_Char.Length;
            Secified_Char_Point = STR.IndexOf(Secified_Char, Start_Point);

            if (Secified_Char_Point == -1 || Secified_Char_Point == 0)
            {
                Return_STR = STR;
            }
            else
            {
                Return_STR = STR.Substring(Start_Point, Secified_Char_Point - Start_Point);
            }
            return Return_STR;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定文字から指定文字までにある文字列を削除
        /// </summary>
        /// <param name="STR"></param>
        /// <param name="Start_Str"></param>
        /// <param name="End_Str"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Str_Remove_Range(string STR, string Start_Str, string End_Str)
        {
            int Start_Point;
            int End_Point;

            Start_Point = STR.IndexOf(Start_Str, 0) + Start_Str.Length;
            End_Point = STR.IndexOf(End_Str, Start_Point);

            if ((End_Point - Start_Point) <= 0)
            {
                //MessageBox.Show("文字列配置が異常です。" + New_Line + "プログラムを確認してください。");
                return "ERR";
            }

            return STR.Remove(Start_Point, End_Point - Start_Point);
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字列を削除
        /// </summary>
        /// <param name="STR"></param>
        /// <param name="Del_Char"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Specified_Character_Delete(string STR, string Del_Char)
        {
            int Del_Point;
            int Del_Char_Len;

            Del_Char_Len = Del_Char.Length;
            Del_Point = STR.IndexOf(Del_Char);

            return STR.Remove(Del_Point, Del_Char_Len);
        }


        /// <summary>
        /// 指定文字を全て削除する。
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public string StrWordRemove(string str, string strRemove)
        {
            string strSet = str;
            strSet = strSet.Replace(strRemove, "");
            return strSet;
        }

        //文字列の種類判定**********************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、全角又は半角数字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Number_Judge(char c)
        {
            return '0' <= c && c <= '9' || '０' <= c && c <= '９';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、半角数字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Half_Size_Number_Judge(char c)
        {
            return '0' <= c && c <= '9';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、全角数字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Full_Size_Number_Judge(char c)
        {
            return '０' <= c && c <= '９';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、全角又は半角英字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Alphabet_Judge(char c)
        {
            return 'A' <= c && c <= 'Z' || 'Ａ' <= c && c <= 'Ｚ';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、半角英字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Half_Size_Alphabet_Judge(char c)
        {
            return 'A' <= c && c <= 'Z';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、全角英字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Full_Size_Alphabet_Judge(char c)
        {
            return 'Ａ' <= c && c <= 'Ｚ';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字が、英数字かを判定
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Number_Or_Alphabet_Judge(char c)
        {
            return
                'A' <= c && c <= 'Z' || 'Ａ' <= c && c <= 'Ｚ' ||
                '0' <= c && c <= '9' || '０' <= c && c <= '９';
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定した文字列に指定文字があるかを判定
        /// </summary>
        /// <param name="STR"></param>
        /// <param name="Judge_STR"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool String_Judge(string STR, string Judge_STR)
        {
            if (STR.IndexOf(Judge_STR) > -1) return true;
            else return false;
        }
        //文字列の種類判定**********************************************************************************


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// データ文字列を(多次元)配列化
        /// </summary>
        /// <param name="data_text"></param>
        /// <param name="s_group_word"></param>
        /// <param name="s_separate_word"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[,] m_ArrayMake(string data_text, string s_group_word, string s_separate_word)
        {
            int x, y;
            double d_p_start, d_p_end;
            double d_p_group_start, d_p_group_end;
            string s_data_1set;
            string[,] s_return; //後でセット

            //x = 1;
            //y = 2;
            //データのスタート位置を取得
            //d_start_point = this.Str_Serch_Number_Of_Times(data_text, s_group_word, 0, i_start_line - 1, true);
            //データ前のテキストを削除
            //data_text = data_text.Remove(0, (int)d_start_point);

            //データ数が何グループあるかを取得
            x = this.Character_Figure(data_text, s_group_word);
#if debug_m_ArrayMake
            MessageBox.Show("グループ数" + New_Line + x);
#endif

            //ポインタを先頭にセット
            d_p_group_start = 0;
            d_p_group_end = data_text.IndexOf(s_group_word, (int)d_p_group_start);
            //データの１セットを格納
            s_data_1set = data_text.Substring((int)d_p_group_start, (int)d_p_group_end - (int)d_p_group_start);

#if debug_m_ArrayMake
            MessageBox.Show("グループデータ" + New_Line + s_data_1set);
#endif
            //グループにデータがいくつあるか取得
            if (s_separate_word == null)
                y = 1;
            else
            {
                y = this.Character_Figure(s_data_1set, s_separate_word) + 1;
                if (y == 0)
                    y = 1;
            }
#if debug_m_ArrayMake
            MessageBox.Show("データ数" + New_Line + y + "/Group");
#endif

            s_return = new string[x, y];
            //-----------------------------------------------------------------------
            for (double d = 0; d < x; d++)
            {
                //データの１セットを格納
                s_data_1set = data_text.Substring((int)d_p_group_start, (int)d_p_group_end - (int)d_p_group_start);
#if debug_m_ArrayMake
                MessageBox.Show("グループデータ" + New_Line + s_data_1set);
#endif
                //データの先頭位置をセット
                d_p_start = 0;
                //-------------------------------------------------------------------
                for (int i = 0; i < y; i++)
                {
                    if (y == 1)
                        d_p_end = s_data_1set.Length;
                    else
                        //区切り位置を取得
                        d_p_end = this.Str_Serch_Number_Of_Times(s_data_1set, s_separate_word, (int)d_p_start, i, false);
                    //データを格納
                    s_return[(int)d, i] = s_data_1set.Substring((int)d_p_start, (int)d_p_end - (int)d_p_start);
#if debug_m_ArrayMake
                    MessageBox.Show("データ " + d + "-" + i + New_Line + s_return[(int)d, i]);
#endif
                    //次の先頭位置をセット
                    d_p_start += d_p_end + s_separate_word.Length;
                }
                //-------------------------------------------------------------------
                d_p_group_start = d_p_group_end + s_group_word.Length;
                d_p_group_end = data_text.IndexOf(s_group_word, (int)d_p_group_start);
                //d_p_group_end = this.Str_Serch_Number_Of_Times(data_text, s_group_word, (int)d_p_group_start, (int)d + 1, false);
            }
            //-----------------------------------------------------------------------
            return s_return;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lst_s_data"></param>
        /// <param name="s_separate_word"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[,] m_ListToAry_Data(List<string> lst_s_data, string s_separate_word)
        {
            int x, y;
            int i_p_start, i_p_end;
            string s_data_1set;
            string[,] s_return; //後でセット

            x = lst_s_data.Count;
            s_data_1set = lst_s_data[0];
            y = this.Character_Figure(s_data_1set, s_separate_word) + 1;
            s_return = new string[x, y];

            for (int i1 = 0; i1 < x; i1++)
            {
                s_data_1set = lst_s_data[i1];
                i_p_start = 0;
                for (int i2 = 0; i2 < y; i2++)
                {
                    //i_p_end = Str_Serch_Number_Of_Times(s_data_1set, s_separate_word, i_p_start, i2, false);
                    i_p_end = s_data_1set.IndexOf(s_separate_word, i_p_start);
                    if (i_p_end < 0)
                        s_return[i1, i2] = s_data_1set.Substring(i_p_start);
                    else
                        s_return[i1, i2] = s_data_1set.Substring(i_p_start, i_p_end - i_p_start);
                    i_p_start = i_p_end + s_separate_word.Length;
                }
            }
            return s_return;
        }



        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字列が数値であるか判定
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mStrIsNum_Judge(string str)
        {
            int i = 0;
            return int.TryParse(str, out i);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public List<string> mReadLine_TextFile(string str_path)
        {
            int num;
            string strBuf;
            List<string> lst_strReturn = new List<string>();

            //StreamReader sr = new StreamReader(str_data, Encoding.GetEncoding("SHIFT_JIS"));
            StreamReader sr = new StreamReader(str_path, Encoding.GetEncoding("SHIFT_JIS"));

            //num = Character_Figure(str_data, New_Line);

            while(true)
             {
                strBuf = sr.ReadLine();
                if (strBuf == null)
                    break;
                else
                    lst_strReturn.Add(strBuf);
            }

            sr.Close();
            return lst_strReturn;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// テキストを１行区切でリスト化
        /// </summary>
        /// <param name="str_data">対象文字列</param>
        /// <param name="str_separator">センサ区切り文字</param>
        /// <param name="i_data_num">１行内のセンサ数</param>
        /// <returns></returns>
        public List<string> mReadLine_String(string str_data, string str_separator, int i_data_num)
        {
            string strBuf;
            List<string> lst_strReturn = new List<string>();

            //StreamReader sr = new StreamReader(str_data, Encoding.GetEncoding("SHIFT_JIS"));
            System.IO.StringReader rs = new System.IO.StringReader(str_data);


            //num = Character_Figure(str_data, New_Line);

            while (true)
            {
                strBuf = rs.ReadLine();

                if (strBuf == null)
                    break;

                else if (i_data_num != Character_Figure(strBuf, str_separator))
                    continue;
                else
                    lst_strReturn.Add(strBuf);
            }

            rs.Close();
            return lst_strReturn;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 文字の出現回数をカウント
        /// </summary>
        /// <param name="s"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        public int mCountChar(string s, char c)
        {
            return s.Length - s.Replace(c.ToString(), "").Length;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
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


/*
#if debug
MessageBox.Show("" + New_Line + );
#endif
*/
