using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Windows.Forms;


//excel操作のために必要
using Excel = Microsoft.Office.Interop.Excel;

namespace Ctrl_Dll
{
    public class cls_Excel_Cont
    {

        //Etc1 ETC = new Etc1();
        //Text_Cont TC = new Text_Cont();

        public static Excel.Application EXL_App = null;
        public static Excel.Workbooks EXL_Book = null;
        public static Excel.Worksheet EXL_Sheet = null;
        public static Excel.Range EXL_Cell = null;


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Excelファイルを開く
        /// </summary>
        /// <param name="File_Name"></param>
        /// <param name="ReadOnly"></param>
        /// <param name="Visible_Enable"></param>
        /// <param name="Sheet_No"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void EXL_Open(string File_Name, bool ReadOnly, bool Visible_Enable, int Sheet_No)
        {

            //Excel起動
            EXL_App = new Excel.Application();

            //Excel表示
            EXL_App.Visible = Visible_Enable;
            EXL_App.DisplayAlerts = false;
            //EXL_App.Visible = true;
            //EXL_App.DisplayAlerts = true;

            //Excelブック取得
            EXL_Book = EXL_App.Workbooks;
            EXL_Book.Open(File_Name,  　//　指定ファイルPath
                                Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                                ReadOnly, // （省略可能）ReadOnly (True / False )
                                Type.Missing, // （省略可能）Format
                // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                // 5:なし / 6:引数 Delimiterで指定された文字
                                Type.Missing, // （省略可能）Password
                                Type.Missing, // （省略可能）WriteResPassword
                                Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
                                Type.Missing, // （省略可能）Origin
                                Type.Missing, // （省略可能）Delimiter
                                Type.Missing, // （省略可能）Editable
                                Type.Missing, // （省略可能）Notify
                                Type.Missing, // （省略可能）Converter
                                Type.Missing, // （省略可能）AddToMru
                                Type.Missing, // （省略可能）Local
                                Type.Missing  // （省略可能）CorruptLoad
                                );
            EXL_Sheet = (Excel.Worksheet)EXL_App.Worksheets[Sheet_No];
            EXL_Cell = (Excel.Range)EXL_Sheet.Cells[1, 1];
            //EXL_Cell = EXL_Sheet.get_Range(EXL_Sheet.Cells[1, 1], EXL_Sheet.Cells[1, 1]);
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell書き込み
        /// </summary>
        /// <param name="Sheet_No"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="STR"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Write(int Sheet_No, int X, int Y, string STR)
        {
            EXL_Sheet = (Excel.Worksheet)EXL_App.Worksheets[Sheet_No];
            //EXL_Sheet = (Excel.Worksheet)EXL_App.Worksheets[1];
            //Sheet_Change(Sheet_No);
            EXL_Cell = (Excel.Range)EXL_Sheet.Cells[X, Y];
            //EXL_Cell.Cells[X,Y] = STR;// = (Excel.Range)EXL_Sheet.Cells[X, Y];

            EXL_Cell.Value2 = STR;
        }

        //セルの読み込み***********************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 現在のセル位置の文字列取得
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Cell_Read_Now()
        {
            if (EXL_Cell.Value2 == null) return "";
            else return EXL_Cell.Value2.ToString();
        }



        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell読み込み
        /// </summary>
        /// <param name="Sheet_No"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Cell_Read(int Sheet_No, int X, int Y)
        {
            EXL_Sheet = (Excel.Worksheet)EXL_App.Worksheets[Sheet_No];
            EXL_Cell = (Excel.Range)EXL_Sheet.Cells[X, Y];
            if (EXL_Cell.Value2 == null) return "";
            return EXL_Cell.Value2.ToString();
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell複数行を配列格
        /// 指定行に記入されているCellの文字列を格納
        /// 指定数量で格納をストップ　
        /// </summary>
        /// <param name="Cell_Start_X"></param>
        /// <param name="Cell_Start_Y"></param>
        /// <param name="Quantity"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[] Cell_Read_Virtical(int Cell_Start_X, int Cell_Start_Y, int Quantity)
        {
            int X = Cell_Start_X;
            int Y = Cell_Start_Y;

            string[] STR = new string[Quantity];

            for (int i = 0; i < Quantity; i++)
            {
                Cell_Change(X + i, Y);
                if (EXL_Cell.Value2 == null) STR[i] = "";
                else STR[i] = EXL_Cell.Value2.ToString();
            }
            return STR;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定行の記入のあるCellの数を取得
        /// </summary>
        /// <param name="Start_X"></param>
        /// <param name="Y"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Cell_Check_Virtical(int Start_X, int Y)
        {
            int Cnt = 0;
            int Blank_Cnt = 0;
            int i = 0;

            while (true)
            {
                Cell_Change(Start_X + i, Y);

                if (EXL_Cell.Value2 == null)
                {
                    Blank_Cnt++;
                    i++;
                    if (Blank_Cnt >= 50) break;
                }
                else
                {
                    Blank_Cnt = 0;
                    Cnt++;
                    i++;
                }
            }
            return Cnt;
        }

        //セルの読み込み***********************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 選択中のCellを罫線で囲う
        /// </summary>
        /// <param name="Style_Type_No"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Enclose(int Style_Type_No)
        {
            //0：線無し
            //1:全囲い
            EXL_Cell.Borders.LineStyle = Style_Type_No;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定のCellを罫線で囲う
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Enclose(int X, int Y)
        {
            Cell_Change(X, Y);
            EXL_Cell.Borders.LineStyle = 1;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル記入箇所数取得
        /// ※列Yの並びの途中にスペースが無い事が条件
        /// </summary>
        /// <param name="Start_X"></param>
        /// <param name="Y"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Cell_Fill_In_The_Num_Virtical(int Start_X, int Y)
        {
            //Excel_Cont EC = new Excel_Cont();

            int X = 0;
            int Add_Q = 10000;
            int Return_Int;

            X = Add_Q;
            Cell_Change(X, Y);
            if (EXL_Cell.Value2 == null)
            {
                Add_Q = Add_Q / 10;// *-1;
                for (int i = 0; i < 10; i++)
                {
                    X = X - Add_Q;
                    if (X < Start_X)
                        X = Start_X;
                    Cell_Change(X, Y);

                    if (EXL_Cell.Value2 != null)
                        break;
                }

                Add_Q = Add_Q / 10;
                for (int i = 0; i < 10; i++)
                {
                    X = X + Add_Q;
                    Cell_Change(X, Y);

                    if (EXL_Cell.Value2 == null)
                        break;
                }

                while (true)
                {
                    X--;
                    Cell_Change(X, Y);
                    if (EXL_Cell.Value2 != null)
                    {
                        Return_Int = X;
                        break;
                    }
                }
            }

            else
            {
                //下り
                while (true)
                {
                    X = X + Add_Q;
                    Cell_Change(X, Y);
                    if (EXL_Cell.Value2 == null)
                        break;
                }
            }

            Add_Q = Add_Q / 10;
            for (int i = 0; i < 10; i++)
            {
                X = X - Add_Q;
                if (X < Start_X) X = Start_X;
                Cell_Change(X, Y);

                if (EXL_Cell.Value2 != null)
                    break;
            }

            Add_Q = Add_Q / 10;
            for (int i = 0; i < 10; i++)
            {
                X = X + Add_Q;
                Cell_Change(X, Y);

                if (EXL_Cell.Value2 == null)
                    break;
            }

            while (true)
            {
                Cell_Change(X, Y);
                if (EXL_Cell.Value2 != null)
                {
                    Return_Int = X;
                    break;
                }
                else
                    X--;
            }
            return Return_Int - (Start_X - 1);
        }


        //セルの追加*****************************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cellの追加
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="Holizontal1_or_Virtical2"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Add(int X, int Y, int Holizontal1_or_Virtical2)
        {
            EXL_Cell = (Excel.Range)EXL_Sheet.Cells[1, 1];
            EXL_Cell.Insert(Holizontal1_or_Virtical2, Type.Missing);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 選択分のセルを追加する
        /// </summary>
        /// <param name="S_X"></param>
        /// <param name="S_Y"></param>
        /// <param name="E_X"></param>
        /// <param name="E_Y"></param>
        /// <param name="Add_Holizontal1_or_Add_Virtical2"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Add_Multiple(int S_X, int S_Y, int E_X, int E_Y, int Add_Holizontal1_or_Add_Virtical2)
        //public void Cell_Add_Multiple(int X_or_Y, int Add_Holizontal1_or_Add_Virtical2, int Add_Start, int Add_End)
        {
            Cell_Change(1, 1);

            EXL_Cell.get_Range
                (EXL_Cell[S_X, S_Y], EXL_Cell[E_X, E_Y]).Insert(Add_Holizontal1_or_Add_Virtical2, Type.Missing);

            /*//指定されたスタートポイントからエンドポイントまでループ
            for (int i = Add_Start; i <= Add_End; i++)
            {
                if (Add_Holizontal1_or_Add_Virtical2 == 1)
                {
                    Cell_Change(i, X_or_Y);
                }

                if (Add_Holizontal1_or_Add_Virtical2 == 2)
                {
                    Cell_Change(X_or_Y, i);
                }
                else { MessageBox.Show("【Holizontal1_or_Virtical2】は１か２のみです。"); return; }

                EXL_Cell.Insert(Add_Holizontal1_or_Add_Virtical2, Type.Missing);
            }*/
        }
        //セルの追加*****************************************************************************************************



        //セルの削除*****************************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cellの削除
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="Holizontal1_or_Virtical2"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Dele(int X, int Y, int Holizontal1_or_Virtical2)
        {
            Cell_Change(X, Y);
            EXL_Cell.Delete(Holizontal1_or_Virtical2);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cellをまとめて削除
        /// </summary>
        /// <param name="Start_X"></param>
        /// <param name="Start_Y"></param>
        /// <param name="End_X"></param>
        /// <param name="End_Y"></param>
        /// <param name="Holizontal1_or_Virtical2"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Dele_Multi(int Start_X, int Start_Y, int End_X, int End_Y, int Holizontal1_or_Virtical2)
        {
            Cell_Change(1, 1);
            EXL_Cell.get_Range(EXL_Cell[Start_X, Start_Y], EXL_Cell[End_X, End_Y]).Delete(Holizontal1_or_Virtical2);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell行(横)を削除
        /// </summary>
        /// <param name="Cell_X"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Dele_Holizontal(int Cell_X)
        {
            Cell_Change(1, 1);
            EXL_Cell.get_Range(EXL_Cell[Cell_X, 1], EXL_Cell[Cell_X, 256]).Delete(1);
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell行(縦)を削除
        /// </summary>
        /// <param name="Cell_Y"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Dele_Virtical(int Cell_Y)
        {
            Cell_Change(1, 1);
            EXL_Cell.get_Range(EXL_Cell[1, Cell_Y], EXL_Cell[65536, Cell_Y]).Delete(2);
        }
        //セルの削除*****************************************************************************************************


        //セルの移動*****************************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell位置移動
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Change(int X, int Y)
        {
            EXL_Cell = (Excel.Range)EXL_Sheet.Cells[X, Y];
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell複数選択
        /// </summary>
        /// <param name="Start_X"></param>
        /// <param name="Start_Y"></param>
        /// <param name="End_X"></param>
        /// <param name="End_Y"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Multi_Select(int Start_X, int Start_Y, int End_X, int End_Y)
        {
            Cell_Change(1, 1);
            EXL_Cell.get_Range(EXL_Cell[Start_X, Start_Y], EXL_Cell[End_X, End_Y]);
        }
        //セルの移動*****************************************************************************************************


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell指定範囲の色を変える
        /// </summary>
        /// <param name="Start_X"></param>
        /// <param name="Start_Y"></param>
        /// <param name="End_X"></param>
        /// <param name="End_Y"></param>
        /// <param name="Color_Code"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Multiple_Color_Change(int Start_X, int Start_Y, int End_X, int End_Y, int Color_Code)
        {
            Cell_Change(1, 1);
            EXL_Cell.get_Range(EXL_Cell[Start_X, Start_Y], EXL_Cell[End_X, End_Y]).Interior.ColorIndex = Color_Code;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell指定範囲に同じ文字列を記入
        /// </summary>
        /// <param name="Start_X"></param>
        /// <param name="Start_Y"></param>
        /// <param name="End_X"></param>
        /// <param name="End_Y"></param>
        /// <param name="STR"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Multiple_Write(int Start_X, int Start_Y, int End_X, int End_Y, string STR)
        {
            EXL_Cell.get_Range(EXL_Cell[Start_X, Start_Y], EXL_Cell[End_X, End_Y]).Value2 = STR;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定文字列のあるCellを検索
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="Holizontal1_or_Virtical2"></param>
        /// <param name="Serch_STR"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Cell_Text_Serch(int X, int Y, int Holizontal1_or_Virtical2, string Serch_STR)
        {
            int count = 0;
            //指定文字が見つかるまでループ
            while (true)
            {
                if (Holizontal1_or_Virtical2 == 1)
                {
                    Cell_Change(X, Y + count);
                    if (EXL_Cell.Value2 == null)
                    {
                        return Y + count;
                    }
                    else if (EXL_Cell.Value2.ToString() == Serch_STR)
                    {
                        return Y + count;
                    }
                    count++;
                }
                if (Holizontal1_or_Virtical2 == 2)
                {
                    Cell_Change(X + count, Y);
                    if (EXL_Cell.Value2 == null)
                    {
                        return X + count;
                    }
                    else if (EXL_Cell.Value2.ToString() == Serch_STR)
                    {
                        return X + count;
                    }
                    count++;
                }
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Cell書き込み
        /// </summary>
        /// <param name="STR"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Write(string STR)
        {
            EXL_Cell.Value2 = STR;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// シート選択
        /// </summary>
        /// <param name="No"></param>
        public void Sheet_Change(int No)
        {
            EXL_Sheet = (Excel.Worksheet)EXL_App.Worksheets[No];
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// シート数取得
        /// ※注意※　これを実行するとプロセスが残ってしまう。
        /// </summary>
        /// <param name="F_P_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public int Sheet_figure(string F_P_N)
        {
            Excel.Application Ap = null;
            Excel.Workbooks Bk = null;
            //Excel.Worksheet St = null;

            //Excel起動
            Ap = new Excel.Application();
            //Excel表示
            Ap.Visible = false;
            Ap.DisplayAlerts = false;
            //Excelブック取得
            Bk = Ap.Workbooks;
            Bk.Open(F_P_N,  　//　指定ファイルPath
                                Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
                                false, // （省略可能）ReadOnly (True / False )
                                Type.Missing, // （省略可能）Format
                // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                // 5:なし / 6:引数 Delimiterで指定された文字
                                Type.Missing, // （省略可能）Password
                                Type.Missing, // （省略可能）WriteResPassword
                                Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
                                Type.Missing, // （省略可能）Origin
                                Type.Missing, // （省略可能）Delimiter
                                Type.Missing, // （省略可能）Editable
                                Type.Missing, // （省略可能）Notify
                                Type.Missing, // （省略可能）Converter
                                Type.Missing, // （省略可能）AddToMru
                                Type.Missing, // （省略可能）Local
                                Type.Missing  // （省略可能）CorruptLoad
                                );

            int figure = Ap.Worksheets.Count;

            //■Excelファイルを閉じる
            Bk.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Bk);
            Ap.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Ap);

            //プロセス強制終了(プロセスが残ってしまった時の為　※但しなぜか２度目以降は効かない。)
            //System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("EXCEL.EXE");
            //foreach (System.Diagnostics.Process p in ps) p.Kill(); 

            return figure;

        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル幅変更
        /// </summary>
        /// <param name="I"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Width(int I)
        {
            //EXL_Cell.RowHeight = X;
            EXL_Cell.ColumnWidth = I;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル高さ変更
        /// </summary>
        /// <param name="I"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Hight(int I)
        {
            EXL_Cell.RowHeight = I;
            //EXL_Cell.ColumnWidth = Y;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セルフォントサイズ
        /// </summary>
        /// <param name="Size"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Font_Size(int Size)
        {
            EXL_Cell.Font.Size = Size;
        }


        //セルの文字位置横****************************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル横位置を左揃え
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Horizontal_Left()
        {
            EXL_Cell.HorizontalAlignment = Excel.Constants.xlLeft;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル横位置を中央揃え
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Horizontal_Center()
        {
            EXL_Cell.HorizontalAlignment = Excel.Constants.xlCenter;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル横位置を右揃え
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Horizontal_Right()
        {
            EXL_Cell.HorizontalAlignment = Excel.Constants.xlRight;
        }
        //セルの文字位置横****************************************************************************************************



        //セルの文字位置縦***************************************************************************************************
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル縦位置を上部
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Vertical_Top()
        {
            EXL_Cell.VerticalAlignment = Excel.Constants.xlTop;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル縦位置を中央部
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Vertical_Center()
        {
            EXL_Cell.VerticalAlignment = Excel.Constants.xlCenter;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル縦位置を下部
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Vertical_Bottom()
        {
            EXL_Cell.VerticalAlignment = Excel.Constants.xlBottom;
        }
        //セルの文字位置縦***************************************************************************************************


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// セル変更
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Cell_Hight_Width(int X, int Y)
        {
            EXL_Cell.RowHeight = X;
            EXL_Cell.ColumnWidth = Y;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Excelファイルの作成
        /// </summary>
        /// <param name="F_P_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool File_Create(string F_P_N)
        {
            //Excel.Application Ap = new Excel.Application();
            //Excel.Workbooks Bk = Ap.Workbooks.Add();

            Excel.Application Ap = new Excel.Application();
            Excel.Workbooks Bk = Ap.Workbooks;
            Bk.Add(string.Empty);
            Ap.ActiveWorkbook.Save();
            //Ap.Save(F_P_N);

            Bk.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Bk);
            Ap.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Ap);


            return true;
            //Excel.Worksheet St = null;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Excelの保存
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool EXL_Save()
        {
            EXL_App.ActiveWorkbook.Save();
            return true;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// Excelファイルを閉じる
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void EXL_Close()
        {



            System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Cell);
            //ETC.wait(1000);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Sheet);

            //■Excelファイルを閉じる
            //EXL_Cell.Clear();
            //ETC.wait(1000);
            //EXL_Sheet.Application.Quit();
            //ETC.wait(1000);
            EXL_Book.Close();
            //ETC.wait(1000);
            EXL_App.Quit();
            //ETC.wait(1000);

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Cell);
            //ETC.wait(1000);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Sheet);
            //ETC.wait(1000);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Book);
            //ETC.wait(1000);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_App);
            //ETC.wait(1000);

            EXL_Cell = null;
            EXL_Sheet = null;
            EXL_Book = null;
            EXL_App = null;


            //プロセス強制終了(プロセスが残ってしまった時の為　※但しなぜか２度目以降は効かない。)
            System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("EXCEL.EXE");
            //foreach (System.Diagnostics.Process p in ps) p.Kill();
            
            //ETC.wait(1000);
            //ps = null;
            //★★★確実に「EXCEL.EXE」を終了させるために、「dllhost.exe」もプロセス終了させる。
            //ps = System.Diagnostics.Process.GetProcessesByName("dllhost.exe");
            //ETC.wait(1000);
            //ps = null;

            //上記だけではプロセスが消えなかったため、調査してこの一行を実行で解決
            GC.Collect();
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 動作異常終了時のファイルを閉じる処理
        /// </summary>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void EXL_ERR_Close()
        {
            if (EXL_Cell != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Cell);
                EXL_Cell = null;
            }

            if (EXL_Sheet != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Sheet);
                EXL_Sheet = null;
            }

            if (EXL_Book != null)
            {
                EXL_Book.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_Book);
                EXL_Book = null;
            }
            if (EXL_App != null)
            {
                EXL_App.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EXL_App);
                EXL_App = null;
            }


            //プロセス強制終了(プロセスが残ってしまった時の為　※但しなぜか２度目以降は効かない。)
            System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("EXCEL.EXE");
            foreach (System.Diagnostics.Process p in ps) p.Kill();

            //ETC.wait(1000);
            //ps = null;
            //★★★確実に「EXCEL.EXE」を終了させるために、「dllhost.exe」もプロセス終了させる。
            //ps = System.Diagnostics.Process.GetProcessesByName("dllhost.exe");
            //ETC.wait(1000);
            //ps = null;
        }
    }
}
