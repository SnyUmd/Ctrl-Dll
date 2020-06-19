//**************************************************************************************
//
//**************************************************************************************



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms;
using System.Drawing;

namespace Ctrl_Dll
{

    public class cls_Chart_Drawing
    {

        cls_Text_Cont clsTC = new cls_Text_Cont();

        //**************************************************************************************
        //チャートをクリア
        //**************************************************************************************
        public void mChartClear(ref Chart chrt)
        {
            chrt.Series.Clear();
            chrt.ChartAreas.Clear();
            chrt.Titles.Clear();
        }


        //**************************************************************************************
        /// <summary>
        /// グラフエリアを追加
        /// </summary>
        /// <param name="chrt"></param>
        /// <param name="str_area_name"></param>
        //**************************************************************************************
        public void mAddChartArea(ref Chart chrt,
                              string str_area_name)
        {
            chrt.ChartAreas.Add(new ChartArea(str_area_name));
        }


        //**************************************************************************************
        /// <summary>
        /// グラフを追加
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_add_area_no"グラフを追加するエリア番号</param>
        /// <param name="type">グラフタイプ</param>
        /// <param name="str_legend_name">グラフ名</param>
        /// <param name="bl_marker_circle">ポイントにドットをプロットする？</param>
        //**************************************************************************************
        public void mAddSeries(ref Chart chrt,
                               int i_add_area_no,
                               System.Windows.Forms.DataVisualization.Charting.SeriesChartType type,
                               string str_legend_name,
                               bool bl_marker_circle)
        {
            chrt.Series.Add(str_legend_name);
            chrt.Series[str_legend_name].ChartType = type;
            chrt.Series[str_legend_name].ChartArea = chrt.ChartAreas[i_add_area_no].Name;

            if (bl_marker_circle)
                chrt.Series[str_legend_name].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle;
        }


        //**************************************************************************************
        /// <summary>
        /// 目盛スタイルをセット
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_area_no">エリア番号</param>
        /// <param name="x_color">X軸目盛の色</param>
        /// <param name="y_color">Y軸目盛の色</param>
        /// <param name="x_dash_style">X軸目盛のスタイル</param>
        /// <param name="y_dath_style">Y軸目盛のスタイル</param>
        //**************************************************************************************
        public void mSetGrid(ref Chart chrt,
                             int i_area_no,
                             Color x_color,
                             Color y_color,
                             ChartDashStyle x_dash_style,
                             ChartDashStyle y_dash_style)
        {
            chrt.ChartAreas[i_area_no].AxisX.MajorGrid.LineColor = x_color;
            chrt.ChartAreas[i_area_no].AxisX.MajorGrid.LineDashStyle = x_dash_style;
            chrt.ChartAreas[i_area_no].AxisY.MajorGrid.LineColor = y_color;
            chrt.ChartAreas[i_area_no].AxisY.MajorGrid.LineDashStyle = y_dash_style;
        }


        //**************************************************************************************
        /// <summary>
        /// X軸メモリのスケールをセット
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_area_no">セットするエリア番号</param>
        /// <param name="i_max">X軸目盛の最大値</param>
        /// <param name="i_min">X軸目盛の最小値</param>
        /// <param name="i_interval">X軸区切り軸</param>
        //**************************************************************************************
        public void mSetScaleValue_X(ref Chart chrt, int i_area_no, int i_max, int i_min, int i_interval)
        {
            chrt.ChartAreas[i_area_no].AxisX.Maximum = i_max;
            chrt.ChartAreas[i_area_no].AxisX.Minimum = i_min;
            chrt.ChartAreas[i_area_no].AxisX.Interval = i_interval;
        }


        //**************************************************************************************
        /// <summary>
        /// Y軸目盛のスケールをセット
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_area_no">セットするエリア番号</param>
        /// <param name="i_max">Y軸目盛の最大値</param>
        /// <param name="i_min">Y軸目盛の最小値</param>
        /// <param name="i_interval">Y軸区切り軸</param>
        //**************************************************************************************
        public void mSetScaleValue_Y(ref Chart chrt, int i_area_no, int i_max, int i_min, int i_interval)
        {
            chrt.ChartAreas[i_area_no].AxisY.Maximum = i_max;
            chrt.ChartAreas[i_area_no].AxisY.Minimum = i_min;
            chrt.ChartAreas[i_area_no].AxisY.Interval = i_interval;
        }


        //**************************************************************************************
        /// <summary>
        /// データを配列化
        /// </summary>
        /// <param name="str_data">データ文字列</param>
        /// <param name="str_data_separator">データの区切り文字</param>
        /// <param name="str_set_separator">セット区切り文字</param>
        /// <returns></returns>
        //**************************************************************************************
        //★OR
        public string[,] mDataSet(string str_data, string str_data_separator, string str_set_separator, int i_data_num)
        //public string[,] mDataSet(string str_data, string str_data_separator, string str_set_separator, int i_data_num, Form FRM, ref TextBox TXB1, ref TextBox TXB2)
        {
            int iSetSeparator;
            int iDataSeparator;

            int iSetNum = 0;


            string strBuf_OneData;
            string strBuf_OneSet;
            string strDataBuf = str_data;



            //★OR
            //double[,] ary_dblReturn;
            string[,] ary_strReturn;

            //データのセット数を取得
            iSetSeparator = clsTC.Character_Figure(str_data, str_set_separator);

            //TXB1.Text = iSetSeparator.ToString();

            if (iSetSeparator <= 0)
            {
                MessageBox.Show("エラー　セット区切り文字が見つかりません");
                return null;
            }
            else
            {
                //★OR
                //ary_dblReturn = new double[i_data_num + 1, iSetSeparator + 1];
                ary_strReturn = new string[i_data_num, iSetSeparator + 1];

                //1セット内にデータ区切が何文字あるかを取得
                //iDataSeparator = clsTC.Character_Figure(strBuf_OneSet, str_data_separator);
                //配列数をセット

                for (double i = 0; i <= iSetSeparator; i++)
                {
                    //1セットのデータをバッファに格納
                    strBuf_OneSet = clsTC.Secified_Char_Read(strDataBuf, str_set_separator, 0);/////

                    if (i != iSetSeparator)//ループの最後は削除処理はスキップ
                                           //格納した1セットをデータ文字列から削除
                        strDataBuf = strDataBuf.Remove(0, strBuf_OneSet.Length + str_set_separator.Length);/////

                    //１セットのデータをチェック
                    iDataSeparator = clsTC.Character_Figure(strBuf_OneSet, str_data_separator);///////
                    //デーセパレータ数が一致しなければスキップ
                    if (i_data_num > (iDataSeparator + 1)) continue;



                    //最後セットのデータ数が少ない可能性があるので、毎回1セットのデータ数を取得
                    //iDataSeparator = clsTC.Character_Figure(strBuf_OneSet, str_data_separator);

                    //1セットのデータをデータ毎に配列
                    for (int i2 = 0; i2 < i_data_num; i2++)
                    {
                        strBuf_OneData = clsTC.Secified_Char_Read(strBuf_OneSet, str_data_separator, 0);

                        //★OR
                        //ary_dblReturn[i2, i] = double.Parse(strBuf_OneData);
                        ary_strReturn[i2, iSetNum] = strBuf_OneData;

                        if (i2 != i_data_num - 1)//ループの最後は削除処理はスキップ
                            //取得したデータを文字列から削除
                            strBuf_OneSet = strBuf_OneSet.Remove(0, strBuf_OneData.Length + str_data_separator.Length);
                    }

                    //TXB2.Text = iSetNum.ToString();
                    //FRM.Refresh();
                    iSetNum++;
                }
            }

            //★OR
            //return ary_dblReturn;
            return ary_strReturn;
        }










        /*
        //**************************************************************************************
        /// <summary>
        /// 配列データを平均化
        /// </summary>
        /// <param name="ary_str"></param>
        /// <param name="i_element_count"></param>
        /// <param name="i_sensor_num"></param>
        /// <param name="str_data_separator"></param>
        /// <param name="str_set_Separator"></param>
        /// <returns></returns>
        public string mMovingAverageData(string[,] ary_str, int i_element_count, int i_sensor_num, string str_data_separator, string str_set_Separator)
        {
            string strReturn = "";
            string buf;

            int iAryLineLen;
            int iAryColLen;

            double dblCalculation = 0;
            int iSkipNum = 0;


            iAryColLen = ary_str.GetLength(0);
            iAryLineLen = ary_str.GetLength(1);

            if (iAryColLen < i_sensor_num)
            {
                MessageBox.Show("指定されたセンサ数が、データ無いのセンサ数を超えています。");
                return null;
            }

            for (int i = 0; i < iAryLineLen - i_element_count; i++)
            {
                for (int i2 = 0; i2 < i_sensor_num; i2++)
                {
                    for (int i3 = 0; i3 < i_element_count; i3++)
                    {
                        buf = ary_str[i2, (i + i3)];
                        //中身が数字で無ければスキップ
                        if (!clsTC.mStrIsNum_Judge(buf))
                        {
                            iSkipNum++;
                            continue;
                        }
                        dblCalculation += double.Parse(buf);
                    }
                    strReturn += dblCalculation / (i_element_count - iSkipNum);
                    if (i2 < i_sensor_num - 1)
                        strReturn += str_data_separator;

                    iSkipNum = 0;
                    dblCalculation = 0;
                }
                strReturn += str_set_Separator;
            }

            return strReturn;
        }
        */




        //**************************************************************************************
        /// <summary>
        /// グラフデータ書き込み
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_series_no">グラフ番号</param>
        /// <param name="db_x_data">X軸データ</param>
        /// <param name="db_y_data">Y軸データ</param>
        //**************************************************************************************
        public void mDataWrite2(Form frm, ref Chart chrt, int i_series_no, int iX, int iY, bool bl_form_refresh)
        {
            chrt.Series[i_series_no].Points.AddXY(iX, iY);
            if (bl_form_refresh)
                frm.Refresh();
        }




















        //**************************************************************************************
        /// <summary>
        /// 配列データをグラフへ書き込み
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_series_no">グラフ番号</param>
        /// <param name="db_x_data">X軸データ</param>
        /// <param name="db_y_data">Y軸データ</param>
        //**************************************************************************************
        public void mDataWrite_ary(ref Chart chrt, int i_series_no, double[] dbl_data)
        {
            int i = 0;
            foreach (double data in dbl_data)
            {
                chrt.Series[i_series_no].Points.Add(i, dbl_data[i]);

                i++;
            }
        }


        //**************************************************************************************
        /// <summary>
        /// ２軸の配列データをグラフデータ書き込み
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_series_no">グラフ番号</param>
        /// <param name="db_x_data">X軸データ</param>
        /// <param name="db_y_data">Y軸データ</param>
        //**************************************************************************************
        public void mDataWrite_aryXY(ref Chart chrt, int i_series_no, double[] db_x_data, double[] db_y_data)
        {
            int i = 0;
            foreach (double x_data in db_x_data)
            {
                chrt.Series[i_series_no].Points.AddXY(i, db_y_data[i]);

                i++;
            }
        }


        //**************************************************************************************
        /// <summary>
        /// ２軸のデータをグラフデータ書き込み
        /// </summary>
        /// <param name="chrt">チャートコントローラ</param>
        /// <param name="i_series_no">グラフ番号</param>
        /// <param name="db_x_data">X軸データ</param>
        /// <param name="db_y_data">Y軸データ</param>
        //**************************************************************************************
        public void mDataWrite_XY(ref Chart chrt, int i_series_no, double[] db_x_data, double[] db_y_data)
        {
            int i = 0;
            foreach (double x_data in db_x_data)
            {
                chrt.Series[i_series_no].Points.AddXY(i, db_y_data[i]);

                i++;
            }
        }

        //**************************************************************************************
        //public void mChartOffset(ref Chart CHR, int i_series_num, int i_axisX_max, List<int> lst_i_Data, int i_offset, Form frm)
        public void mChartOffset(ref Chart CHR, int i_series_num, List<int> lst_i_Data, int i_offset, Form frm)
        {
            int iBuf = 0;

            CHR.Series[i_series_num].Points.Clear();
            //for (int i = 0; i < i_axisX_max - 1; i++)
            for (int i = 0; i < lst_i_Data.Count - 1; i++)
                {
                iBuf = lst_i_Data[i] + i_offset;

                mDataWrite2(frm, ref CHR, i_series_num, i, iBuf, false);
                iBuf = 0;
            }
        }
    }
}






//**************************************************************************************
//
//**************************************************************************************