#define Debug

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#if Debug
using System.Windows.Forms;
#endif

namespace Ctrl_Dll
{
    public class cls_List_Ctrl
    {

        //public List<string> GetArryItem_Str(string[,] arry_str, int n_element)
        //*******************************************************************************
        /// <summary>
        /// 配列の指定行の内容をリスト化(昇順並び)
        /// </summary>
        /// <param name="arry_str"></param>
        /// <param name="n_dimensions"></param>
        /// <returns></returns>
        //*******************************************************************************
        public List<string> GetArryItem_Str(string[,] arry_str, int n_dimensions)
        {
            //ループ数を取得
            //int n_cnt_loop = arry_str.Length / arry_str.GetLength(n_dimensions);
            List<string> lst_str_Return = new List<string>();

            lst_str_Return.Add("");

            //for (int i = 0; i < arry_str.GetLength(n_dimensions); i++)
            for (int i = 0; i < arry_str.GetLength(1); i++)
            {
                if (lst_str_Return.IndexOf(arry_str[n_dimensions, i]) < 0)
                    lst_str_Return.Add(arry_str[n_dimensions, i]);
            }
            lst_str_Return.Sort();
            return lst_str_Return;
        }

        //配列数取得方法」
        public void list_count()
        {
            int[] array1 = { 1, 2, 3 };
            int[,] array2 = new int[2, 10];

            MessageBox.Show("１次元配列の要素数 = " + array1.Length.ToString());

            MessageBox.Show("多次元配列の次元数 = " + array2.Rank.ToString());
            MessageBox.Show("多次元配列の総要素数 = " + array2.Length.ToString());

            MessageBox.Show("２次元配列の１次元目の要素数 = " + array2.GetLength(0).ToString());
            MessageBox.Show("２次元配列の２次元目の要素数 = " + array2.GetLength(1).ToString());
        }
    }
}
