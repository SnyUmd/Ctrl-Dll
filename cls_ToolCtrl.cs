using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ctrl_Dll
{
    public class cls_ToolCtrl
    {
        public void Lsv_Initial_Set(ListView lsv, string str_name)
        {

            // ListViewコントロールのプロパティを設定
            lsv.FullRowSelect = true;
            lsv.GridLines = true;
            lsv.Sorting = SortOrder.Ascending;
            lsv.View = View.Details;
            
        }

        public void mTxb_View_Last(ref TextBox TXB1)
        {
            //カレット位置を末尾に移動
            TXB1.SelectionStart = TXB1.Text.Length;
            //テキストボックスにフォーカスを移動
            TXB1.Focus();
            //カレット位置までスクロール
            TXB1.ScrollToCaret();
        }
    }
}
