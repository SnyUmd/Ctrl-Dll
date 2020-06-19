//#define Debug

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Windows.Forms;
using System.IO.Compression;



namespace Ctrl_Dll
{
    public class cls_File_Cont
    {
        //ファイルを開く際に使用するための宣言
        public System.Diagnostics.Process App = null;


        //▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲
        //▼▲▼▲クラス宣言▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲
        //▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲▼▲
        //Excel_Cont EX = new Excel_Cont();
        //File_Cont FC = new File_Cont();
        //Etc1 E1 = new Etc1();

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        //音楽を再生する
        [System.Runtime.InteropServices.DllImport
                ("winmm.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int mciSendString
                (string command, System.Text.StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

        public string aliasName = "MediaFile";

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 音楽ファイルを再生する
        /// </summary>
        /// <param name="File_Dir"></param>
        /// <param name="Play_Time"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Music_Play(string File_Dir, int Play_Time)
        {
            //再生するファイル名
            string fileName = File_Dir;

            string cmd;
            //ファイルを開く
            cmd = "open \"" + fileName + "\" alias " + aliasName;
            if (mciSendString(cmd, null, 0, IntPtr.Zero) != 0)
                return;
            //再生する
            cmd = "play " + aliasName;
            mciSendString(cmd, null, 0, IntPtr.Zero);


            System.Threading.Thread.Sleep(Play_Time);

            //string cmd;
            //再生しているWAVEを停止する
            cmd = "stop " + aliasName;
            mciSendString(cmd, null, 0, IntPtr.Zero);
            //閉じる
            cmd = "close " + aliasName;
            mciSendString(cmd, null, 0, IntPtr.Zero);
        }



        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// デスクトップのディレクトリを取得
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Desk_Top_Directory()
        {
            return System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\";
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// マイドキュメントのディレクトリを取得
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Mydocument_Directory()
        {
            //string test = System.Environment.SpecialFolder.Personal.ToString();
            return System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\";
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 実行ファイルDirectoryを取得
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string App_Directory_Acquisition()
        {
            return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\";
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイルフォルダー名取得
        /// </summary>
        /// <param name="F_D_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Get_Folder_Name(string F_D_N)
        {
            return Path.GetDirectoryName(F_D_N);
            /*
            int File_Len;
            int Extension_Point;
            string Folder_Name;

            //文字数を取得
            File_Len = F_D_N.Length;
            //拡張子先頭位置を取得
            Extension_Point = F_D_N.LastIndexOf(@"\") + 1;
            //拡張子を削除
            Folder_Name = F_D_N.Remove(Extension_Point, File_Len - Extension_Point);
            return Folder_Name;*/
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル名取得(拡張子無し)
        /// </summary>
        /// <param name="F_D_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Get_File_Name(string F_D_N)
        {
            return Path.GetFileNameWithoutExtension(F_D_N);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル名取得(拡張子含む)
        /// </summary>
        /// <param name="F_D_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string File_Name_Extraction(string F_D_N)
        {
            return Path.GetFileName(F_D_N);

            /*
            int File_Len;
            int File_Name_Point;
            string File_Name;

            //文字数を取得
            File_Len = F_D_N.Length;
            //拡ファイル名位置を取得
            File_Name_Point = F_D_N.LastIndexOf(@"\") + 1;
            //ファイル名部のみ取得
            File_Name = F_D_N.Substring(File_Name_Point, File_Len - File_Name_Point);
            return File_Name;
            */
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダーを開く
        /// </summary>
        /// <param name="F_D_N"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Folder_Open(string F_D_N)
        {
            System.Diagnostics.Process.Start(F_D_N);
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// テキストファイルの読み込み
        /// </summary>
        /// <param name="TXT_F_D_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Txt_File_Read(string TXT_F_D_N)
        {
            string A = "";
            bool File_Have;

            File_Have = File_Search(TXT_F_D_N);

            if (!File_Have)
            {
                A = "";
            }
            else
            {
                System.IO.StreamReader B = new System.IO.StreamReader(TXT_F_D_N, System.Text.Encoding.GetEncoding("shift_jis"));
                A = B.ReadToEnd();
                B.Close();
                B.Dispose();
            }
            return A;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// テキストファイルの中身を書き込み
        /// </summary>
        /// <param name="File_Dir"></param>
        /// <param name="W_Str"></param>
        /// <param name="Overwrite"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Txt_File_Write(string File_Dir, string W_Str, bool Overwrite)
        {
            System.IO.StreamWriter sw =
                new System.IO.StreamWriter(File_Dir, false, System.Text.Encoding.GetEncoding("shift_jis"));
            if (Overwrite)
            {

            }
            //テキストを内容を書き込む
            sw.Write(W_Str);
            //閉じる
            sw.Close();
            sw.Dispose();

        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル作成
        /// </summary>
        /// <param name="File_Path_Name"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void File_Create(string File_Path_Name)
        {
            // F_Stream が破棄されることを保証するために using を使用する
            // 指定したパスのファイルを作成する
            using (System.IO.FileStream F_Stream = System.IO.File.Create(File_Path_Name))
            {
                // 作成時に返される FileStream を利用して閉じる
                if (F_Stream != null)
                {
                    F_Stream.Close();
                }
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル検索
        /// </summary>
        /// <param name="File_Dir_Name"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool File_Search(string File_Dir_Name)
        {
            int File_Name_Point;
            int File_Name_Length;

            string File_Name;
            string File_Dir;

            string[] All_File;

            //指定されたパス&ファイル名から、検索するファイルのフォルダパスまでの文字位置を取得
            File_Name_Length = File_Dir_Name.Length;
            //指定されたパス&ファイル名から、検索するファイル名の先頭位置を取得(末尾から検索する)
            File_Name_Point = File_Dir_Name.LastIndexOf(@"\") + 1;

            //検索するパス名を抽出
            File_Dir = File_Dir_Name.Substring(0, File_Name_Point);
            //検索するファイル名を抽出
            File_Name = File_Dir_Name.Substring(File_Name_Point, File_Name_Length - File_Name_Point);

            //検索するパス名中に指定ファイル名のファイルがいくつあるかを検索実行
            All_File = System.IO.Directory.GetFiles(File_Dir, File_Name, System.IO.SearchOption.TopDirectoryOnly);

            //もしもファイルが見つからなければ、falseを返す
            //ファイルが見つかれば、trueを返す
            if (All_File.Length > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダ内のファイル検索
        /// </summary>
        /// <param name="File_Dir"></param>
        /// <param name="Serch_File_Name"></param>
        /// <param name="All_Dir"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[] Folder_File_Serch(string File_Dir, string Serch_File_Name, bool All_Dir)
        {
            string[] All_File;

            //指定されたパス&ファイル名から、検索するファイルのフォルダパスまでの文字位置を取得
            //File_Name_Length = File_Dir_Name.Length;
            //指定されたパス&ファイル名から、検索するファイル名の先頭位置を取得(末尾から検索する)
            //File_Name_Point = File_Dir_Name.LastIndexOf(@"\") + 1;

            //検索するパス名を抽出
            //File_Dir = File_Dir_Name.Substring(0, File_Name_Point);
            //検索するファイル名を抽出
            //File_Name = File_Dir_Name.Substring(File_Name_Point, File_Name_Length - File_Name_Point);

            //検索するパス名中に指定ファイル名のファイルがいくつあるかを検索実行
            if (All_Dir) All_File = System.IO.Directory.GetFiles(File_Dir, Serch_File_Name, System.IO.SearchOption.AllDirectories);
            else All_File = System.IO.Directory.GetFiles(File_Dir, Serch_File_Name, System.IO.SearchOption.TopDirectoryOnly);

            return All_File;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダの作成
        /// </summary>
        /// <param name="Dir"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Folder_Create(string Dir)
        {
            //フォルダ"C:\TEST\SUB"を作成する
            //"C:\TEST"フォルダが存在しなくても"C:\TEST\SUB"が作成される
            //"C:\TEST\SUB"が存在していると、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生
            try
            {
                System.IO.DirectoryInfo Di =
                    System.IO.Directory.CreateDirectory(Dir);
                Di = null;
                return true;
            }
            catch
            {
                return false;
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダの削除
        /// </summary>
        /// <param name="Dir"></param>
        /// <param name="Sub_Folder_Delete_EN"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Folder_Delete(string Dir, bool Sub_Folder_Delete_EN)
        {
            //フォルダ"C:\TEST"を削除する
            //第2項をTrueにすると、"C:\TEST"を根こそぎ（サブフォルダ、ファイルも）削除する
            //"C:\TEST"に読み取り専用ファイルがあると、UnauthorizedAccessExceptionが発生
            //"C:\TEST"が存在しないと、DirectoryNotFoundExceptionが発生
            try
            {
                System.IO.Directory.Delete(Dir, Sub_Folder_Delete_EN);
                return true;
            }
            catch
            {
                return false;
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダの移動
        /// </summary>
        /// <param name="Dir"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Folder_Move(string strBeforeDir, string strAfterDir)
        {
            //フォルダ"C:\1"を"C:\2\SUB"に移動（名前を変更）する
            //"C:\2\SUB"が存在していると、IOExceptionが発生
            //移動先が別のドライブ（ボリューム）だと、IOExceptionが発生
            //"C:\1"や"C:\2"が存在しないと、DirectoryNotFoundExceptionが発生
            //"C:\1\SUB"のように移動先が移動元のサブフォルダだと、IOExceptionが発生
            //アクセス許可が無いと、UnauthorizedAccessExceptionが発生
            try
            {
                System.IO.Directory.Move(strBeforeDir, strAfterDir);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// サブフォルダを取得する
        /// </summary>
        /// <param name="str_target_folder"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[] mGetSubFolder(string str_target_folder)
        {
            return System.IO.Directory.GetDirectories(str_target_folder, "*", System.IO.SearchOption.AllDirectories);
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダにサブフォルダがあるかを検索する
        /// </summary>
        /// <param name="str_target_folder"></param>
        /// <param name="str_serch_folder"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mSerchFolder(string str_target_folder, string str_serch_folder)
        {
            string[] ary_strBuf;
            ary_strBuf = System.IO.Directory.GetDirectories(str_target_folder, str_serch_folder, System.IO.SearchOption.AllDirectories);
            if (ary_strBuf.Length < 1)
                return false;
            else
                return true;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// フォルダの存在を確認
        /// </summary>
        /// <param name="Dir"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool Folder_Fined(string Dir)
        {
            // フォルダ (ディレクトリ) が存在しているかどうか確認する
            if (System.IO.Directory.Exists(Dir)) return true;
            //if (System.IO.File.Exists(Dir)) return true;
            else return false;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイルの存在を確認
        /// </summary>
        /// <param name="Dir"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool File_Fined(string Dir)
        {
            // フォルダ (ディレクトリ) が存在しているかどうか確認する
            if (System.IO.File.Exists(Dir)) return true;
            //if (System.IO.File.Exists(Dir)) return true;
            else return false;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル名の変更/移動
        /// </summary>
        /// <param name="Before_F_P_N"></param>
        /// <param name="After_F_P_N"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void File_ReName(string Before_F_P_N, string After_F_P_N)
        {
            File.Move(Before_F_P_N, After_F_P_N);
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイルコピー
        /// </summary>
        /// <param name="Ref_F_N_P"></param>
        /// <param name="Copy_F_N_P"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool File_Copy(string Ref_F_N_P, string Copy_F_N_P)
        {



            //★★★フォルダが無ければ作成する処理があったほうが良い★★★
            //★★★ファイル選択キャンセルされた時のエラー処理が無い★★★



            if (!File_Search(Copy_F_N_P))
            {
                System.IO.File.Copy(Ref_F_N_P, Copy_F_N_P, false);
                return true;
            }
            else
            {
                return false;
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル削除
        /// </summary>
        /// <param name="F_N_P"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool File_Del(string F_N_P)
        {
            try
            {
                System.IO.File.Delete(F_N_P);
                return true;
            }
            catch
            {
                return false;
            }
        }



        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル拡張子取得
        /// </summary>
        /// <param name="F_P_N"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string File_Extension(string F_P_N)
        {
            return Path.GetExtension(F_P_N);
            /*
            int File_Len;
            int Extension_Point;

            File_Len = F_P_N.Length;
            Extension_Point = F_P_N.LastIndexOf(".");
            if (Extension_Point < 0) return "ERR";
            return F_P_N.Substring(Extension_Point, File_Len - Extension_Point);
             */
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル拡張子を変更後のテキストを返す
        /// </summary>
        /// <param name="F_P_N"></param>
        /// <param name="Change_Extension"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string File_Extension_Change(string F_P_N, string Change_Extension)
        {
            //Path.ChangeExtension(F_P_N, Change_Extension);


            int File_Len;
            int Extension_Point;

            string Change_F_P_N;

#if Debug
            if (0 > Change_Extension.IndexOf(".", 0))
            {
                MessageBox.Show("変更拡張子にはドットを入れる。プログラムを確認");
            }
#endif

            //文字数を取得
            File_Len = F_P_N.Length;
            //拡張子先頭位置を取得
            Extension_Point = F_P_N.LastIndexOf(".");
            //拡張子を削除
            Change_F_P_N = F_P_N.Remove(Extension_Point, File_Len - Extension_Point);
            //目的拡張子を追加
            Change_F_P_N += Change_Extension;

            return Change_F_P_N;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル選択ダイアログ操作
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Protel_CSVFile_Read_Dialog()
        {
            //クラス宣言
            OpenFileDialog Read_File_Dialog = new OpenFileDialog();
            //初期表示フォルダの指定
            Read_File_Dialog.InitialDirectory = Desk_Top_Directory();
            //表示ファイルを指定
            Read_File_Dialog.Filter = "CSVファイル(*.CSV) | *.CSV;|すべてのファイル|*.*";
            //ダイアログ表示初期時に、BOMが選択されるようにする。
            Read_File_Dialog.FilterIndex = 1;
            //タイトル設定
            Read_File_Dialog.Title = "CSVファイルを選択してください。";
            if (Read_File_Dialog.ShowDialog() == DialogResult.OK)
            {
                return Read_File_Dialog.FileName;
            }
            else
            {
                return "";
            }
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル選択ダイアログ操作
        /// </summary>
        /// <param name="F_D"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string Address_ExcelFile_Read_Dialog(string F_D)
        {
            //クラス宣言
            OpenFileDialog Read_File_Dialog = new OpenFileDialog();
            //初期表示フォルダの指定
            Read_File_Dialog.InitialDirectory = F_D;
            //表示ファイルを指定
            Read_File_Dialog.Filter = "xlsファイル(*.xls;*.xlsx) | *.xls;*.xlsx;|すべてのファイル|*.*";
            //ダイアログ表示初期時に、Excelが選択されるようにする。
            Read_File_Dialog.FilterIndex = 1;
            //タイトル設定
            Read_File_Dialog.Title = "アドレス表のファイルを選択してください。";
            if (Read_File_Dialog.ShowDialog() == DialogResult.OK)
            {
                return Read_File_Dialog.FileName;
            }
            else
            {
                return "";
            }
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル選択ダイアログ操作　全てのファイル
        /// </summary>
        /// <param name="F_D"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string File_Read_Dialog(string F_D)
        {
            //クラス宣言
            OpenFileDialog Read_File_Dialog = new OpenFileDialog();
            //初期表示フォルダの指定
            Read_File_Dialog.InitialDirectory = F_D;
            //表示ファイルを指定
            Read_File_Dialog.Filter = "すべてのファイル|*.*";
            //ダイアログ表示初期時に、Excelが選択されるようにする。
            Read_File_Dialog.FilterIndex = 1;
            //タイトル設定
            Read_File_Dialog.Title = "ファイルを選択してください。";
            if (Read_File_Dialog.ShowDialog() == DialogResult.OK)
            {
                return Read_File_Dialog.FileName;
            }
            else
            {
                return "";
            }
        }



        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// 指定ファイルの起動
        /// </summary>
        /// <param name="File_Name"></param>
        /// <param name="Wait_For_Exit"></param>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public void Ex_App_Start(string File_Name, bool Wait_For_Exit)
        {
            //ファイルを開いて終了まで待機する
            App = System.Diagnostics.Process.Start(File_Name);
            //if(Wait_For_Exit) App.WaitForExit();
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル更新日取得 Datetime
        /// </summary>
        /// <param name="File_Name"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public DateTime File_Renewal_Date_DateTime(string File_Name)
        {
            return System.IO.File.GetLastWriteTime(File_Name);
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// ファイル更新日取得 String
        /// </summary>
        /// <param name="File_Name"></param>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string File_Renewal_Date_String(string File_Name)
        {
            return System.IO.File.GetLastWriteTime(File_Name).ToString("F");

            //("F")⇒⇒⇒'2000年5月12日 20:30:15
            //("gyyyy年MM月dd日(dddd)")⇒⇒⇒'西暦2000年05月12日(金曜日)
            //("tthh時mm分ss秒fffミリ秒")⇒⇒⇒'午後08時30分15秒123ミリ秒
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mTextFileSave(string strWrite, ref string strFolderPath, bool blUpdateFolder)
        {
            //bool blReturn = false;

            try
            {
                //SaveFileDialogクラスのインスタンスを作成
                SaveFileDialog sfd = new SaveFileDialog();


                //はじめのファイル名を指定する
                //はじめに「ファイル名」で表示される文字列を指定する
                sfd.FileName = "新しいファイル.txt";
                //はじめに表示されるフォルダを指定する
                //sfd.InitialDirectory = Mydocument_Directory();
                sfd.InitialDirectory = strFolderPath;
                //[ファイルの種類]に表示される選択肢を指定する
                //指定しない（空の文字列）の時は、現在のディレクトリが表示される
                sfd.Filter = "テキストファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*";
                //[ファイルの種類]ではじめに選択されるものを指定する
                //1番目の「テキストファイル」が選択されているようにする
                sfd.FilterIndex = 1;
                //タイトルを設定する
                sfd.Title = "保存先のファイルを選択してください";
                //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                sfd.RestoreDirectory = true;
                //既に存在するファイル名を指定したとき警告する
                //デフォルトでTrueなので指定する必要はない
                sfd.OverwritePrompt = true;
                //存在しないパスが指定されたとき警告を表示する
                //デフォルトでTrueなので指定する必要はない
                sfd.CheckPathExists = true;

                //ダイアログを表示する
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    //OKボタンがクリックされたとき、選択されたファイル名を表示する
                    //Console.WriteLine(sfd.FileName);

                    //次回開くフォルダが必要であればアップデート
                    if (blUpdateFolder)
                        strFolderPath = Get_Folder_Name(sfd.FileName);

                    StreamWriter sw = new StreamWriter(sfd.FileName, false, Encoding.ASCII);
                    sw.Write(strWrite);
                    sw.Close();
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

//◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string mFileCreation(string str_top_dir)
        {
            //SaveFileDialogクラスのインスタンスを作成
            SaveFileDialog sfd = new SaveFileDialog();
            //はじめに「ファイル名」で表示される文字列を指定する
            sfd.FileName = "NewFile";
            //はじめに表示されるフォルダを指定する
            sfd.InitialDirectory = str_top_dir;


            //[ファイルの種類]に表示される選択肢を指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            sfd.Filter = "テキストファイル・ログファイル(*.txt;*.log)|*.txt;*.log|すべてのファイル(*.*)|*.*";

            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            sfd.FilterIndex = 2;


            //タイトルを設定する
            sfd.Title = "Please select a save destination file";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            sfd.RestoreDirectory = true;
            //既に存在するファイル名を指定したとき警告する
            //デフォルトでTrueなので指定する必要はない
            sfd.OverwritePrompt = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            sfd.CheckPathExists = true;

            //ダイアログを表示する
            if (sfd.ShowDialog() == DialogResult.OK)
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                return sfd.FileName + ".txt";
            else
                return "";
        }




        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        /// <summary>
        /// PCに存在するドライブを取得
        /// </summary>
        /// <returns></returns>
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[] mGetDriveList()
        {
            string[] ary_strDriver;
            ary_strDriver = System.Environment.GetLogicalDrives();

            return ary_strDriver;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public string[] mDriveInf(string str_drive_name)
        {
            string[] strRTN = new string[5];
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            strRTN[0] = str_drive_name;
            strRTN[1] = drive.Name;
            strRTN[2] = drive.DriveType.ToString();
            strRTN[3] = drive.IsReady.ToString();
            if (strRTN[3] == "True")
                strRTN[4] = drive.VolumeLabel.ToLower();
            else
                strRTN[4] = "Not Ready";

            return strRTN;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDriveReady(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            return drive.IsReady;
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDribe_is_Removable(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            if (drive.DriveType == DriveType.Removable)
                return true;
            else
                return false;
        }
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDribe_is_CDRom(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            if (drive.DriveType == DriveType.CDRom)
                return true;
            else
                return false;
        }
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDribe_is_Fixed(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            if (drive.DriveType == DriveType.Fixed)
                return true;
            else
                return false;
        }
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDribe_is_Networkd(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            if (drive.DriveType == DriveType.Network)
                return true;
            else
                return false;
        }
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDribe_is_Ram(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            if (drive.DriveType == DriveType.Ram)
                return true;
            else
                return false;
        }
        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mDribe_is_NoRootDirectory(string str_drive_name)
        {
            System.IO.DriveInfo drive = new System.IO.DriveInfo(str_drive_name);
            if (drive.DriveType == DriveType.NoRootDirectory)
                return true;
            else
                return false;
        }


        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mFileZIP(string str_file, string str_ext_zipfile_name)
        {
            try
            {
                ZipFile.CreateFromDirectory(str_file, str_ext_zipfile_name);
                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }

        //◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆◆
        public bool mFileUnZIP(string str_zip_file, string str_ext_dir)
        {
            try
            {
                ZipFile.ExtractToDirectory(str_zip_file,
                                           str_ext_dir,
                                           System.Text.Encoding.GetEncoding("shift_jis"));
                return true;
            }
            catch(Exception)
            {
                return false;
            }
        }
    }
}




//System.IO.Path.****(@"C:\My Documents\My Pictures\サンプル.jpg");

//GetDirectoryName              ディレクトリ名の取得	        C:\dir\sub	C:\dir\sub	C:\dir	.\sub	\\pc\s\dir
//GetExtension	                拡張子の取得	                .txt				.txt
//GetFileName	                ファイル名の取得	            f.txt		sub	f.	f.txt
//GetFileNameWithoutExtension	ファイル名（拡張子なし）の取得	f		sub	f	f
//GetPathRoot	                ルートディレクトリ名の取得	    C:\	C:\	C:\		\\pc\s
//GetFullPath	                絶対パスの取得	                C:\dir\sub\f.txt	C:\dir\sub\	C:\dir\sub	C:\cd\sub\f	\\pc\s\dir\f.txt
//HasExtension	                拡張子を持っているか	        True	False	False	False	True
//IsPathRooted	                ルートが含まれているか（詳細）	True	True	True	False	True
//Directory.GetDirectoryRoot	ボリューム、ルート情報の取得	C:\	C:\	C:\	C:\	\\pc\s
//Directory.GetParent           親ディレクトリの取得	        C:\dir\sub	C:\dir\sub	C:\dir	C:\cd\sub	\\pc\s\dir