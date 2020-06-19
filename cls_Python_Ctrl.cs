using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.IO;

namespace Ctrl_Dll
{
    public class cls_Python_Ctrl
    {
        //***************************************************************************************
        /// <summary>
        /// Pythonモデルの戻り値を取得する
        /// </summary>
        /// <param name="str_py_path">実行するPythonモデルのパス</param>
        /// <returns></returns>
        //***************************************************************************************
        public string mPyRead(string str_py_path)
        {
            string strReturn = "";

            //phthonプロセスをセット
            Process PythonProcess = new Process
            {
                StartInfo = new ProcessStartInfo("python.exe")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    Arguments = str_py_path
                }
            };

            //プロセススタート
            PythonProcess.Start();

            StreamReader S_Reader = PythonProcess.StandardOutput;
            //strReturn = S_Reader.ReadLine();
            strReturn = S_Reader.ReadToEnd();
            PythonProcess.WaitForExit();
            PythonProcess.Close();

            return strReturn;
        }



        //***************************************************************************************
        public string mPyRun(string str_py_interpreter_path, string str_py_script_path)
        {
            string strReturn = "";

            var lst_buf = new List<string>
            {
            str_py_script_path ,
            "10",   //第1引数
            "20"    //第2引数
            };

            var process = new Process()
            {
                StartInfo = new ProcessStartInfo(str_py_interpreter_path)
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    Arguments = string.Join(" ", lst_buf),
                },
            };

            process.Start();

            //python側でprintした内容を取得
            var sr = process.StandardOutput;
            var result = sr.ReadToEnd();

            process.WaitForExit();
            process.Close();

            //Console.WriteLine("Result is ... " + result);
            strReturn = result;

            return strReturn;
        }
    }
}
