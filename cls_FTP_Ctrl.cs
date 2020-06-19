using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Windows.Forms;

//■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
using System.Threading;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
//■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■


namespace Ctrl_Dll
{

    // ================================================================================================================================
    // ================================================================================================================================
    // ================================================================================================================================
    // FTP設定
    public class FtpConfig
    {
        private static string ftpUser;
        private static string ftpPassword;
        private static string ftpHost;
        private static bool enableSslPolicyErrorsRemoteCertificateNameMismatch;

        /*
        public FtpConfig(string sUserID, string sPass, string sFTP_Host)
        {
            // 接続情報
            FtpConfig.FtpUser = sUserID;  // FTPユーザー
            FtpConfig.FtpPassword = sPass; // FTPパスワード
            FtpConfig.ftpHost = sFTP_Host;   // FTPホスト（ドメイン）
            // オプション
            FtpConfig.EnableSslPolicyErrorsRemoteCertificateNameMismatch = false;   // SSL証明書名の不一致をチェックするか？
        }
        */

        
        static FtpConfig()
        {
            // 接続情報
            FtpConfig.FtpUser = "plcuser";  // FTPユーザー
            FtpConfig.FtpPassword = "plc1331"; // FTPパスワード
            FtpConfig.ftpHost = "ftp://54.65.200.126/demo_test/";   // FTPホスト（ドメイン）
            // オプション
            FtpConfig.EnableSslPolicyErrorsRemoteCertificateNameMismatch = false;   // SSL証明書名の不一致をチェックするか？
        }
        

        // FTPユーザー
        public static string FtpUser
        {
            get { return ftpUser; }
            private set { ftpUser = value; }
        }

        // FTPパスワード
        public static string FtpPassword
        {
            get { return ftpPassword; }
            private set { ftpPassword = value; }
        }

        // ホストのルートURI
        public static string FtpRoot
        {
            get { return "ftp://" + ftpHost; }
        }

        // SSL証明書名の不一致をエラーとするか
        public static bool EnableSslPolicyErrorsRemoteCertificateNameMismatch
        {
            get { return enableSslPolicyErrorsRemoteCertificateNameMismatch; }
            set { enableSslPolicyErrorsRemoteCertificateNameMismatch = value; }
        }
    }


    // ================================================================================================================================
    // ================================================================================================================================
    // ================================================================================================================================
    // FTP非同期通信の状態管理
    public class FtpState
    {
        private ManualResetEvent wait;
        string status;
        private FtpWebRequest request;
        private string filePath;
        private Exception operationException = null;

        public FtpState() { wait = new ManualResetEvent(false); }

        // シグナル状態管理
        public ManualResetEvent OperationComplete { get { return wait; } }

        // ステータス記述
        public string StatusDescription
        {
            get { return status; }
            set { status = value; }
        }

        // FTPオブジェクト
        public FtpWebRequest Request
        {
            get { return request; }
            set { request = value; }
        }

        // ローカルファイルパス
        public string FilePath
        {
            get { return filePath; }
            set { filePath = value; }
        }

        // 例外
        public Exception OperationException
        {
            get { return operationException; }
            set { operationException = value; }
        }
    }



    // ================================================================================================================================
    // ================================================================================================================================
    // ================================================================================================================================
    // FTP管理
    public class FtpManager
    {
        //証明書の内容を表示
        private static void PrintCertificate(X509Certificate certificate)
        {
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);

            sw.WriteLine("▼証明書の内容");
            sw.WriteLine("サブジェクトの識別名(Subject)：{0}", certificate.Subject);
            sw.WriteLine("発行した証明機関の名前(Issuer)：{0}", certificate.Issuer);
            sw.WriteLine("形式の名前(GetFormat)：{0}", certificate.GetFormat());
            sw.WriteLine("失効日(GetExpirationDateString)：{0}", certificate.GetExpirationDateString());
            sw.WriteLine("発効日(GetEffectiveDateString)：{0}", certificate.GetEffectiveDateString());
            sw.WriteLine("文字列形式のキー アルゴリズム情報(GetKeyAlgorithm)：{0}", certificate.GetKeyAlgorithm());
            sw.WriteLine("16進数文字列形式の公開キー(GetPublicKeyString)：{0}", certificate.GetPublicKeyString());
            sw.WriteLine("16進数文字列形式のシリアル番号(GetSerialNumberString)：{0}", certificate.GetSerialNumberString());
            sw.WriteLine("▲証明書の内容");

            Console.WriteLine(sb);
        }

        // SSL証明書の信頼性を確認
        private static bool OnRemoteCertificateValidationCallback(
          Object sender,
          X509Certificate certificate,
          X509Chain chain,
          SslPolicyErrors sslPolicyErrors)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                Console.WriteLine("SSL のポリシー エラーはありません");
            }
            else
            {
                if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) ==
                    SslPolicyErrors.RemoteCertificateChainErrors)
                {
                    Console.WriteLine("【ERROR】ChainStatusが、空でない配列を返しました");
                }
                else if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateNameMismatch) ==
                  SslPolicyErrors.RemoteCertificateNameMismatch)
                {
                    // 証明書情報を確認
                    PrintCertificate(certificate);

                    if (!FtpConfig.EnableSslPolicyErrorsRemoteCertificateNameMismatch)
                    {
                        Console.WriteLine("【WARNING】設定により証明書名の不一致を容認しました");
                        return true;
                    }
                    Console.WriteLine("【ERROR】証明書の名前が一致していません");
                }
                else if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateNotAvailable) ==
                  SslPolicyErrors.RemoteCertificateNotAvailable)
                {
                    Console.WriteLine("【ERROR】証明書が利用できません");
                }
                else
                {
                    Console.WriteLine("【ERROR】予期しないエラーが発生しました");
                }

                return false;
            }
            return true;
        }

        // FTPのリクエスト用オブジェクト作成
        private static FtpWebRequest CreateFtpWebRequest(string webRequestMethod, string requestUri)
        {
            FtpWebRequest request = null;
            string uri = FtpConfig.FtpRoot + requestUri;

            try
            {
                // アップロード先URI
                Uri targetUri = new Uri(uri);

                request = (FtpWebRequest)WebRequest.Create(targetUri);
                request.Credentials = new NetworkCredential(FtpConfig.FtpUser, FtpConfig.FtpPassword);
                request.Method = webRequestMethod;
                request.KeepAlive = false;
                request.UseBinary = true;
                request.UsePassive = true;

                // SSL通信設定
                ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(OnRemoteCertificateValidationCallback);
                request.EnableSsl = true;
            }
            catch (UriFormatException ex)
            {
                Console.WriteLine("【ERROR】無効なURIを検出 - {0}, uri：{1}", ex.Message, uri);
                throw;
            }
            return request;
        }

        // アップロード(非同期)
        public static void UploadAsync(string requestUri, string filePath)
        {
            FtpWebRequest request = null;
            try
            {
                // 状態管理クラスをオブジェクト化
                FtpState state = new FtpState();

                // FTPのリクエスト用オブジェクト作成
                request = CreateFtpWebRequest(WebRequestMethods.Ftp.UploadFile, requestUri);

                // 非同期通信管理オブジェクト設定
                state.Request = request;
                state.FilePath = filePath;

                // 待機オブジェクト
                ManualResetEvent waitObject = state.OperationComplete;

                // 非同期リクエスト開始
                request.BeginGetRequestStream(new AsyncCallback(EndGetStreamCallback), state);

                // 処理完了まで待機
                waitObject.WaitOne();

                // 処理完了
                if (state.OperationException != null)
                {
                    // 例外発生
                    throw state.OperationException;
                }
                else
                {
                    Console.WriteLine("アップロード処理が正常に終了しました - {0}", state.StatusDescription);
                }
            }
            catch (WebException ex)
            {
                Console.WriteLine("【ERROR】アップロード処理 - {0}", ex.Message);
                throw;
            }
            finally
            {
                request = null;
            }
        }

        // 非同期通信のコールバック
        private static void EndGetStreamCallback(IAsyncResult ar)
        {
            // 通信状態管理オブジェクトを取得
            FtpState state = (FtpState)ar.AsyncState;
            Stream requestStream = null;
            FileStream stream = null;
            try
            {
                // ファイルをストリームに書き込む
                requestStream = state.Request.EndGetRequestStream(ar);
                byte[] buffer = new byte[1024];
                int count = 0;
                int readBytes = 0;
                stream = File.OpenRead(state.FilePath);
                do
                {
                    readBytes = stream.Read(buffer, 0, buffer.Length);
                    requestStream.Write(buffer, 0, readBytes);
                    count += readBytes;
                }
                while (readBytes != 0);
                Console.WriteLine("{0} byteストリームに書き込みました", count);

                state.Request.BeginGetResponse(new AsyncCallback(EndGetResponseCallback), state);
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("【ERROR】アップロード処理 ファイルがありません。 - {0}", ex.Message);
                state.OperationException = ex;
                state.OperationComplete.Set();
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("【ERROR】アップロード処理 - {0}", ex.Message);
                state.OperationException = ex;
                state.OperationComplete.Set();
                return;
            }
            finally
            {
                if (stream != null) stream.Close();
                if (requestStream != null) requestStream.Close();
            }
        }

        // 非同期通信完了処理
        private static void EndGetResponseCallback(IAsyncResult ar)
        {
            FtpState state = (FtpState)ar.AsyncState;
            FtpWebResponse response = null;

            try
            {
                response = (FtpWebResponse)state.Request.EndGetResponse(ar);
                state.StatusDescription = response.StatusDescription;
                state.OperationComplete.Set();
            }
            catch (Exception ex)
            {
                Console.WriteLine("【ERROR】アップロード処理 - {0}", ex.Message);
                state.OperationException = ex;
                state.OperationComplete.Set();
            }
            finally
            {
                if (response != null) response.Close();
            }
        }

        // ファイルを削除
        public static void Delete(string requestUri)
        {
            FtpWebRequest request = null;
            FtpWebResponse response = null;

            try
            {
                // リクエスト用オブジェクト作成
                request = CreateFtpWebRequest(WebRequestMethods.Ftp.DeleteFile, requestUri);
                // レスポンス用オブジェクト取得
                response = (FtpWebResponse)request.GetResponse();
                Console.WriteLine("削除処理が正常に終了しました - {0}", response.StatusDescription);
            }
            catch (WebException ex)
            {
                Console.WriteLine("【ERROR】削除処理 - {0}", ex.Message);
                throw;
            }
            finally
            {
                if (response != null) response.Close();
                request = null;
            }
        }

        /* ダウンロード */
        public static void Download(string requestUri, string filePath)
        {
            FtpWebRequest request = null;
            FtpWebResponse response = null;
            Stream responseStream = null;
            FileStream stream = null;
            try
            {
                // リクエスト用オブジェクト作成
                request = CreateFtpWebRequest(WebRequestMethods.Ftp.DownloadFile, requestUri);
                // レスポンス用オブジェクト取得
                response = (FtpWebResponse)request.GetResponse();

                // ストリームに書き込み
                responseStream = response.GetResponseStream();
                stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                byte[] buffer = new byte[1024];
                while (true)
                {
                    int readSize = responseStream.Read(buffer, 0, buffer.Length);
                    if (readSize == 0) break;
                    stream.Write(buffer, 0, readSize);
                }
                Console.WriteLine("ダウンロード処理が正常に終了しました - {0}", response.StatusDescription);
            }
            catch (WebException ex)
            {
                Console.WriteLine("【ERROR】ダウンロード処理 - {0}", ex.Message);
                throw;
            }
            finally
            {
                if (responseStream != null) responseStream.Close();
                if (stream != null) stream.Close();
                if (response != null) response.Close();
                request = null;
            }
        }

        /* ディレクトリ存在チェック */
        public static bool ExistsDirectory(string requestUri)
        {
            FtpWebRequest request = null;
            FtpWebResponse response = null;

            try
            {
                // リクエスト用オブジェクト作成
                request = CreateFtpWebRequest(WebRequestMethods.Ftp.ListDirectory, requestUri);
                // レスポンス用オブジェクト取得
                response = (FtpWebResponse)request.GetResponse();
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    FtpWebResponse r = (FtpWebResponse)ex.Response;
                    if (r.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                    {
                        Console.WriteLine("ディレクトリ無");
                        return false;
                    }
                }
                Console.WriteLine("【ERROR】ディレクトリ存在チェック処理 - {0}", ex.Message);
                throw;
            }
            finally
            {
                if (response != null) response.Close();
                request = null;
            }

            Console.WriteLine("ディレクトリ有");
            return true;
        }

        /* ディレクトリを作成 */
        public static void MakeDirectory(string requestUri)
        {
            FtpWebRequest request = null;
            FtpWebResponse response = null;

            try
            {
                // 既にディレクトリが存在していたら何もしない
                if (ExistsDirectory(requestUri)) return;
                // リクエスト用オブジェクト作成
                request = CreateFtpWebRequest(WebRequestMethods.Ftp.MakeDirectory, requestUri);
                // レスポンス用オブジェクト取得
                response = (FtpWebResponse)request.GetResponse();
                Console.WriteLine("ディレクトリ作成処理が正常に終了しました - {0}", response.StatusDescription);
            }
            catch (WebException ex)
            {
                Console.WriteLine("【ERROR】ディレクトリを作成処理 - {0}", ex.Message);
                throw;
            }
            finally
            {
                if (response != null) response.Close();
                request = null;
            }
        }
    }


    
    // ================================================================================================================================
    // 製作部=========================================================================================================================
    // ================================================================================================================================
    public class cls_FTP_Ctrl
    {
        //**********************************************************************************************
        /// <summary>
        /// FTPサーバへファイルをアップロード
        /// </summary>
        /// <param name="s_uri">FTPサーバアドレス</param>
        /// <param name="s_user_id">ユーザーID</param>
        /// <param name="s_pass">パスワード</param>
        /// <param name="s_file_name">ファイルネーム</param>
        //**********************************************************************************************
        public void mFTP_FileUp(string s_uri, string s_user_id, string s_pass, string s_file_name)
        {
            //FtpWebRequestの作成
            System.Net.FtpWebRequest ftpReq = (System.Net.FtpWebRequest)System.Net.WebRequest.Create(s_uri);
            //IDとパスワードを設定
            ftpReq.Credentials = new System.Net.NetworkCredential(s_user_id, s_pass);
            //MethodにWebRequestMethods.Ftp.UploadFile("STOR")を設定
            ftpReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile;
            //要求の完了後に接続を閉じる
            ftpReq.KeepAlive = false;
            //ASCIIモードで転送する
            //ftpReq.UseBinary = false;
            //PASVモードを無効にする
            //ftpReq.UsePassive = false;//ファイルをアップロードするためのStreamを取得
            System.IO.Stream reqStrm = ftpReq.GetRequestStream();
            //アップロードするファイルを開く
            System.IO.FileStream Fs = new System.IO.FileStream(s_file_name, 
                                                               System.IO.FileMode.Open, 
                                                               System.IO.FileAccess.Read);

            //アップロードStreamに書き込む
            byte[] buffer = new byte[1024];

            while (true)
            {
                int readSize = Fs.Read(buffer, 0, buffer.Length);
                if (readSize == 0)
                    break;
                reqStrm.Write(buffer, 0, readSize);
            }
            Fs.Close();
            reqStrm.Close(); 

            //FtpWebResponseを取得
            System.Net.FtpWebResponse FtpRes =
                        (System.Net.FtpWebResponse)ftpReq.GetResponse();
            //FTPサーバーから送信されたステータスを表示
            Console.WriteLine("{0}: {1}", FtpRes.StatusCode, FtpRes.StatusDescription);
            //閉じる
            FtpRes.Close();
        }
    }
}



//**********************************************************************************************
//**********************************************************************************************
//**********************************************************************************************