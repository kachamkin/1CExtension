using EasyHook;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Timer = System.Timers.Timer;

namespace exptcp
{

    [Guid("F80C68DA-1DE1-4C3F-B1F6-D4228C5ADE69")]
    public interface IExpTcp
    {
        bool StartTCPListener(string uid, int port, out string errorMessage);
        
        void StopTCPListener();
        
        bool SendMessage(string ip, int port, string message, out string errorMessage);
        
        void StartTimer(int delay);
        
        void StopTimer();
        
        void EnableHook(string uid);
        
        void DisableHook();
        //bool Execute(string code, string path, out string errorMessage);
        
        bool Encrypt(string plainText, string key, out string iv, out string cipherText, out string errorMessage);

        bool Decrypt(string cipherText, string key, string iv, out string plainText, out string errorMessage);

        bool GenerateKey(string otherPublicKey, out string key, out string errorMessage);

        bool GeneratePublicKey(out string key, out string errorMessage);

    }

    [Guid("E02109C9-EAB8-4608-AAC9-FDE97FA86AC8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExpTcpEvent
    {
        [DispId(1)]
        void OnDataReceived(string uid, string message);

        [DispId(2)]
        void OnTimerElapsed(DateTime time);

        [DispId(3)]
        void OnGetClipboard(string uid, string fileName, string filePath);
    }

    [Guid("A8971C43-510C-4A0F-A51A-E379D6612367")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IExpTcpEvent))]
    [ProgId("1C.Extensions")]
    public class ExpTcp : IExpTcp
    {

        public delegate void DataReceivedEvent(string uid, string message);
        public event DataReceivedEvent OnDataReceived;

        public delegate void TimerElapsedEvent(DateTime time);
        public event TimerElapsedEvent OnTimerElapsed;

        public delegate void GetClipboardEvent(string uid, string fileName, string filePath);
        public event GetClipboardEvent OnGetClipboard;

        private Timer timer;
        private TcpListener listener;

        private LocalHook hook;
        private bool hookEnabled;

        private string formUID;
        private string formUIDListener;

        private ECDiffieHellmanCng ecdh;

        public bool Encrypt(string plainText, string key, out string iv, out string cipherText, out string errorMessage)
        {
            errorMessage = "";
            iv = "";
            cipherText = "";

            try
            {
                using AesCryptoServiceProvider aes = new();
                aes.Key = Convert.FromBase64String(key);
                iv = Convert.ToBase64String(aes.IV);

                using MemoryStream ms = new();
                using CryptoStream cs = new(ms, aes.CreateEncryptor(), CryptoStreamMode.Write);

                byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
                Task.Run(() => cs.Write(plainTextBytes, 0, plainTextBytes.Length)).Wait();
                cs.Close();

                cipherText = Convert.ToBase64String(ms.ToArray());

                ms.Close();
            }
            catch (Exception ex)
            {
                errorMessage = ex.ToString();
                return false;
            }

            return true;
        }

        public bool Decrypt(string cipherText, string key, string iv, out string plainText, out string errorMessage)
        {
            errorMessage = "";
            plainText = "";

            try
            {
                using AesCryptoServiceProvider aes = new();
                aes.Key = Convert.FromBase64String(key);
                aes.IV = Convert.FromBase64String(iv);

                using MemoryStream ms = new();
                using CryptoStream cs = new(ms, aes.CreateDecryptor(), CryptoStreamMode.Write);

                byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
                Task.Run(() => cs.Write(cipherTextBytes, 0, cipherTextBytes.Length)).Wait();
                cs.Close();

                plainText = Encoding.UTF8.GetString(ms.ToArray());

                ms.Close();
            }
            catch (Exception ex)
            {
                errorMessage = ex.ToString();
                return false;
            }

            return true;
        }

        public bool GenerateKey(string otherPublicKey, out string key, out string errorMessage)
        {
            errorMessage = "";
            key = "";

            try
            {
                if (ecdh is null)
                    ecdh = new();
                key = Convert.ToBase64String(ecdh.DeriveKeyMaterial(CngKey.Import(Convert.FromBase64String(otherPublicKey), CngKeyBlobFormat.EccPublicBlob)));
            }
            catch (Exception ex)
            {
                errorMessage = ex.ToString();
                return false;
            }

            return true;
        }

        public bool GeneratePublicKey(out string key, out string errorMessage)
        {
            errorMessage = "";
            key = "";

            try
            {
                if (ecdh is null)
                    ecdh = new();
                key = Convert.ToBase64String(ecdh.Key.Export(CngKeyBlobFormat.EccPublicBlob));
            }
            catch (Exception ex)
            {
                errorMessage = ex.ToString();
                return false;
            }

            return true;
        }

        //public bool Execute(string code, string path, out string errorMessage)
        //{
        //    errorMessage = "";

        //    try
        //    {
        //        ScriptState state = ExecuteAsync(string.IsNullOrWhiteSpace(code) ? File.ReadAllText(path) : code).Result;
        //        if (state.Exception is not null)
        //        {
        //            errorMessage = state.Exception.ToString();
        //            return false;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        errorMessage = ex.ToString();
        //        return false;
        //    }
            
        //    return true;
        //}

        //async Task<ScriptState> ExecuteAsync(string code)
        //{
        //    ScriptOptions options = ScriptOptions.Default
        //       .AddImports("System", "System.IO", "System.Collections.Generic",
        //           "System.Console", "System.Diagnostics", "System.Dynamic",
        //           "System.Linq", "System.Text",
        //           "System.Threading.Tasks", "System.Windows.Forms")
        //       .AddReferences("System", "System.Core", "Microsoft.CSharp", "System.Windows.Forms");

        //    return await CSharpScript.RunAsync(code, options);
        //}

        public void StartTimer(int delay)
        {
            timer = new Timer(delay);
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }

        public void StopTimer()
        {
            try
            {
                timer?.Stop();
                timer?.Dispose();
            }
            catch { };
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            OnTimerElapsed?.Invoke(e.SignalTime);
        }

        public bool StartTCPListener(string uid, int port, out string errorMessage)
        {

            errorMessage = "";
            formUIDListener = uid;

            try
            {

                listener = new TcpListener(IPAddress.Any, port);
                listener.Start();

                Accept();

            }
            catch (Exception ex)
            {
                errorMessage = ex.ToString();
                return false;
            }

            return true;

        }

        public void StopTCPListener()
        {
            try
            {
                listener?.Stop();
            }
            catch { };
        }

        async private void Accept()
        {
            try
            {
                using TcpClient client = await listener.AcceptTcpClientAsync();
                using NetworkStream stream = client.GetStream();
                using BinaryReader reader = new(stream);

                try
                {
                    OnDataReceived?.Invoke(formUIDListener, await Task.Run(() => reader.ReadString()));
                }
                catch { }
                finally
                {
                    reader.Close();
                    stream.Close();
                    client.Close();
                    GC.Collect();
                }
            }
            catch { };
            Accept();
        }

        public bool SendMessage(string ip, int port, string message, out string errorMessage)
        {

            errorMessage = "";

            try
            {
                SendMessageAsync(ip, port, message);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.ToString();
                return false;
            }

        }

        async private void SendMessageAsync(string ip, int port, string message)
        {
            using TcpClient client = new();
            await client.ConnectAsync(ip, port);

            using NetworkStream stream = client.GetStream();
            using BinaryWriter writer = new(stream);

            await Task.Run(() => writer.Write(message));

            writer.Close();
            stream.Close();
            client.Close();

            GC.Collect();
        }

        public void EnableHook(string uid)
        {
            hookEnabled = true;
            formUID = uid;
            hook.ThreadACL.SetExclusiveACL(new Int32[] { });
        }

        public void DisableHook()
        {
            hookEnabled = false;
            hook.ThreadACL.SetInclusiveACL(new Int32[] { });
        }

        private void RunHook()
        {
            try
            {

                hook = LocalHook.Create(EasyHook.LocalHook.GetProcAddress("ole32.dll", "OleGetClipboard"),
                            new NativeMethods.GetClipboardDelegate(GetClipboardHook), null);
            }
            catch
            {
            }
        }

        private long GetClipboardHook(out NativeMethods.IDataObject pDataObj)
        {
            try
            {
                long res = NativeMethods.OleGetClipboard(out pDataObj);

                if (hookEnabled && (DataObjectHelper.GetDataPresent(pDataObj, "FileGroupDescriptorW") || DataObjectHelper.GetDataPresent(pDataObj, "FileGroupDescriptor")))
                {
                    string[] fileNames = DataObjectHelper.GetFilenames(pDataObj);

                    string tempPath = Path.GetTempPath() + "1CExtension\\";
                    if (!Directory.Exists(tempPath))
                        Directory.CreateDirectory(tempPath);

                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        using FileStream stream = new(tempPath + fileNames[i], FileMode.Create);
                        DataObjectHelper.ReadFileContents(pDataObj, i, stream);

                        long streamLength = stream.Length;
                        stream.Close();

                        if (streamLength > 0)
                            OnGetClipboard?.Invoke(formUID, fileNames[i], stream.Name);
                    }

                    Clipboard.Clear();
                }

                return res;
            }
            catch
            {
                pDataObj = null;
                return 1;
            }
        }

        public ExpTcp()
        {
            hookEnabled = false;
            RunHook();
        }

        ~ExpTcp()
        {
            try
            {
                StopTCPListener();
                StopTimer();
                ecdh?.Dispose();

                if (hook != null)
                {
                    hook.ThreadACL.SetInclusiveACL(new Int32[] { });
                    hook.Dispose();
                };
                LocalHook.Release();
            }
            catch { };
        }

    }
}
