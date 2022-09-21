//using System;
//using System.Collections;
//using System.ComponentModel;
//using System.Reflection;
//using System.Runtime.InteropServices;
//using SYSKIND = System.Runtime.InteropServices.ComTypes.SYSKIND;

//namespace exptcp
//{
//    //[ComImport,
//    //GuidAttribute("00020406-0000-0000-C000-000000000046"),
//    //InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
//    //ComVisible(false)]
//    //public interface ICreateITypeLib
//    //{
//    //    void CreateTypeInfo();
//    //    void SetName();
//    //    void SetVersion();
//    //    void SetGuid();
//    //    void SetDocString();
//    //    void SetHelpFileName();
//    //    void SetHelpContext();
//    //    void SetLcid();
//    //    void SetLibFlags();
//    //    void SaveAllChanges();
//    //}

//    //public class ConversionEventHandler : ITypeLibExporterNotifySink
//    //{
//    //    public void ReportEvent(ExporterEventKind eventKind, int eventCode, string eventMsg)
//    //    {
//    //    }

//    //    public Object ResolveRef(Assembly asm)
//    //    {
//    //        return null;
//    //    }
//    //}

//    [RunInstaller(true)]
//    public partial class ExpTcpInstaller : System.Configuration.Install.Installer
//    {
//        //[DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
//        //public static extern long RegisterTypeLib([MarshalAs(UnmanagedType.Interface)] object ptlib, [MarshalAs(UnmanagedType.LPWStr)] string szFullPath, [MarshalAs(UnmanagedType.LPWStr)] string szHelpDir);
//        //[DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
//        //public extern static long UnRegisterTypeLib(ref Guid libID, ushort wVerMajor, ushort wVerMinor, int lcid, SYSKIND syskind);

//        public ExpTcpInstaller()
//        {
//            InitializeComponent();
//        }

//        protected override void OnAfterInstall(IDictionary savedState)
//        {
//            base.OnAfterInstall(savedState);

//            Assembly asm = typeof(ExpTcp).Assembly;
//            new RegistrationServices().RegisterAssembly(asm, AssemblyRegistrationFlags.SetCodeBase);

//            //string appFolder = base.Context.Parameters["targetdir"];
//            //appFolder = appFolder.Substring(0, appFolder.Length - 1) + asm.GetName().Name + ".tlb";

//            //ICreateITypeLib tlb = (ICreateITypeLib)new TypeLibConverter().ConvertAssemblyToTypeLib(asm, appFolder, Environment.Is64BitOperatingSystem ? TypeLibExporterFlags.ExportAs64Bit : TypeLibExporterFlags.ExportAs32Bit, new ConversionEventHandler());
//            //tlb.SaveAllChanges();
//            //RegisterTypeLib(tlb, appFolder, null);
//        }

//        protected override void OnBeforeUninstall(IDictionary savedState)
//        {
//            base.OnBeforeUninstall(savedState);

//            Assembly asm = typeof(ExpTcp).Assembly;
//            new RegistrationServices().UnregisterAssembly(asm);

//            //Version version = asm.GetName().Version;
//            //Guid guid = new(((GuidAttribute)asm.GetCustomAttributes(typeof(GuidAttribute), false)[0]).Value);
//            //UnRegisterTypeLib(ref guid, (ushort)version.Major, (ushort)version.Minor, 0, Environment.Is64BitOperatingSystem ? SYSKIND.SYS_WIN64 : SYSKIND.SYS_WIN32);
//        }
//    }
//}
