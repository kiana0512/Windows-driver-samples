// Program.cs — MyCompany APO Control (.NET 9 + WinForms + IKsControl)
// 在当前 app 目录 dotnet new winforms 后，删除 Form1.*，用本文件覆盖 Program.cs 即可运行。

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

internal static class KsGuids
{
    // 与 ApoCtl.cpp 保持一致：属性集 GUID + PID（包括一次性下发的 PID_ParamsBlob=10）
    public static readonly Guid MYCOMPANY_APO_PROPSETID = new Guid("d4d9a040-8b5f-4c0e-aad1-aabbccddeeff");
    public const uint PID_Gain = 1, PID_EQBand = 2, PID_Reverb = 3, PID_Limiter = 4, PID_ParamsBlob = 10;

    public const uint KSPROPERTY_TYPE_GET = 0x00000001;
    public const uint KSPROPERTY_TYPE_SET = 0x00000002;
}

// =======================
// FIX: 顶层公共 KSPROPERTY，IKsControl 与调用方统一用它
// =======================
[StructLayout(LayoutKind.Sequential)]
public struct KSPROPERTY
{
    public Guid Set;
    public uint Id;
    public uint Flags;
}

// ---------- 模型（与 MyApoParams.h 对齐：12 段 EQ + Reverb + Limiter + 512B Opcode） ----------
public enum EqType : int { Peak = 0, LowShelf = 1, HighShelf = 2 }

[TypeConverter(typeof(ExpandableObjectConverter))]
public class EqBandModel
{
    public bool Enabled { get; set; } = false;
    public float Freq { get; set; } = 1000f;
    public float Q { get; set; } = 1.0f;
    public float GainDb { get; set; } = 0f;
    public EqType Type { get; set; } = EqType.Peak;
    public override string ToString() => $"{(Enabled ? "On" : "Off")} {Type} {Freq}Hz Q={Q} {GainDb:+0.0;-0.0;0}dB";
}

[TypeConverter(typeof(ExpandableObjectConverter))]
public class ReverbModel
{
    public bool Enabled { get; set; } = false;
    public float Wet { get; set; } = 0.2f;
    public float Room { get; set; } = 0.7f;
    public float Damp { get; set; } = 0.3f;
    public float PreMs { get; set; } = 20f;
    public override string ToString() => $"{(Enabled ? "On" : "Off")} wet={Wet:0.00} room={Room:0.00} damp={Damp:0.00} pre={PreMs}ms";
}

[TypeConverter(typeof(ExpandableObjectConverter))]
public class DspParamsModel
{
    [Category("Main")] public float Gain { get; set; } = 1.0f;

    [Category("EQ (12 bands)")] public EqBandModel Eq0  { get; set; } = new EqBandModel { Type = EqType.LowShelf,  Freq = 100,   Q = 0.707f };
    [Category("EQ (12 bands)")] public EqBandModel Eq1  { get; set; } = new EqBandModel { Type = EqType.Peak,      Freq = 250,   Q = 1.0f  };
    [Category("EQ (12 bands)")] public EqBandModel Eq2  { get; set; } = new EqBandModel { Type = EqType.Peak,      Freq = 500,   Q = 1.0f  };
    [Category("EQ (12 bands)")] public EqBandModel Eq3  { get; set; } = new EqBandModel { Type = EqType.Peak,      Freq = 1000,  Q = 1.0f  };
    [Category("EQ (12 bands)")] public EqBandModel Eq4  { get; set; } = new EqBandModel { Type = EqType.Peak,      Freq = 2000,  Q = 1.0f  };
    [Category("EQ (12 bands)")] public EqBandModel Eq5  { get; set; } = new EqBandModel { Type = EqType.Peak,      Freq = 4000,  Q = 1.0f  };
    [Category("EQ (12 bands)")] public EqBandModel Eq6  { get; set; } = new EqBandModel { Type = EqType.Peak,      Freq = 8000,  Q = 1.0f  };
    [Category("EQ (12 bands)")] public EqBandModel Eq7  { get; set; } = new EqBandModel { Type = EqType.HighShelf, Freq = 12000, Q = 0.707f };
    [Category("EQ (12 bands)")] public EqBandModel Eq8  { get; set; } = new EqBandModel();
    [Category("EQ (12 bands)")] public EqBandModel Eq9  { get; set; } = new EqBandModel();
    [Category("EQ (12 bands)")] public EqBandModel Eq10 { get; set; } = new EqBandModel();
    [Category("EQ (12 bands)")] public EqBandModel Eq11 { get; set; } = new EqBandModel();

    [Category("FX")] public ReverbModel Reverb { get; set; } = new ReverbModel();
    [Category("FX")] public bool LimiterEnabled { get; set; } = true;

    [Browsable(false)] public byte[] Opcode { get; set; } = Array.Empty<byte>();

    public IEnumerable<EqBandModel> Bands()
    {
        yield return Eq0; yield return Eq1; yield return Eq2; yield return Eq3; yield return Eq4; yield return Eq5;
        yield return Eq6; yield return Eq7; yield return Eq8; yield return Eq9; yield return Eq10; yield return Eq11;
    }
}

// 二进制打包（pack(1) 对齐到 784 字节）
internal static class ParamsPacker
{
    public static byte[] Pack(DspParamsModel m)
    {
        // 4 + 12*20 + 20 + 4 + 4 + 512 = 784 bytes
        byte[] buf = new byte[784];
        int off = 0;
        void W32(int v)   { BitConverter.GetBytes(v).CopyTo(buf, off); off += 4; }
        void WU32(uint v) { BitConverter.GetBytes(v).CopyTo(buf, off); off += 4; }
        void WF(float f)  { BitConverter.GetBytes(f).CopyTo(buf, off); off += 4; }

        WF(m.Gain);
        foreach (var b in m.Bands()) { W32(b.Enabled ? 1 : 0); WF(b.Freq); WF(b.Q); WF(b.GainDb); W32((int)b.Type); }
        W32(m.Reverb.Enabled ? 1 : 0); WF(m.Reverb.Wet); WF(m.Reverb.Room); WF(m.Reverb.Damp); WF(m.Reverb.PreMs);
        W32(m.LimiterEnabled ? 1 : 0);

        var op = (m.Opcode ?? Array.Empty<byte>()).Take(512).ToArray();
        WU32((uint)op.Length);
        Array.Copy(op, 0, buf, off, op.Length);
        return buf;
    }
}

// -------------------- WinForms UI --------------------
public class MainForm : Form
{
    ComboBox cboDev = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 420 };
    Button btnRefresh = new Button { Text = "刷新设备" };
    Button btnForce = new Button { Text = "强制拉流(150ms)" };
    PropertyGrid grid = new PropertyGrid { Dock = DockStyle.Fill };
    Button btnApply = new Button { Text = "Apply (KsProperty)", Dock = DockStyle.Bottom, Height = 40 };

    List<MMDevice> devices = new();
    DspParamsModel model = new();

    public MainForm()
    {
        Text = "MyCompany APO Control (IKsControl + 12-band EQ)";
        Width = 900; Height = 700;

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40 };
        top.Controls.Add(new Label { Text = "Render 设备：", AutoSize = true, Padding = new Padding(0, 10, 0, 0) });
        top.Controls.Add(cboDev);
        top.Controls.Add(btnRefresh);
        top.Controls.Add(btnForce);
        Controls.Add(top);

        grid.SelectedObject = model;
        Controls.Add(grid);
        Controls.Add(btnApply);

        btnRefresh.Click += (_, __) => RefreshDevices();
        btnForce.Click += (_, __) =>
        {
            var dev = CurrentDevice(); if (dev == null) return;
            using var ac = dev.ActivateAudioClient();
            ac.InitializeSharedNoPersist(48000, 2, 150); // 触发 audiodg 懒加载 APO
        };
        btnApply.Click += (_, __) =>
        {
            var dev = CurrentDevice();
            if (dev == null) { MessageBox.Show("请选择设备"); return; }
            byte[] blob = ParamsPacker.Pack(model);
            using var ks = dev.OpenKsControl(forceStream: true, forceMs: 150);
            ks.SetBlob(blob);
            MessageBox.Show($"参数已下发：{blob.Length} bytes (PID=10)");
        };

        Load += (_, __) => RefreshDevices();
    }

    void RefreshDevices()
    {
        devices = CoreAudio.ListActiveEndpoints(DataFlow.Render);
        cboDev.Items.Clear();
        foreach (var d in devices) cboDev.Items.Add(d.FriendlyName);
        if (cboDev.Items.Count > 0) cboDev.SelectedIndex = 0;
    }
    MMDevice? CurrentDevice() => (cboDev.SelectedIndex >= 0 && cboDev.SelectedIndex < devices.Count) ? devices[cboDev.SelectedIndex] : null;
}

// -------------------- CoreAudio + IKsControl 互操作封装 --------------------
public enum DataFlow { Render = 0, Capture = 1 }

public sealed class MMDevice : IDisposable
{
    internal CoreAudio.IMMDevice dev;
    public string FriendlyName { get; }
    internal MMDevice(CoreAudio.IMMDevice d) { dev = d; FriendlyName = CoreAudio.GetDeviceFriendlyName(d); }
    public void Dispose() { if (dev != null) Marshal.ReleaseComObject(dev); }

    public AudioClient ActivateAudioClient() => new AudioClient(CoreAudio.ActivateAudioClient(dev));
    public KsControl OpenKsControl(bool forceStream = false, int forceMs = 150)
    {
        if (forceStream) { using var ac = ActivateAudioClient(); ac.InitializeSharedNoPersist(48000, 2, forceMs); }
        var ks = CoreAudio.GetKsControl(dev);
        return new KsControl(ks);
    }
}

public sealed class AudioClient : IDisposable
{
    internal CoreAudio.IAudioClient ac;
    public AudioClient(CoreAudio.IAudioClient a) { ac = a; }

    public void InitializeSharedNoPersist(int sampleRate, int channels, int runMs)
    {
        IntPtr pwfx;
        int hr = ac.GetMixFormat(out pwfx);
        if (hr >= 0 && pwfx != IntPtr.Zero)
        {
            // 用设备混音格式初始化（pFormat 为 CoTaskMem 指针）
            hr = ac.Initialize(0 /*shared*/, 0x80000 /*NOPERSIST*/, 10_000_000, 0, pwfx, IntPtr.Zero);
            Marshal.FreeCoTaskMem(pwfx);
        }
        else
        {
            // FIX: Initialize 的 pFormat 是 IntPtr，不能 ref 结构；改为分配非托管内存传指针
            var fmt = CoreAudio.WaveFormatEx.S16Stereo48k();
            IntPtr pFmt = Marshal.AllocHGlobal(Marshal.SizeOf<CoreAudio.WAVEFORMATEX>()); // 分配内存
            try
            {
                Marshal.StructureToPtr(fmt, pFmt, false);
                hr = ac.Initialize(0, 0x80000, 10_000_000, 0, pFmt, IntPtr.Zero);
            }
            finally
            {
                Marshal.FreeHGlobal(pFmt);
            }
        }
        if (hr >= 0) { ac.Start(); System.Threading.Thread.Sleep(runMs); ac.Stop(); }
    }

    public void Dispose() { if (ac != null) Marshal.ReleaseComObject(ac); }
}

public sealed class KsControl : IDisposable
{
    internal CoreAudio.IKsControl ks;
    public KsControl(CoreAudio.IKsControl k) { ks = k; }
    public void Dispose() { if (ks != null) Marshal.ReleaseComObject(ks); }

    public void SetBlob(byte[] bytes)
    {
        var prop = new KSPROPERTY { Set = KsGuids.MYCOMPANY_APO_PROPSETID, Id = KsGuids.PID_ParamsBlob, Flags = KsGuids.KSPROPERTY_TYPE_SET };
        uint ret = 0;
        int hr = ks.KsProperty(ref prop, (uint)Marshal.SizeOf<KSPROPERTY>(), bytes, (uint)bytes.Length, ref ret);
        if (hr < 0) throw new COMException("KsProperty(SET BLOB) failed", hr);
    }
}

// =======================
// FIX: CoreAudio 设为 public 静态类；内部仅把对 IMMDevice 作为参数的方法设为 internal
// =======================
public static class CoreAudio
{
    // 注意：static readonly Guid 不能 by-ref 传；调用处用“本地变量副本”解决（见下方 ListActiveEndpoints）
    static readonly Guid CLSID_MMDeviceEnumerator = new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E");
    static readonly Guid IID_IMMDeviceEnumerator = typeof(IMMDeviceEnumerator).GUID;

    public static List<MMDevice> ListActiveEndpoints(DataFlow df)
    {
        var list = new List<MMDevice>();

        // FIX: 不能 ref 传 static readonly 字段 —— 拷到本地变量再 by-ref
        Guid clsid = CLSID_MMDeviceEnumerator;
        Guid iidEnum = IID_IMMDeviceEnumerator;
        CoCreateInstance(ref clsid, null, 23 /*CLSCTX_ALL*/, ref iidEnum, out object obj);

        var en = (IMMDeviceEnumerator)obj;
        en.EnumAudioEndpoints((EDataFlow)df, 0x1 /*DEVICE_STATE_ACTIVE*/, out IMMDeviceCollection col);
        col.GetCount(out uint n);
        for (uint i = 0; i < n; i++) { col.Item(i, out IMMDevice dev); list.Add(new MMDevice(dev)); }
        Marshal.ReleaseComObject(col); Marshal.ReleaseComObject(en);
        return list;
    }

    // IMMDevice 是 internal，因此把以下 3 个方法也设为 internal，避免把 IMMDevice 暴露到公共 API
    internal static string GetDeviceFriendlyName(IMMDevice dev)
    {
        dev.OpenPropertyStore(0 /*STGM_READ*/, out IPropertyStore ps);
        var key = PKEY_Device_FriendlyName;
        ps.GetValue(ref key, out PROPVARIANT v);
        string name = v.GetString(); PropVariantClear(ref v); Marshal.ReleaseComObject(ps);
        return name;
    }

    internal static IAudioClient ActivateAudioClient(IMMDevice dev)
    {
        Guid iid = typeof(IAudioClient).GUID; dev.Activate(ref iid, 23 /*CLSCTX_ALL*/, IntPtr.Zero, out object o); return (IAudioClient)o;
    }

    internal static IKsControl GetKsControl(IMMDevice dev)
    {
        // Path 1) IAudioClient.GetService(IKsControl)
        try { var ac = ActivateAudioClient(dev); Guid iidKs = typeof(IKsControl).GUID; ac.GetService(ref iidKs, out object svc); if (svc is IKsControl k1) return k1; } catch { }

        // Path 2) DeviceTopology Part::Activate(IKsControl)
        Guid iidTopo = typeof(IDeviceTopology).GUID;
        dev.Activate(ref iidTopo, 23, IntPtr.Zero, out object topoObj);
        var topo = (IDeviceTopology)topoObj;
        topo.GetConnectorCount(out uint cc);
        if (cc > 0)
        {
            topo.GetConnector(0, out IConnector epConn);
            if (epConn.GetConnectedTo(out IConnector devConn) == 0)
            {
                var IID_IPart = typeof(IPart).GUID;
                if (devConn.QueryInterface(ref IID_IPart, out object partObj) == 0)
                {
                    var part = (IPart)partObj; Guid iidKs = typeof(IKsControl).GUID;
                    if (part.Activate(23, ref iidKs, out object ksObj) == 0 && ksObj is IKsControl k2)
                    {
                        Marshal.ReleaseComObject(devConn);
                        Marshal.ReleaseComObject(epConn);
                        Marshal.ReleaseComObject(topo);
                        return k2;
                    }
                }
                // Path 3) 兜底：在设备侧 Connector 上 QI(IKsControl)
                Guid iidKs2 = typeof(IKsControl).GUID;
                if (devConn.QueryInterface(ref iidKs2, out object ksObj2) == 0 && ksObj2 is IKsControl k3)
                {
                    Marshal.ReleaseComObject(devConn);
                    Marshal.ReleaseComObject(epConn);
                    Marshal.ReleaseComObject(topo);
                    return k3;
                }
                Marshal.ReleaseComObject(devConn);
            }
            Marshal.ReleaseComObject(epConn);
        }
        Marshal.ReleaseComObject(topo);
        throw new NotSupportedException("无法获取 IKsControl");
    }

    [DllImport("ole32.dll")] static extern int CoCreateInstance(ref Guid rclsid, [MarshalAs(UnmanagedType.IUnknown)] object? pUnkOuter, uint dwClsContext, ref Guid riid, out object ppv);
    [DllImport("ole32.dll")] internal static extern int PropVariantClear(ref PROPVARIANT pvar);

    // ---- COM & PROP 互操作定义 ----
    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow dataFlow, uint dwStateMask, out IMMDeviceCollection ppDevices);
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
        int RegisterEndpointNotificationCallback(IntPtr pClient);
        int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }
    internal enum EDataFlow { eRender = 0, eCapture = 1, eAll = 2 }
    internal enum ERole { eConsole = 0, eMultimedia = 1, eCommunications = 2 }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-C0F66554870A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection
    {
        int GetCount(out uint pcDevices);
        int Item(uint nDevice, out IMMDevice ppDevice);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        int Activate(ref Guid iid, uint dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore(uint stgmAccess, out IPropertyStore ppProperties);
        int GetId(out IntPtr ppstrId);
        int GetState(out uint pdwState);
    }

    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore
    {
        int GetCount(out uint cProps);
        int GetAt(uint iProp, out PROPERTYKEY pkey);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        int Commit();
    }

    [StructLayout(LayoutKind.Sequential)] internal struct PROPERTYKEY { public Guid fmtid; public uint pid; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct PROPVARIANT { public ushort vt; public ushort wReserved1, wReserved2, wReserved3; public IntPtr p; public int p2, p3; public string GetString() { const ushort VT_LPWSTR = 31; return (vt == VT_LPWSTR && p != IntPtr.Zero) ? Marshal.PtrToStringUni(p)! : ""; } }
    static readonly PROPERTYKEY PKEY_Device_FriendlyName = new PROPERTYKEY { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 14 };

    // public：供外部类型使用（AudioClient/KsControl 构造函数里需要）
    [ComImport, Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioClient
    {
        int Initialize(int shareMode, int streamFlags, long hnsBufferDuration, long hnsPeriodicity, IntPtr pFormat, IntPtr audioSessionGuid);
        int GetBufferSize(out uint pNumBufferFrames);
        int GetStreamLatency(out long phnsLatency);
        int GetCurrentPadding(out uint pNumPaddingFrames);
        int IsFormatSupported(int shareMode, IntPtr pFormat, IntPtr ppClosestMatch);
        int GetMixFormat(out IntPtr ppDeviceFormat);
        int GetDevicePeriod(out long phnsDefaultDevicePeriod, out long phnsMinimumDevicePeriod);
        int Start();
        int Stop();
        int Reset();
        int SetEventHandle(IntPtr eventHandle);
        int GetService(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct WAVEFORMATEX
    {
        public ushort wFormatTag, nChannels;
        public uint nSamplesPerSec, nAvgBytesPerSec;
        public ushort nBlockAlign, wBitsPerSample, cbSize;
    }
    internal static class WaveFormatEx
    {
        public static WAVEFORMATEX S16Stereo48k() => new WAVEFORMATEX
        {
            wFormatTag = 1, nChannels = 2, nSamplesPerSec = 48000,
            wBitsPerSample = 16, nBlockAlign = (ushort)(2 * 16 / 8),
            nAvgBytesPerSec = 48000 * (uint)(2 * 16 / 8), cbSize = 0
        };
    }

    [ComImport, Guid("2A07407E-6497-4A18-9787-32F79BD0D98F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDeviceTopology
    {
        int GetConnectorCount(out uint pCount);
        int GetConnector(uint Index, out IConnector ppConnector);
    }

    [ComImport, Guid("9c2c4058-23f5-41de-877a-df3af236a09e"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IConnector
    {
        int GetType(out int pType);
        int GetDataFlow(out int pFlow);
        int ConnectTo(IConnector pConnectTo);
        int Disconnect();
        int IsConnected(out int pbConnected);
        int GetConnectedTo(out IConnector ppConTo);
        int GetConnectorIdConnectedTo(out IntPtr ppwstrConnectorId);
        int GetDeviceIdConnectedTo(out IntPtr ppwstrDeviceId);
        int QueryInterface(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        // 释放统一用 Marshal.ReleaseComObject
    }

    [ComImport, Guid("AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPart
    {
        int GetName(out IntPtr ppwstrName);
        int GetLocalId(out uint pnId);
        int GetGlobalId(out IntPtr pguid);
        int GetPartType(out int pPartType);
        int GetSubType(out Guid pSubType);
        int GetControlInterfaceCount(out uint pCount);
        int GetControlInterface(uint nControl, out IntPtr ppInterfaceDesc);
        int EnumPartsIncoming(out IntPtr ppParts);
        int EnumPartsOutgoing(out IntPtr ppParts);
        int GetTopologyObject(out IDeviceTopology ppTopology);
        int Activate(uint dwClsContext, ref Guid refiid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);
        int RegisterControlChangeCallback(Guid riid, IntPtr pNotify);
        int UnregisterControlChangeCallback(IntPtr pNotify);
    }

    // public：供 KsControl 类使用
    [ComImport, Guid("28F54685-06FD-11d2-B27A-00A0C9223196"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IKsControl
    {
        int KsProperty(ref KSPROPERTY Property, uint PropertyLength, byte[] PropertyData, uint DataLength, ref uint BytesReturned);
        int KsMethod(IntPtr Method, uint MethodLength, IntPtr MethodData, uint DataLength, ref uint BytesReturned);
        int KsEvent(IntPtr Event, uint EventLength, IntPtr EventData, uint DataLength, ref uint BytesReturned);
    }
}

internal static class Entry
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
