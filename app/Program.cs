// Program.cs — MyCompany APO Control (.NET 9 + WinForms + IKsControl)
// 目的：枚举 Render 端点 ->（可选）强制拉流 -> 用 IKsControl 下发 12 段 EQ/混响/限幅参数（PID=10）
//
// 本版增强（UI 排列/布局修复版 + 稳定性修复）：
// 1) PropertyGrid 关闭 Help 面板与工具栏，避免占用大量垂直空间
// 2) 仅按“分类”显示；分类名加入序号 01/02/03，保证 Main→EQ→FX 的排序
// 3) 12 段 EQ 的 DisplayName 使用零填充（Eq 00…Eq 11），避免字母序导致 Eq10/11 排在 Eq1 后面
// 4) 顶部工具条 FlowLayoutPanel 自适应宽度、单行不换行；窗体启用 DPI 自适应与最小尺寸
// 5) “强制拉流”支持自定义毫秒数；两个“打开日志”按钮文案区分；“仅 GET 探针”带说明
// 6) 修复拓扑路径 RCW 误释放导致的 InvalidComObjectException；修正 NOPERSIST 常量；完善释放与错误提示
//
// 温馨提示：如果点击 Apply 后进程仍“瞬间退出无提示”，通常是 ksproxy/audiodg/驱动端崩了；
// 请配合 ProcDump/WER/WinDbg 指南抓取 dump。

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

// =============== 公共定义（GUID/PID、日志、KSPROPERTY 头） ===============

internal static class KsGuids
{
    public static readonly Guid MYCOMPANY_APO_PROPSETID = new Guid("d4d9a040-8b5f-4c0e-aad1-aabbccddeeff");
    public const uint PID_Gain = 1, PID_EQBand = 2, PID_Reverb = 3, PID_Limiter = 4, PID_ParamsBlob = 10;

    public const uint KSPROPERTY_TYPE_GET = 0x00000001;
    public const uint KSPROPERTY_TYPE_SET = 0x00000002;
}

[StructLayout(LayoutKind.Sequential)]
public struct KSPROPERTY { public Guid Set; public uint Id; public uint Flags; }

static class DebugLog
{
    // 日志优先保存在 exe 目录的 logs/ 下，失败再回退到 %TEMP%
    public static readonly string LogPath = InitLogPath();
    public static readonly string PayloadPath = Path.Combine(Path.GetDirectoryName(LogPath)!, "last-payload.bin");

    static string InitLogPath()
    {
        try
        {
            var baseDir = AppContext.BaseDirectory;
            var dir = Path.Combine(baseDir, "logs");
            Directory.CreateDirectory(dir);
            var p = Path.Combine(dir, "MyApoCtl.log");
            File.AppendAllText(p, $"{DateTime.Now:O}  <app start>\n", Encoding.UTF8);
            return p;
        }
        catch
        {
            var p = Path.Combine(Path.GetTempPath(), "MyApoCtl.log");
            try { File.AppendAllText(p, $"{DateTime.Now:O}  <app start>\n", Encoding.UTF8); } catch { }
            return p;
        }
    }

    [DllImport("kernel32.dll")] static extern bool AllocConsole();
    public static void AttachConsole()
    {
        try { AllocConsole(); } catch { }
        Info("=== Console attached ===");
        Info("Log file: " + LogPath);
    }

    public static void Info(string msg)
    {
        var line = $"{DateTime.Now:O}  {msg}";
        try { File.AppendAllText(LogPath, line + Environment.NewLine, Encoding.UTF8); } catch { }
        try { Console.WriteLine(line); } catch { }
        System.Diagnostics.Debug.WriteLine(line);
    }

    public static void HexPreview(string title, byte[] data, int max = 64)
    {
        int n = Math.Min(max, data?.Length ?? 0);
        var sb = new StringBuilder();
        for (int i = 0; i < n; i++) sb.Append(data[i].ToString("X2")).Append(i % 16 == 15 ? " " : " ");
        Info($"{title} (first {n} bytes): {sb}");
    }

    [HandleProcessCorruptedStateExceptions]
    public static void ShowAndLog(string title, Exception ex)
    {
        var hr = (ex as COMException)?.ErrorCode ?? 0;
        Info($"[{title}] {ex.GetType().Name}: 0x{(uint)hr:X8}\n{ex}");
        try { MessageBox.Show($"{title}\n{ex}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
    }
}

// ===================== 参数模型 & 打包（784B） =====================

public enum EqType : int { Peak = 0, LowShelf = 1, HighShelf = 2 }

[TypeConverter(typeof(ExpandableObjectConverter))]
public class EqBandModel
{
    public bool Enabled { get; set; } = false;
    public float Freq { get; set; } = 1000f;
    public float Q { get; set; } = 1.0f;
    public float GainDb { get; set; } = 0f;
    public EqType Type { get; set; } = EqType.Peak;
    public override string ToString() =>
        $"{(Enabled ? "On" : "Off")} {Type} {Freq}Hz Q={Q} {GainDb:+0.0;-0.0;0}dB";
}

[TypeConverter(typeof(ExpandableObjectConverter))]
public class ReverbModel
{
    public bool Enabled { get; set; } = false;
    public float Wet { get; set; } = 0.2f;
    public float Room { get; set; } = 0.7f;
    public float Damp { get; set; } = 0.3f;
    public float PreMs { get; set; } = 20f;
    public override string ToString() =>
        $"{(Enabled ? "On" : "Off")} wet={Wet:0.00} room={Room:0.00} damp={Damp:0.00} pre={PreMs}ms";
}

[TypeConverter(typeof(ExpandableObjectConverter))]
public class DspParamsModel
{
    [Category("01 Main")]
    public float Gain { get; set; } = 1.0f;

    [Category("02 EQ (12 bands)"), DisplayName("Eq 00")]
    public EqBandModel Eq0 { get; set; } = new EqBandModel { Type = EqType.LowShelf, Freq = 100, Q = 0.707f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 01")]
    public EqBandModel Eq1 { get; set; } = new EqBandModel { Type = EqType.Peak, Freq = 250, Q = 1.0f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 02")]
    public EqBandModel Eq2 { get; set; } = new EqBandModel { Type = EqType.Peak, Freq = 500, Q = 1.0f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 03")]
    public EqBandModel Eq3 { get; set; } = new EqBandModel { Type = EqType.Peak, Freq = 1000, Q = 1.0f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 04")]
    public EqBandModel Eq4 { get; set; } = new EqBandModel { Type = EqType.Peak, Freq = 2000, Q = 1.0f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 05")]
    public EqBandModel Eq5 { get; set; } = new EqBandModel { Type = EqType.Peak, Freq = 4000, Q = 1.0f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 06")]
    public EqBandModel Eq6 { get; set; } = new EqBandModel { Type = EqType.Peak, Freq = 8000, Q = 1.0f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 07")]
    public EqBandModel Eq7 { get; set; } = new EqBandModel { Type = EqType.HighShelf, Freq = 12000, Q = 0.707f };
    [Category("02 EQ (12 bands)"), DisplayName("Eq 08")]
    public EqBandModel Eq8 { get; set; } = new EqBandModel();
    [Category("02 EQ (12 bands)"), DisplayName("Eq 09")]
    public EqBandModel Eq9 { get; set; } = new EqBandModel();
    [Category("02 EQ (12 bands)"), DisplayName("Eq 10")]
    public EqBandModel Eq10 { get; set; } = new EqBandModel();
    [Category("02 EQ (12 bands)"), DisplayName("Eq 11")]
    public EqBandModel Eq11 { get; set; } = new EqBandModel();

    [Category("03 FX")]
    public ReverbModel Reverb { get; set; } = new ReverbModel();

    [Category("03 FX")]
    public bool LimiterEnabled { get; set; } = true;

    [Browsable(false)] public byte[] Opcode { get; set; } = Array.Empty<byte>();

    public IEnumerable<EqBandModel> Bands()
    {
        yield return Eq0; yield return Eq1; yield return Eq2; yield return Eq3; yield return Eq4; yield return Eq5;
        yield return Eq6; yield return Eq7; yield return Eq8; yield return Eq9; yield return Eq10; yield return Eq11;
    }
}

internal static class ParamsPacker
{
    public static byte[] Pack(DspParamsModel m)
    {
        byte[] buf = new byte[784];
        int off = 0;
        void W32(int v) { BitConverter.GetBytes(v).CopyTo(buf, off); off += 4; }
        void WU32(uint v) { BitConverter.GetBytes(v).CopyTo(buf, off); off += 4; }
        void WF(float f) { BitConverter.GetBytes(f).CopyTo(buf, off); off += 4; }

        WF(m.Gain);
        foreach (var b in m.Bands())
        { W32(b.Enabled ? 1 : 0); WF(b.Freq); WF(b.Q); WF(b.GainDb); W32((int)b.Type); }
        W32(m.Reverb.Enabled ? 1 : 0); WF(m.Reverb.Wet); WF(m.Reverb.Room); WF(m.Reverb.Damp); WF(m.Reverb.PreMs);
        W32(m.LimiterEnabled ? 1 : 0);

        var op = (m.Opcode ?? Array.Empty<byte>()).Take(512).ToArray();
        WU32((uint)op.Length);
        Array.Copy(op, 0, buf, off, op.Length);

        try { File.WriteAllBytes(DebugLog.PayloadPath, buf); } catch { }
        return buf;
    }
}

// ===================== UI =====================

public enum KsPath { Auto, GetService, Topology }

public class MainForm : Form
{
    ComboBox cboDev = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 420 };
    ComboBox cboPath = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
    CheckBox chkProbe = new CheckBox { Text = "仅 GET 探针", AutoSize = true };
    Button btnRefresh = new Button { Text = "刷新设备" };
    Button btnForce = new Button { Text = "强制拉流" };
    NumericUpDown nudMs = new NumericUpDown { Minimum = 10, Maximum = 2000, Increment = 10, Value = 150, Width = 80 };
    Button btnOpenLog = new Button { Text = "打开日志" };
    Button btnOpenDir = new Button { Text = "打开目录" };
    PropertyGrid grid = new PropertyGrid { Dock = DockStyle.Fill };
    Button btnApply = new Button { Text = "Apply (KsProperty)", Dock = DockStyle.Bottom, Height = 40 };

    List<MMDevice> devices = new();
    DspParamsModel model = new();

    public MainForm()
    {
        Text = $"MyCompany APO Control (IKsControl + 12-band EQ) | log: {DebugLog.LogPath}";
        Width = 980; Height = 720;

        this.AutoScaleMode = AutoScaleMode.Dpi;
        this.MinimumSize = new Size(960, 640);

        var top = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Padding = new Padding(0, 6, 0, 6)
        };
        top.Controls.Add(new Label { Text = "Render 设备：", AutoSize = true, Padding = new Padding(0, 10, 0, 0) });
        top.Controls.Add(cboDev);

        top.Controls.Add(new Label { Text = "  IKs 路径：", AutoSize = true, Padding = new Padding(10, 10, 0, 0) });
        cboPath.Items.AddRange(new object[] { "Auto", "GetService", "Topology" });
        cboPath.SelectedIndex = 0;
        top.Controls.Add(cboPath);

        var tips = new ToolTip();
        tips.SetToolTip(chkProbe,
            "只做 KsProperty(GET) 探测：\n" +
            "• 成功：返回需要的缓冲大小（常见 STATUS_BUFFER_OVERFLOW）\n" +
            "• 不支持：返回 E_PROP_ID_UNSUPPORTED/STATUS_NOT_FOUND\n" +
            "取消勾选再点 Apply 会执行 SET（下发 PID=10 参数包）。");

        top.Controls.Add(chkProbe);
        top.Controls.Add(btnRefresh);

        top.Controls.Add(new Label { Text = "  拉流(ms)：", AutoSize = true, Padding = new Padding(10, 10, 0, 0) });
        top.Controls.Add(nudMs);
        top.Controls.Add(btnForce);

        btnOpenLog.MinimumSize = new Size(88, 0);
        btnOpenDir.MinimumSize = new Size(88, 0);
        top.Controls.Add(btnOpenLog);
        top.Controls.Add(btnOpenDir);
        Controls.Add(top);

        grid.PropertySort = PropertySort.Categorized;
        grid.HelpVisible = false;
        grid.ToolbarVisible = false;
        grid.SelectedObject = model;
        Controls.Add(grid);

        Controls.Add(btnApply);

        btnRefresh.Click += (_, __) => RefreshDevices();
        btnForce.Click += (_, __) =>
        {
            var d = CurrentDevice(); if (d == null) return;
            using var ac = d.ActivateAudioClient();
            ac.InitializeSharedNoPersist(48000, 2, (int)nudMs.Value);
            MessageBox.Show($"已拉流 {(int)nudMs.Value} ms（共享模式，NOPERSIST）");
        };
        btnOpenLog.Click += (_, __) => OpenPath(DebugLog.LogPath, select: true);
        btnOpenDir.Click += (_, __) => OpenPath(Path.GetDirectoryName(DebugLog.LogPath)!, select: false);

        btnApply.Click += (_, __) =>
        {
            try
            {
                var dev = CurrentDevice();
                if (dev == null) { MessageBox.Show("请选择设备"); return; }

                byte[] blob = ParamsPacker.Pack(model);
                DebugLog.Info($"Apply clicked. device='{dev.FriendlyName}', blobLen={blob.Length}");
                DebugLog.HexPreview("payload", blob);

                using var ks = dev.OpenKsControl(GetSelectedPath(), forceStream: true, forceMs: (int)nudMs.Value);

                // 先 GET 探针
                int hrProbe = ks.ProbeBlobSupport(out uint need);
                DebugLog.Info($"Probe(GET) => hr=0x{(uint)hrProbe:X8}, bytesNeeded={need}");

                if (!chkProbe.Checked)
                {
                    int hrSet = ks.TrySetBlob(blob, out uint ret);
                    DebugLog.Info($"SET => hr=0x{(uint)hrSet:X8}, bytesReturned={ret}");
                    if (hrSet < 0) throw new COMException($"KsProperty(SET) 失败 hr=0x{hrSet:X8}", hrSet);
                    MessageBox.Show($"参数已下发：{blob.Length} bytes (PID=10)");
                }
                else
                {
                    MessageBox.Show("已完成 GET 探针（未下发 SET）。");
                }
            }
            catch (COMException ex) { DebugLog.ShowAndLog("COM 调用失败", ex); }
            catch (SEHException ex) { DebugLog.ShowAndLog("Native/SEH 异常", ex); }
            catch (Exception ex) { DebugLog.ShowAndLog("未处理异常", ex); }
        };

        Load += (_, __) => RefreshDevices();
    }

    KsPath GetSelectedPath() => (KsPath)cboPath.SelectedIndex;

    void OpenPath(string path, bool select)
    {
        try
        {
            if (select)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"/select,\"{path}\"") { UseShellExecute = true });
            else
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", $"\"{path}\"") { UseShellExecute = true });
        }
        catch { MessageBox.Show(path); }
    }

    void RefreshDevices()
    {
        devices = CoreAudio.ListActiveEndpoints(DataFlow.Render);
        cboDev.Items.Clear();
        foreach (var d in devices) cboDev.Items.Add(d.FriendlyName);
        if (cboDev.Items.Count > 0) cboDev.SelectedIndex = 0;
    }

    MMDevice? CurrentDevice()
        => (cboDev.SelectedIndex >= 0 && cboDev.SelectedIndex < devices.Count) ? devices[cboDev.SelectedIndex] : null;
}

// ===================== CoreAudio + IKsControl 封装 =====================

public enum DataFlow { Render = 0, Capture = 1 }

public sealed class MMDevice : IDisposable
{
    internal CoreAudio.IMMDevice dev;
    public string FriendlyName { get; }
    internal MMDevice(CoreAudio.IMMDevice d) { dev = d; FriendlyName = CoreAudio.GetDeviceFriendlyName(d); }
    public void Dispose() { if (dev != null) Marshal.ReleaseComObject(dev); }

    public AudioClient ActivateAudioClient() => new AudioClient(CoreAudio.ActivateAudioClient(dev));

    public KsControl OpenKsControl(KsPath path = KsPath.Auto, bool forceStream = false, int forceMs = 150)
    {
        if (forceStream) { using var ac = ActivateAudioClient(); ac.InitializeSharedNoPersist(48000, 2, forceMs); }
        var ks = CoreAudio.GetKsControl(dev, path);
        return new KsControl(ks);
    }
}

public sealed class AudioClient : IDisposable
{
    internal CoreAudio.IAudioClient ac;
    public AudioClient(CoreAudio.IAudioClient a) { ac = a; }

    public void InitializeSharedNoPersist(int sampleRate, int channels, int runMs)
    {
        const int AUDCLNT_STREAMFLAGS_NOPERSIST = 0x00008000; // 修正常量
        int hr = ac.GetMixFormat(out IntPtr pwfx);
        if (hr >= 0 && pwfx != IntPtr.Zero)
        {
            hr = ac.Initialize(0 /*shared*/, AUDCLNT_STREAMFLAGS_NOPERSIST, 10_000_000, 0, pwfx, IntPtr.Zero);
            Marshal.FreeCoTaskMem(pwfx);
        }
        else
        {
            var fmt = CoreAudio.WaveFormatEx.S16Stereo48k();
            IntPtr pFmt = Marshal.AllocHGlobal(Marshal.SizeOf<CoreAudio.WAVEFORMATEX>());
            try { Marshal.StructureToPtr(fmt, pFmt, false); hr = ac.Initialize(0, AUDCLNT_STREAMFLAGS_NOPERSIST, 10_000_000, 0, pFmt, IntPtr.Zero); }
            finally { Marshal.FreeHGlobal(pFmt); }
        }
        if (hr >= 0) { ac.Start(); System.Threading.Thread.Sleep(runMs); ac.Stop(); }
    }

    public void Dispose() { if (ac != null) Marshal.ReleaseComObject(ac); }
}

// IKsControl —— 全部走 IntPtr，规避托管封送；每步都打日志
public sealed class KsControl : IDisposable
{
    internal CoreAudio.IKsControl ks;
    public KsControl(CoreAudio.IKsControl k) { ks = k; }
    public void Dispose() { if (ks != null) Marshal.ReleaseComObject(ks); }

    public int ProbeBlobSupport(out uint bytesNeeded)
    {
        var prop = new KSPROPERTY { Set = KsGuids.MYCOMPANY_APO_PROPSETID, Id = KsGuids.PID_ParamsBlob, Flags = KsGuids.KSPROPERTY_TYPE_GET };
        DebugLog.Info("Probe(GET) -> calling KsProperty(GET)...");
        var hr = ks.KsProperty(ref prop, (uint)Marshal.SizeOf<KSPROPERTY>(), IntPtr.Zero, 0, out bytesNeeded);
        DebugLog.Info($"Probe(GET) <- hr=0x{(uint)hr:X8}, need={bytesNeeded}");
        return hr;
    }

    public int TrySetBlob(byte[] bytes, out uint bytesReturned)
    {
        var prop = new KSPROPERTY { Set = KsGuids.MYCOMPANY_APO_PROPSETID, Id = KsGuids.PID_ParamsBlob, Flags = KsGuids.KSPROPERTY_TYPE_SET };
        IntPtr p = IntPtr.Zero;
        try
        {
            p = Marshal.AllocHGlobal(bytes.Length);
            Marshal.Copy(bytes, 0, p, bytes.Length);
            DebugLog.Info($"SET -> calling KsProperty(SET) len={bytes.Length}...");
            var hr = ks.KsProperty(ref prop, (uint)Marshal.SizeOf<KSPROPERTY>(), p, (uint)bytes.Length, out bytesReturned);
            DebugLog.Info($"SET <- hr=0x{(uint)hr:X8}, ret={bytesReturned}");
            return hr;
        }
        finally { if (p != IntPtr.Zero) Marshal.FreeHGlobal(p); }
    }
}

public static class CoreAudio
{
    const uint DEVICE_STATE_ACTIVE = 0x00000001;
    const uint CLSCTX_ALL = 23;

    public static List<MMDevice> ListActiveEndpoints(DataFlow df)
    {
        var list = new List<MMDevice>();
        var en = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
        try
        {
            if (en.EnumAudioEndpoints((EDataFlow)df, DEVICE_STATE_ACTIVE, out IMMDeviceCollection col) >= 0 && col != null)
            {
                try
                {
                    if (col.GetCount(out uint n) >= 0)
                    {
                        for (uint i = 0; i < n; i++)
                            if (col.Item(i, out IMMDevice dev) >= 0) list.Add(new MMDevice(dev));
                        if (list.Count > 0) return list;
                    }
                }
                finally { SafeRelease(col); }
            }
        }
        catch { }
        finally { SafeRelease(en); }
        return GetDefaultRenderEndpointsSafe();
    }

    private static List<MMDevice> GetDefaultRenderEndpointsSafe()
    {
        var ret = new List<MMDevice>();
        var en2 = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
        try
        {
            void Add(ERole role)
            {
                if (en2.GetDefaultAudioEndpoint(EDataFlow.eRender, role, out IMMDevice dev) >= 0 && dev != null)
                {
                    string id = GetDeviceId(dev);
                    if (!ret.Any(d => GetDeviceId(d.dev) == id)) ret.Add(new MMDevice(dev));
                    else SafeRelease(dev);
                }
            }
            Add(ERole.eConsole); Add(ERole.eMultimedia); Add(ERole.eCommunications);
        }
        catch { }
        finally { SafeRelease(en2); }
        return ret;
    }

    internal static string GetDeviceId(IMMDevice dev)
    {
        try
        {
            if (dev.GetId(out IntPtr pwsz) >= 0 && pwsz != IntPtr.Zero)
            { string id = Marshal.PtrToStringUni(pwsz) ?? "<device>"; Marshal.FreeCoTaskMem(pwsz); return id; }
        }
        catch { }
        return "<unknown>";
    }

    internal static string GetDeviceFriendlyName(IMMDevice dev)
    {
        try
        {
            if (dev.OpenPropertyStore(0, out IPropertyStore ps) >= 0)
            {
                try
                {
                    var key = PKEY_Device_FriendlyName; PROPVARIANT v = default;
                    if (ps.GetValue(ref key, out v) >= 0)
                    {
                        string name = PropVariantToString(ref v);
                        PropVariantClear(ref v);
                        if (!string.IsNullOrWhiteSpace(name)) return name;
                    }
                }
                finally { SafeRelease(ps); }
            }
        }
        catch { }
        try
        {
            if (dev.GetId(out IntPtr pwsz) >= 0 && pwsz != IntPtr.Zero)
            { string id = Marshal.PtrToStringUni(pwsz) ?? "<device>"; Marshal.FreeCoTaskMem(pwsz); return id; }
        }
        catch { }
        return "<unknown>";
    }

    private static string PropVariantToString(ref PROPVARIANT v)
    {
        const ushort VT_EMPTY = 0, VT_BSTR = 8, VT_LPWSTR = 31;
        try
        {
            if (v.vt == VT_LPWSTR && v.p != IntPtr.Zero) return Marshal.PtrToStringUni(v.p) ?? "";
            if (v.vt == VT_BSTR && v.p != IntPtr.Zero) return Marshal.PtrToStringBSTR(v.p) ?? "";
            if (v.vt == VT_EMPTY || v.p == IntPtr.Zero) return "";
        }
        catch { }
        return "";
    }

    internal static IAudioClient ActivateAudioClient(IMMDevice dev)
    {
        Guid iid = typeof(IAudioClient).GUID;
        int hr = dev.Activate(ref iid, CLSCTX_ALL, IntPtr.Zero, out object o);
        if (hr < 0) throw new COMException("IMMDevice.Activate(IAudioClient) failed", hr);
        return (IAudioClient)o;
    }

    internal static IKsControl GetKsControl(IMMDevice dev, KsPath path)
    {
        Guid iidKs = typeof(IKsControl).GUID;

        // 0) 直连：IMMDevice.Activate(IKsControl)
        if (path == KsPath.Auto)
        {
            try
            {
                int hr0 = dev.Activate(ref iidKs, CLSCTX_ALL, IntPtr.Zero, out object ks0);
                DebugLog.Info($"IMMDevice.Activate(IKsControl) hr=0x{(uint)hr0:X8}");
                if (hr0 >= 0 && ks0 is IKsControl k0) { DebugLog.Info("IKsControl: via IMMDevice.Activate"); return k0; }
            }
            catch (Exception ex) { DebugLog.Info("IMMDevice.Activate(IKsControl) ex: " + ex.Message); }
        }

        // 1) IAudioClient.GetService(IKsControl)
        if (path == KsPath.Auto || path == KsPath.GetService)
        {
            if (TryGetKsViaAudioClient(dev, out IKsControl k1)) return k1;
            if (path == KsPath.GetService)
                throw new COMException("AudioClient.GetService 未提供 IKsControl（此路径并非所有设备都支持）", unchecked((int)0x80004002)); // E_NOINTERFACE
        }

        // 2) DeviceTopology / IPart.Activate(IKsControl)
        if (TryGetKsViaTopology(dev, out IKsControl k2)) return k2;

        throw new NotSupportedException("无法获取 IKsControl（IMMDevice.Activate / GetService / Topology 均失败）");
    }

    private static bool TryGetKsViaAudioClient(IMMDevice dev, out IKsControl ks)
    {
        ks = null;
        IAudioClient ac = null;
        try
        {
            ac = ActivateAudioClient(dev);
            Guid iidKs = typeof(IKsControl).GUID;
            int hr = ac.GetService(ref iidKs, out object svc);
            DebugLog.Info($"IAudioClient.GetService(IKsControl) hr=0x{(uint)hr:X8}");
            if (hr >= 0 && svc is IKsControl k) { ks = k; DebugLog.Info("IKsControl: via IAudioClient.GetService"); return true; }
        }
        catch (Exception ex) { DebugLog.Info("GetService path ex: " + ex.Message); }
        finally { SafeRelease(ac); }
        return false;
    }

    private static bool TryGetKsViaTopology(IMMDevice dev, out IKsControl ks)
    {
        ks = null;
        Guid iidTopo = typeof(IDeviceTopology).GUID;
        int hrTopo = dev.Activate(ref iidTopo, CLSCTX_ALL, IntPtr.Zero, out object topoObj);
        DebugLog.Info($"IMMDevice.Activate(IDeviceTopology) hr=0x{(uint)hrTopo:X8}");
        if (hrTopo < 0) return false;

        var topo = (IDeviceTopology)topoObj;
        try
        {
            topo.GetConnectorCount(out uint cc);
            DebugLog.Info($"Topology: connectorCount={cc}");
            for (uint i = 0; i < cc; i++)
            {
                topo.GetConnector(i, out IConnector epConn);
                try
                {
                    // (A) 先在 endpoint 侧 connector 上尝试 IPart.Activate
                    if (AsPartActivateKs(epConn, out ks))
                    { DebugLog.Info($"IKsControl: via endpoint connector[{i}]"); return true; }

                    // (B) 再到 device 侧
                    if (epConn.GetConnectedTo(out IConnector devConn) == 0 && devConn != null)
                    {
                        try
                        {
                            if (AsPartActivateKs(devConn, out ks))
                            { DebugLog.Info($"IKsControl: via device connector[{i}]"); return true; }

                            // (C) 兜底：用 deviceId 找到真正设备，再试 Activate / GetService
                            if (devConn.GetDeviceIdConnectedTo(out IntPtr pDevId) == 0 && pDevId != IntPtr.Zero)
                            {
                                try
                                {
                                    string devId = Marshal.PtrToStringUni(pDevId) ?? "";
                                    DebugLog.Info($"Connected deviceId: {devId}");
                                    var en = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
                                    try
                                    {
                                        if (en.GetDevice(devId, out IMMDevice realDev) == 0 && realDev != null)
                                        {
                                            try
                                            {
                                                Guid iidKs = typeof(IKsControl).GUID;
                                                int hrA = realDev.Activate(ref iidKs, CLSCTX_ALL, IntPtr.Zero, out object ksObj);
                                                DebugLog.Info($"Connected IMMDevice.Activate(IKsControl) hr=0x{(uint)hrA:X8}");
                                                if (hrA >= 0 && ksObj is IKsControl kA)
                                                { ks = kA; DebugLog.Info("IKsControl: via connected IMMDevice.Activate"); return true; }

                                                if (TryGetKsViaAudioClient(realDev, out ks))
                                                { DebugLog.Info("IKsControl: via connected IAudioClient.GetService"); return true; }
                                            }
                                            finally { SafeRelease(realDev); }
                                        }
                                    }
                                    finally { SafeRelease(en); }
                                }
                                finally { FreeCoTaskMemSafe(pDevId); }
                            }
                        }
                        finally { SafeRelease(devConn); }
                    }
                }
                finally { SafeRelease(epConn); }
            }
        }
        finally { SafeRelease(topo); }

        return false;
    }

    // ---- 内部工具 ----
    private static void SafeRelease(object com)
    {
        try
        {
            if (com != null && Marshal.IsComObject(com))
                Marshal.ReleaseComObject(com);
        }
        catch { }
    }
    private static void FreeCoTaskMemSafe(IntPtr p)
    {
        if (p != IntPtr.Zero)
        {
            try { Marshal.FreeCoTaskMem(p); } catch { }
        }
    }

    // 关键：用 QI 指针版，避免误释放外层 RCW 导致 InvalidComObjectException
    private static bool AsPartActivateKs(object connObj, out IKsControl ks)
    {
        ks = null;
        IntPtr unk = IntPtr.Zero;
        IntPtr pPart = IntPtr.Zero;
        try
        {
            // 从任意 RCW 取 IUnknown*，手动 QI 到 IPart
            unk = Marshal.GetIUnknownForObject(connObj);
            Guid iidPart = typeof(IPart).GUID;
            int hrQI = Marshal.QueryInterface(unk, ref iidPart, out pPart);
            DebugLog.Info($"QI(IID_IPart) hr=0x{(uint)hrQI:X8}");
            if (hrQI < 0 || pPart == IntPtr.Zero) return false;

            // 用临时 RCW 包一层，仅用于 Activate 调用
            var part = (IPart)Marshal.GetObjectForIUnknown(pPart);

            Guid iidKs = typeof(IKsControl).GUID;
            int hr = part.Activate(CLSCTX_ALL, ref iidKs, out object ksObj);
            DebugLog.Info($"IPart.Activate(IKsControl) hr=0x{(uint)hr:X8}");
            if (hr >= 0 && ksObj is IKsControl k) { ks = k; return true; }
        }
        catch (Exception ex) { DebugLog.Info("AsPartActivateKs(QI) ex: " + ex.Message); }
        finally
        {
            if (pPart != IntPtr.Zero) { try { Marshal.Release(pPart); } catch { } }
            if (unk   != IntPtr.Zero) { try { Marshal.Release(unk);   } catch { } }
            // 不对 connObj / part 做 SafeRelease，避免影响外层 RCW
        }
        return false;
    }

    // ===== COM & PROP 互操作定义 =====

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumeratorComObject { } // coclass

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        [PreserveSig] int EnumAudioEndpoints(EDataFlow dataFlow, uint dwStateMask, out IMMDeviceCollection ppDevices);
        [PreserveSig] int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
        [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
        [PreserveSig] int RegisterEndpointNotificationCallback(IntPtr pClient);
        [PreserveSig] int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }
    public enum EDataFlow { eRender = 0, eCapture = 1, eAll = 2 }
    public enum ERole { eConsole = 0, eMultimedia = 1, eCommunications = 2 }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-C0F66554870A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceCollection
    {
        [PreserveSig] int GetCount(out uint pcDevices);
        [PreserveSig] int Item(uint nDevice, out IMMDevice ppDevice);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        [PreserveSig] int Activate(ref Guid iid, uint dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        [PreserveSig] int OpenPropertyStore(uint stgmAccess, out IPropertyStore ppProperties);
        [PreserveSig] int GetId(out IntPtr ppstrId);
        [PreserveSig] int GetState(out uint pdwState);
    }

    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore
    {
        [PreserveSig] int GetCount(out uint cProps);
        [PreserveSig] int GetAt(uint iProp, out PROPERTYKEY pkey);
        [PreserveSig] int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        [PreserveSig] int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        [PreserveSig] int Commit();
    }

    [StructLayout(LayoutKind.Sequential)] public struct PROPERTYKEY { public Guid fmtid; public uint pid; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PROPVARIANT
    {
        public ushort vt;
        public ushort wReserved1, wReserved2, wReserved3;
        public IntPtr p;
        public IntPtr p2;
        public IntPtr p3;
    }

    static readonly PROPERTYKEY PKEY_Device_FriendlyName = new PROPERTYKEY { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 14 };
    [DllImport("ole32.dll")] internal static extern int PropVariantClear(ref PROPVARIANT pvar);

    [ComImport, Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioClient
    {
        [PreserveSig] int Initialize(int shareMode, int streamFlags, long hnsBufferDuration, long hnsPeriodicity, IntPtr pFormat, IntPtr audioSessionGuid);
        [PreserveSig] int GetBufferSize(out uint pNumBufferFrames);
        [PreserveSig] int GetStreamLatency(out long phnsLatency);
        [PreserveSig] int GetCurrentPadding(out uint pNumPaddingFrames);
        [PreserveSig] int IsFormatSupported(int shareMode, IntPtr pFormat, IntPtr ppClosestMatch);
        [PreserveSig] int GetMixFormat(out IntPtr ppDeviceFormat);
        [PreserveSig] int GetDevicePeriod(out long phnsDefaultDevicePeriod, out long phnsMinimumDevicePeriod);
        [PreserveSig] int Start();
        [PreserveSig] int Stop();
        [PreserveSig] int Reset();
        [PreserveSig] int SetEventHandle(IntPtr eventHandle);
        [PreserveSig] int GetService(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WAVEFORMATEX
    {
        public ushort wFormatTag, nChannels;
        public uint nSamplesPerSec, nAvgBytesPerSec;
        public ushort nBlockAlign, wBitsPerSample, cbSize;
    }
    public static class WaveFormatEx
    {
        public static WAVEFORMATEX S16Stereo48k() => new WAVEFORMATEX
        {
            wFormatTag = 1,
            nChannels = 2,
            nSamplesPerSec = 48000,
            wBitsPerSample = 16,
            nBlockAlign = (ushort)(2 * 16 / 8),
            nAvgBytesPerSec = 48000 * (uint)(2 * 16 / 8),
            cbSize = 0
        };
    }

    [ComImport, Guid("2A07407E-6497-4A18-9787-32F79BD0D98F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDeviceTopology
    {
        [PreserveSig] int GetConnectorCount(out uint pCount);
        [PreserveSig] int GetConnector(uint Index, out IConnector ppConnector);
    }

    [ComImport, Guid("9c2c4058-23f5-41de-877a-df3af236a09e"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IConnector
    {
        [PreserveSig] int GetType(out int pType);
        [PreserveSig] int GetDataFlow(out int pFlow);
        [PreserveSig] int ConnectTo(IConnector pConnectTo);
        [PreserveSig] int Disconnect();
        [PreserveSig] int IsConnected(out int pbConnected);
        [PreserveSig] int GetConnectedTo(out IConnector ppConTo);
        [PreserveSig] int GetConnectorIdConnectedTo(out IntPtr ppwstrConnectorId);
        [PreserveSig] int GetDeviceIdConnectedTo(out IntPtr ppwstrDeviceId);
    }

    [ComImport, Guid("AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPart
    {
        [PreserveSig] int GetName(out IntPtr ppwstrName);
        [PreserveSig] int GetLocalId(out uint pnId);
        [PreserveSig] int GetGlobalId(out IntPtr pguid);
        [PreserveSig] int GetPartType(out int pPartType);
        [PreserveSig] int GetSubType(out Guid pSubType);
        [PreserveSig] int GetControlInterfaceCount(out uint pCount);
        [PreserveSig] int GetControlInterface(uint nControl, out IntPtr ppInterfaceDesc);
        [PreserveSig] int EnumPartsIncoming(out IntPtr ppParts);
        [PreserveSig] int EnumPartsOutgoing(out IntPtr ppParts);
        [PreserveSig] int GetTopologyObject(out IDeviceTopology ppTopology);
        [PreserveSig] int Activate(uint dwClsContext, ref Guid refiid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);
        [PreserveSig] int RegisterControlChangeCallback(Guid riid, IntPtr pNotify);
        [PreserveSig] int UnregisterControlChangeCallback(IntPtr pNotify);
    }

    [ComImport, Guid("28F54685-06FD-11d2-B27A-00A0C9223196"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IKsControl
    {
        [PreserveSig]
        int KsProperty(ref KSPROPERTY Property, uint PropertyLength,
                       IntPtr PropertyData, uint DataLength,
                       out uint BytesReturned);

        [PreserveSig]
        int KsMethod(IntPtr Method, uint MethodLength,
                     IntPtr MethodData, uint DataLength,
                     out uint BytesReturned);

        [PreserveSig]
        int KsEvent(IntPtr Event, uint EventLength,
                    IntPtr EventData, uint DataLength,
                    out uint BytesReturned);
    }
}

// ===================== 入口 =====================

internal static class Entry
{
    [STAThread]
    static void Main()
    {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (s, e) => DebugLog.ShowAndLog("UI ThreadException", e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            DebugLog.ShowAndLog("UnhandledException", e.ExceptionObject as Exception ?? new Exception("unknown"));

        if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift) DebugLog.AttachConsole();

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}
