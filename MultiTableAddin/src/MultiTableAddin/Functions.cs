using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace MultiTableAddin;

public static class Functions
{
    [ExcelFunction(
        Name = "MULTITABLEADDIN_HELLO",
        Category = "MultiTableAddin",
        Description = "返回问候语，用于确认 Excel-DNA 插件已成功加载，并演示 IntelliSense 参数说明。")]
    public static string Hello(
        [ExcelArgument(Name = "name", Description = "任意名称，留空时默认返回“世界”。")] string name)
    {
        string safeName = string.IsNullOrWhiteSpace(name) ? "世界" : name.Trim();
        string host = HostEnvironment.GetHostDisplayName();
        return string.Format("你好，{0}。当前宿主：{1}", safeName, host);
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_OPTIONAL",
        Category = "MultiTableAddin",
        Description = "演示 Excel-DNA 1.9 对可选参数与默认值的直接支持。")]
    public static string OptionalMessage(
        [ExcelArgument(Name = "name", Description = "名称，省略时默认使用 ExcelDNA。")] string name = "ExcelDNA",
        [ExcelArgument(Name = "createdAt", Description = "时间戳，可省略。")] DateTime? createdAt = null)
    {
        string timeText = createdAt.HasValue
            ? createdAt.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
            : "未提供";

        return string.Format("name={0}; createdAt={1}", name, timeText);
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_PARAMS_SUM",
        Category = "MultiTableAddin",
        Description = "演示 Excel-DNA 1.9 对 params 可变参数数组的支持。")]
    public static double SumMany(
        [ExcelArgument(Name = "values", Description = "可变数量的数字参数。")] params double[] values)
    {
        double total = 0d;
        foreach (double value in values)
        {
            total += value;
        }

        return total;
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_TABLE",
        Category = "MultiTableAddin",
        Description = "演示直接返回二维数组，适合动态数组版本的 Excel / WPS。")]
    public static object[,] MakeTable(
        [ExcelArgument(Name = "rows", Description = "行数。")] int rows,
        [ExcelArgument(Name = "cols", Description = "列数。")] int cols)
    {
        return CreateTable(Math.Max(1, Math.Min(rows, 20)), Math.Max(1, Math.Min(cols, 10)), "同步");
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_LIVE_CLOCK",
        Category = "MultiTableAddin",
        Description = "演示 RTD 包装函数，持续推送当前时间。")]
    public static object LiveClock()
    {
        return XlCall.RTD(DemoRtdServer.ServerProgId, null, "CLOCK");
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_LIVE_WAVE",
        Category = "MultiTableAddin",
        Description = "演示 RTD 连续推送数值流，适合状态、行情、进度等场景。")]
    public static object LiveWave(
        [ExcelArgument(Name = "amplitude", Description = "振幅。")] double amplitude = 1d,
        [ExcelArgument(Name = "speed", Description = "变化速度。")] double speed = 0.4d)
    {
        return XlCall.RTD(
            DemoRtdServer.ServerProgId,
            null,
            "WAVE",
            amplitude.ToString(CultureInfo.InvariantCulture),
            speed.ToString(CultureInfo.InvariantCulture));
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_HANDLE_CREATE",
        Category = "MultiTableAddin",
        Description = "演示 Excel-DNA 1.9 内置对象句柄支持，创建一个会话对象并返回句柄。")]
    [return: ExcelHandle]
    public static DemoHandleSession CreateHandle(
        [ExcelArgument(Name = "owner", Description = "会话名称。")] string owner,
        [ExcelArgument(Name = "seedValue", Description = "初始值。")] double seedValue)
    {
        string safeOwner = string.IsNullOrWhiteSpace(owner) ? "匿名用户" : owner.Trim();
        return new DemoHandleSession(safeOwner, seedValue);
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_HANDLE_READ",
        Category = "MultiTableAddin",
        Description = "读取对象句柄中的当前内容。")]
    public static string ReadHandle(
        [ExcelHandle][ExcelArgument(Name = "handle", Description = "由 HANDLE_CREATE 返回的对象句柄。")] DemoHandleSession handle)
    {
        return handle == null
            ? "句柄无效"
            : string.Format("Owner={0}; Value={1:F2}; CreatedAt={2:yyyy-MM-dd HH:mm:ss}", handle.Owner, handle.Value, handle.CreatedAt);
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_HANDLE_ADD",
        Category = "MultiTableAddin",
        Description = "基于对象句柄累加数值。")]
    public static double AddHandleValue(
        [ExcelHandle][ExcelArgument(Name = "handle", Description = "由 HANDLE_CREATE 返回的对象句柄。")] DemoHandleSession handle,
        [ExcelArgument(Name = "delta", Description = "增量值。")] double delta)
    {
        if (handle == null)
        {
            return double.NaN;
        }

        handle.Value += delta;
        return handle.Value;
    }

    [ExcelFunction(
        Name = "MULTITABLEADDIN_LOGPATH",
        Category = "MultiTableAddin",
        Description = "返回插件日志路径。")]
    public static string LogPath()
    {
        return AddInLog.LogFilePath;
    }

    private static object[,] CreateTable(int rows, int cols, string prefix)
    {
        object[,] result = new object[rows, cols];
        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < cols; col++)
            {
                result[row, col] = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0}[{1},{2}] {3:HH:mm:ss}",
                    prefix,
                    row,
                    col,
                    DateTime.Now);
            }
        }

        return result;
    }
}

[ExcelHandle]
public sealed class DemoHandleSession
{
    public DemoHandleSession(string owner, double value)
    {
        Owner = owner;
        Value = value;
        CreatedAt = DateTime.Now;
    }

    public string Owner { get; }

    public DateTime CreatedAt { get; }

    public double Value { get; set; }
}

[ComVisible(true)]
[ProgId(DemoRtdServer.ServerProgId)]
public class DemoRtdServer : ExcelRtdServer
{
    internal const string ServerProgId = "MultiTableAddin.LiveDataRtdServer";
    private readonly object _syncRoot = new();
    private readonly Dictionary<Topic, IDisposable> _subscriptions = new();

    protected override bool ServerStart()
    {
        AddInLog.Write("RTD.ServerStart");
        return true;
    }

    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        string topicName = topicInfo.FirstOrDefault() ?? string.Empty;
        IDisposable subscription = CreateSubscription(topic, topicName, topicInfo);
        lock (_syncRoot)
        {
            _subscriptions[topic] = subscription;
        }

        AddInLog.Write("RTD.ConnectData", string.Join("|", topicInfo));
        return "数据加载中...";
    }

    protected override void DisconnectData(Topic topic)
    {
        lock (_syncRoot)
        {
            if (_subscriptions.TryGetValue(topic, out IDisposable? subscription))
            {
                subscription.Dispose();
                _subscriptions.Remove(topic);
            }
        }
    }

    protected override void ServerTerminate()
    {
        lock (_syncRoot)
        {
            foreach (IDisposable subscription in _subscriptions.Values)
            {
                subscription.Dispose();
            }

            _subscriptions.Clear();
        }

        AddInLog.Write("RTD.ServerTerminate");
    }

    private static IDisposable CreateSubscription(Topic topic, string topicName, IList<string> topicInfo)
    {
        if (string.Equals(topicName, "CLOCK", StringComparison.OrdinalIgnoreCase))
        {
            return new TimerTopicSubscription(
                TimeSpan.FromSeconds(1),
                () => DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
                topic.UpdateValue);
        }

        if (string.Equals(topicName, "WAVE", StringComparison.OrdinalIgnoreCase))
        {
            double amplitude = ParseDouble(topicInfo, 1, 1d);
            double speed = ParseDouble(topicInfo, 2, 0.4d);
            double phase = 0d;
            return new TimerTopicSubscription(
                TimeSpan.FromMilliseconds(500),
                () =>
                {
                    phase += Math.Max(0.15d, Math.Abs(speed));
                    return Math.Round(amplitude * Math.Sin(phase), 4);
                },
                topic.UpdateValue);
        }

        throw new ArgumentOutOfRangeException(nameof(topicInfo), "未知 RTD 主题：" + topicName);
    }

    private static double ParseDouble(IList<string> topicInfo, int index, double fallback)
    {
        if (topicInfo.Count <= index)
        {
            return fallback;
        }

        return double.TryParse(topicInfo[index], NumberStyles.Float, CultureInfo.InvariantCulture, out double result)
            ? result
            : fallback;
    }
}

internal sealed class TimerTopicSubscription : IDisposable
{
    private readonly System.Threading.Timer _timer;
    private readonly Func<object> _valueFactory;
    private readonly Action<object> _onValue;

    internal TimerTopicSubscription(TimeSpan interval, Func<object> valueFactory, Action<object> onValue)
    {
        _valueFactory = valueFactory;
        _onValue = onValue;
        _timer = new System.Threading.Timer(OnTick, null, TimeSpan.Zero, interval);
    }

    public void Dispose()
    {
        _timer.Dispose();
    }

    private void OnTick(object? state)
    {
        try
        {
            _onValue(_valueFactory());
        }
        catch (Exception ex)
        {
            AddInLog.Write("RTD.Tick.Error", ex.ToString());
        }
    }
}
