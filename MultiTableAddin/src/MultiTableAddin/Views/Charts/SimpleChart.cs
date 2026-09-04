using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;
using MultiTableAddin.Core;
using Pen = System.Windows.Media.Pen;
using Size = System.Windows.Size;
using FontFamily = System.Windows.Media.FontFamily;
using ColorConverter = System.Windows.Media.ColorConverter;

namespace MultiTableAddin.Views.Charts;

/// <summary>
/// 轻量级自绘图表控件。
/// 只依赖 WPF 自身的 DrawingContext，不引入 SkiaSharp 等原生依赖，
/// 避免在 Excel / WPS 宿主进程中出现原生 DLL 加载失败。
/// 支持：柱状图 / 条形图 / 折线图 / 面积图 / 饼图 / 环形图 / 仪表盘。
/// </summary>
public class SimpleChart : FrameworkElement
{
    // ── 主题色 ────────────────────────────────────────────────
    private static readonly Brush AxisBrush = Frozen("#4A5163");
    private static readonly Brush GridBrush = Frozen("#333947");
    private static readonly Brush TextBrush = Frozen("#B8BFCC");
    private static readonly Brush StrongTextBrush = Frozen("#E8ECF4");
    private static readonly Brush EmptyTextBrush = Frozen("#6C7488");
    private static readonly Brush TrackBrush = Frozen("#2C3140");

    private static readonly Typeface Font =
        new(new FontFamily("Microsoft YaHei UI, Microsoft YaHei, Segoe UI"),
            FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);

    private const double PadLeft = 58;
    private const double PadRight = 16;
    private const double PadTop = 14;
    private const double PadBottom = 42;
    private const double LegendHeight = 22;

    private double _dpi = 1.0;
    private readonly List<HitArea> _hitAreas = new();

    private sealed class HitArea
    {
        public Rect Bounds;
        public string Text = string.Empty;
    }

    private static Brush Frozen(string hex)
    {
        var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        b.Freeze();
        return b;
    }

    private static Brush BrushOf(string hex, double opacity = 1.0)
    {
        try
        {
            var c = (Color)ColorConverter.ConvertFromString(hex);
            var b = new SolidColorBrush(c) { Opacity = opacity };
            b.Freeze();
            return b;
        }
        catch { return Brushes.SteelBlue; }
    }

    // ── 依赖属性 ──────────────────────────────────────────────

    public static readonly DependencyProperty DataProperty = DependencyProperty.Register(
        nameof(Data), typeof(ChartDataSet), typeof(SimpleChart),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender, OnDataChanged));

    public ChartDataSet? Data
    {
        get => (ChartDataSet?)GetValue(DataProperty);
        set => SetValue(DataProperty, value);
    }

    public static readonly DependencyProperty ChartTypeProperty = DependencyProperty.Register(
        nameof(ChartType), typeof(ChartType), typeof(SimpleChart),
        new FrameworkPropertyMetadata(ChartType.Column, FrameworkPropertyMetadataOptions.AffectsRender));

    public ChartType ChartType
    {
        get => (ChartType)GetValue(ChartTypeProperty);
        set => SetValue(ChartTypeProperty, value);
    }

    public static readonly DependencyProperty GaugeTargetProperty = DependencyProperty.Register(
        nameof(GaugeTarget), typeof(double), typeof(SimpleChart),
        new FrameworkPropertyMetadata(100d, FrameworkPropertyMetadataOptions.AffectsRender));

    public double GaugeTarget
    {
        get => (double)GetValue(GaugeTargetProperty);
        set => SetValue(GaugeTargetProperty, value);
    }

    /// <summary>0→1 的入场动画进度</summary>
    public static readonly DependencyProperty ProgressProperty = DependencyProperty.Register(
        nameof(Progress), typeof(double), typeof(SimpleChart),
        new FrameworkPropertyMetadata(1d, FrameworkPropertyMetadataOptions.AffectsRender));

    public double Progress
    {
        get => (double)GetValue(ProgressProperty);
        set => SetValue(ProgressProperty, value);
    }

    /// <summary>是否播放入场动画</summary>
    public bool Animated { get; set; } = true;

    public SimpleChart()
    {
        ClipToBounds = true;
        SnapsToDevicePixels = true;
    }

    private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is SimpleChart chart && chart.Animated)
            chart.PlayAnimation();
    }

    private void PlayAnimation()
    {
        BeginAnimation(ProgressProperty, null);
        Progress = 0;
        var anim = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(420))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        BeginAnimation(ProgressProperty, anim);
    }

    // ── 交互提示 ──────────────────────────────────────────────

    protected override void OnMouseMove(System.Windows.Input.MouseEventArgs e)
    {
        base.OnMouseMove(e);
        var p = e.GetPosition(this);
        var hit = _hitAreas.FirstOrDefault(h => h.Bounds.Contains(p));
        string? tip = hit?.Text;
        if (!Equals(ToolTip, tip)) ToolTip = tip;
    }

    protected override void OnMouseLeave(System.Windows.Input.MouseEventArgs e)
    {
        base.OnMouseLeave(e);
        ToolTip = null;
    }

    protected override HitTestResult HitTestCore(PointHitTestParameters p) =>
        new PointHitTestResult(this, p.HitPoint);

    // ── 绘制入口 ──────────────────────────────────────────────

    protected override void OnRender(DrawingContext dc)
    {
        base.OnRender(dc);
        _hitAreas.Clear();
        _dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;

        double w = ActualWidth, h = ActualHeight;
        if (w < 40 || h < 40) return;

        // 透明底以接收鼠标事件
        dc.DrawRectangle(Brushes.Transparent, null, new Rect(0, 0, w, h));

        var data = Data;
        if (data == null || data.IsEmpty)
        {
            DrawCenterText(dc, data?.Message is { Length: > 0 } m ? m : "暂无数据", w, h);
            return;
        }

        switch (ChartType)
        {
            case ChartType.Pie:
                DrawPie(dc, data, w, h, false);
                break;
            case ChartType.Doughnut:
                DrawPie(dc, data, w, h, true);
                break;
            case ChartType.Gauge:
                DrawGauge(dc, data, w, h);
                break;
            case ChartType.Bar:
                DrawHorizontalBars(dc, data, w, h);
                break;
            case ChartType.Line:
                DrawLineOrArea(dc, data, w, h, false);
                break;
            case ChartType.Area:
                DrawLineOrArea(dc, data, w, h, true);
                break;
            default:
                DrawColumns(dc, data, w, h);
                break;
        }
    }

    // ── 通用辅助 ──────────────────────────────────────────────

    private FormattedText Text(string s, double size, Brush brush, FontWeight? weight = null)
    {
        var tf = weight == null
            ? Font
            : new Typeface(Font.FontFamily, FontStyles.Normal, weight.Value, FontStretches.Normal);

        return new FormattedText(s, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            tf, size, brush, _dpi);
    }

    private void DrawCenterText(DrawingContext dc, string msg, double w, double h)
    {
        var ft = Text(msg, 13, EmptyTextBrush);
        dc.DrawText(ft, new Point((w - ft.Width) / 2, (h - ft.Height) / 2));
    }

    /// <summary>算出刻度上限与步长，让 Y 轴刻度是好看的整数</summary>
    private static (double niceMax, double step) NiceScale(double max, int ticks = 4)
    {
        if (max <= 0) return (1, 1);
        double raw = max / ticks;
        double mag = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        double norm = raw / mag;
        double stepNorm = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
        double step = stepNorm * mag;
        return (Math.Ceiling(max / step) * step, step);
    }

    private double DrawLegend(DrawingContext dc, ChartDataSet data, double w)
    {
        if (data.Series.Count <= 1) return 0;

        double x = PadLeft, y = 2;
        foreach (var s in data.Series)
        {
            var ft = Text(s.Name, 11, TextBrush);
            if (x + 16 + ft.Width > w - PadRight) break;

            dc.DrawRoundedRectangle(BrushOf(s.Color), null, new Rect(x, y + 4, 10, 10), 2, 2);
            dc.DrawText(ft, new Point(x + 14, y + 1));
            x += 14 + ft.Width + 16;
        }
        return LegendHeight;
    }

    private void DrawAxes(DrawingContext dc, Rect plot, double niceMax, double step)
    {
        var axisPen = new Pen(AxisBrush, 1);
        var gridPen = new Pen(GridBrush, 1) { DashStyle = new DashStyle(new double[] { 3, 3 }, 0) };

        for (double v = 0; v <= niceMax + step * 0.01; v += step)
        {
            double y = plot.Bottom - (v / niceMax) * plot.Height;
            y = Math.Round(y) + 0.5;
            dc.DrawLine(v == 0 ? axisPen : gridPen, new Point(plot.Left, y), new Point(plot.Right, y));

            var ft = Text(ValueFormatter.ToAxisLabel(v), 10, TextBrush);
            dc.DrawText(ft, new Point(plot.Left - ft.Width - 6, y - ft.Height / 2));
        }
    }

    /// <summary>横轴标签，过密时抽稀，过长时旋转</summary>
    private void DrawCategoryLabels(DrawingContext dc, ChartDataSet data, Rect plot, double slotWidth)
    {
        int n = data.Categories.Count;
        if (n == 0) return;

        // 估算每个标签需要的宽度，决定抽稀间隔
        double maxTextW = data.Categories.Max(c => Text(c, 10, TextBrush).Width);
        int stepIdx = 1;
        if (maxTextW + 6 > slotWidth)
            stepIdx = (int)Math.Ceiling((maxTextW + 6) / slotWidth);

        bool rotate = maxTextW > slotWidth * 1.6 && n > 3;

        for (int i = 0; i < n; i++)
        {
            if (i % stepIdx != 0 && i != n - 1) continue;

            var ft = Text(data.Categories[i], 10, TextBrush);
            double cx = plot.Left + slotWidth * (i + 0.5);

            if (rotate)
            {
                var origin = new Point(cx + ft.Height / 2, plot.Bottom + 6);
                dc.PushTransform(new RotateTransform(35, origin.X, origin.Y));
                dc.DrawText(ft, origin);
                dc.Pop();
            }
            else
            {
                dc.DrawText(ft, new Point(cx - ft.Width / 2, plot.Bottom + 6));
            }
        }
    }

    private void AddHit(Rect bounds, string text) => _hitAreas.Add(new HitArea { Bounds = bounds, Text = text });

    // ── 柱状图 ────────────────────────────────────────────────

    private void DrawColumns(DrawingContext dc, ChartDataSet data, double w, double h)
    {
        double legendH = DrawLegend(dc, data, w);
        var plot = new Rect(PadLeft, PadTop + legendH, w - PadLeft - PadRight, h - PadTop - PadBottom - legendH);
        if (plot.Width < 20 || plot.Height < 20) return;

        var (niceMax, step) = NiceScale(Math.Max(data.MaxValue, 1e-9));
        DrawAxes(dc, plot, niceMax, step);

        int cats = data.Categories.Count;
        int seriesCount = data.Series.Count;
        double slot = plot.Width / Math.Max(cats, 1);
        double groupW = slot * 0.68;
        double barW = Math.Max(3, groupW / Math.Max(seriesCount, 1));
        double p = Progress;

        for (int si = 0; si < seriesCount; si++)
        {
            var s = data.Series[si];
            var fill = BrushOf(s.Color);
            for (int i = 0; i < s.Points.Count && i < cats; i++)
            {
                double val = s.Points[i].Value;
                double hgt = Math.Max(0, val / niceMax * plot.Height) * p;
                double x = plot.Left + slot * i + (slot - groupW) / 2 + barW * si;
                double y = plot.Bottom - hgt;

                var rect = new Rect(x, y, Math.Max(1, barW - 2), hgt);
                dc.DrawRoundedRectangle(fill, null, rect, 3, 3);

                AddHit(new Rect(x, plot.Top, Math.Max(6, barW), plot.Height),
                    $"{data.Categories[i]}\n{s.Name}：{ValueFormatter.ToCompactNumber(val)}");

                // 数据量不大时直接标数值
                if (cats * seriesCount <= 14 && hgt > 12)
                {
                    var ft = Text(ValueFormatter.ToAxisLabel(val), 10, StrongTextBrush);
                    dc.DrawText(ft, new Point(x + (barW - 2 - ft.Width) / 2, y - ft.Height - 2));
                }
            }
        }

        DrawCategoryLabels(dc, data, plot, slot);
    }

    // ── 条形图（横向）──────────────────────────────────────────

    private void DrawHorizontalBars(DrawingContext dc, ChartDataSet data, double w, double h)
    {
        double legendH = DrawLegend(dc, data, w);

        // 左侧留给类别名
        double labelW = Math.Min(140, Math.Max(60,
            data.Categories.Count == 0 ? 60 : data.Categories.Max(c => Text(c, 11, TextBrush).Width) + 10));

        var plot = new Rect(labelW + 8, PadTop + legendH, w - labelW - 8 - PadRight - 40, h - PadTop - legendH - 24);
        if (plot.Width < 20 || plot.Height < 20) return;

        var (niceMax, _) = NiceScale(Math.Max(data.MaxValue, 1e-9));
        int cats = data.Categories.Count;
        int seriesCount = data.Series.Count;
        double slot = plot.Height / Math.Max(cats, 1);
        double groupH = slot * 0.66;
        double barH = Math.Max(3, groupH / Math.Max(seriesCount, 1));
        double p = Progress;

        var axisPen = new Pen(AxisBrush, 1);
        dc.DrawLine(axisPen, new Point(plot.Left - 0.5, plot.Top), new Point(plot.Left - 0.5, plot.Bottom));

        for (int i = 0; i < cats; i++)
        {
            var ft = Text(data.Categories[i], 11, TextBrush);
            double cy = plot.Top + slot * (i + 0.5);
            dc.DrawText(ft, new Point(Math.Max(2, labelW - ft.Width), cy - ft.Height / 2));
        }

        for (int si = 0; si < seriesCount; si++)
        {
            var s = data.Series[si];
            var fill = BrushOf(s.Color);
            for (int i = 0; i < s.Points.Count && i < cats; i++)
            {
                double val = s.Points[i].Value;
                double len = Math.Max(0, val / niceMax * plot.Width) * p;
                double y = plot.Top + slot * i + (slot - groupH) / 2 + barH * si;

                dc.DrawRoundedRectangle(fill, null,
                    new Rect(plot.Left, y, Math.Max(1, len), Math.Max(1, barH - 2)), 3, 3);

                AddHit(new Rect(plot.Left, y, Math.Max(len, 6), Math.Max(6, barH)),
                    $"{data.Categories[i]}\n{s.Name}：{ValueFormatter.ToCompactNumber(val)}");

                if (barH >= 12)
                {
                    var ft = Text(ValueFormatter.ToAxisLabel(val), 10, StrongTextBrush);
                    dc.DrawText(ft, new Point(plot.Left + len + 5, y + (barH - 2 - ft.Height) / 2));
                }
            }
        }
    }

    // ── 折线图 / 面积图 ────────────────────────────────────────

    private void DrawLineOrArea(DrawingContext dc, ChartDataSet data, double w, double h, bool area)
    {
        double legendH = DrawLegend(dc, data, w);
        var plot = new Rect(PadLeft, PadTop + legendH, w - PadLeft - PadRight, h - PadTop - PadBottom - legendH);
        if (plot.Width < 20 || plot.Height < 20) return;

        var (niceMax, step) = NiceScale(Math.Max(data.MaxValue, 1e-9));
        DrawAxes(dc, plot, niceMax, step);

        int cats = data.Categories.Count;
        double slot = plot.Width / Math.Max(cats, 1);
        double p = Progress;

        foreach (var s in data.Series)
        {
            if (s.Points.Count == 0) continue;

            var pts = new List<Point>();
            for (int i = 0; i < s.Points.Count && i < cats; i++)
            {
                double val = s.Points[i].Value;
                double x = plot.Left + slot * (i + 0.5);
                double y = plot.Bottom - (val / niceMax) * plot.Height * p;
                pts.Add(new Point(x, y));
            }
            if (pts.Count == 0) continue;

            var stroke = new Pen(BrushOf(s.Color), 2)
            {
                LineJoin = PenLineJoin.Round,
                StartLineCap = PenLineCap.Round,
                EndLineCap = PenLineCap.Round
            };

            if (area)
            {
                var geo = new StreamGeometry();
                using (var ctx = geo.Open())
                {
                    ctx.BeginFigure(new Point(pts[0].X, plot.Bottom), true, true);
                    ctx.LineTo(pts[0], true, false);
                    for (int i = 1; i < pts.Count; i++) ctx.LineTo(pts[i], true, true);
                    ctx.LineTo(new Point(pts[^1].X, plot.Bottom), true, false);
                }
                geo.Freeze();

                var grad = new LinearGradientBrush(
                    ((SolidColorBrush)BrushOf(s.Color, 0.42)).Color,
                    Color.FromArgb(8, 0, 0, 0),
                    new Point(0, 0), new Point(0, 1));
                grad.Freeze();
                dc.DrawGeometry(grad, null, geo);
            }

            for (int i = 1; i < pts.Count; i++)
                dc.DrawLine(stroke, pts[i - 1], pts[i]);

            bool showDots = pts.Count <= 40;
            for (int i = 0; i < pts.Count; i++)
            {
                if (showDots)
                {
                    dc.DrawEllipse(Brushes.White, stroke, pts[i], 3, 3);
                    dc.DrawEllipse(BrushOf(s.Color), null, pts[i], 1.6, 1.6);
                }
                AddHit(new Rect(pts[i].X - slot / 2, plot.Top, Math.Max(slot, 8), plot.Height),
                    $"{data.Categories[Math.Min(i, cats - 1)]}\n{s.Name}：{ValueFormatter.ToCompactNumber(s.Points[i].Value)}");
            }
        }

        DrawCategoryLabels(dc, data, plot, slot);
    }

    // ── 饼图 / 环形图 ──────────────────────────────────────────

    private void DrawPie(DrawingContext dc, ChartDataSet data, double w, double h, bool doughnut)
    {
        var points = data.Series.SelectMany(s => s.Points).Where(pt => pt.Value > 0).ToList();
        double total = points.Sum(pt => pt.Value);
        if (total <= 0)
        {
            DrawCenterText(dc, "数据合计为 0，无法绘制占比图", w, h);
            return;
        }

        // 右侧留给图例
        double legendW = Math.Min(160, w * 0.38);
        double areaW = w - legendW;
        double radius = Math.Max(20, Math.Min(areaW, h) / 2 - 16);
        var center = new Point(areaW / 2, h / 2);
        double inner = doughnut ? radius * 0.58 : 0;

        double angle = -90;
        double p = Progress;

        for (int i = 0; i < points.Count; i++)
        {
            double sweep = points[i].Value / total * 360 * p;
            if (sweep <= 0.01) continue;

            var brush = BrushOf(ChartDataBuilder.ColorAt(i));
            dc.DrawGeometry(brush, null, BuildSlice(center, radius, inner, angle, sweep));
            angle += sweep;
        }

        if (doughnut)
        {
            var ft = Text(ValueFormatter.ToCompactNumber(total), 17, StrongTextBrush, FontWeights.SemiBold);
            dc.DrawText(ft, new Point(center.X - ft.Width / 2, center.Y - ft.Height / 2 - 7));
            var ft2 = Text("合计", 10, TextBrush);
            dc.DrawText(ft2, new Point(center.X - ft2.Width / 2, center.Y + ft.Height / 2 - 6));
        }

        // 图例
        double ly = Math.Max(8, (h - points.Count * 20) / 2);
        double lx = areaW + 6;
        for (int i = 0; i < points.Count && ly < h - 16; i++)
        {
            dc.DrawRoundedRectangle(BrushOf(ChartDataBuilder.ColorAt(i)), null,
                new Rect(lx, ly + 4, 10, 10), 2, 2);

            double pct = points[i].Value / total * 100;
            string label = Ellipsis(points[i].Label, 8);
            var ft = Text($"{label}  {pct:0.#}%", 11, TextBrush);
            dc.DrawText(ft, new Point(lx + 15, ly + 1));

            AddHit(new Rect(lx, ly, legendW, 18),
                $"{points[i].Label}\n{ValueFormatter.ToCompactNumber(points[i].Value)}（{pct:0.##}%）");
            ly += 20;
        }
    }

    private static string Ellipsis(string s, int maxChars) =>
        s.Length <= maxChars ? s : s.Substring(0, maxChars) + "…";

    private static Geometry BuildSlice(Point c, double outerR, double innerR, double startDeg, double sweepDeg)
    {
        double s = startDeg * Math.PI / 180;
        double e = (startDeg + sweepDeg) * Math.PI / 180;
        bool large = sweepDeg > 180;

        Point PO(double a, double r) => new(c.X + r * Math.Cos(a), c.Y + r * Math.Sin(a));

        var geo = new StreamGeometry();
        using (var ctx = geo.Open())
        {
            if (innerR <= 0.01)
            {
                ctx.BeginFigure(c, true, true);
                ctx.LineTo(PO(s, outerR), true, false);
                ctx.ArcTo(PO(e, outerR), new Size(outerR, outerR), 0, large,
                    SweepDirection.Clockwise, true, false);
            }
            else
            {
                ctx.BeginFigure(PO(s, innerR), true, true);
                ctx.LineTo(PO(s, outerR), true, false);
                ctx.ArcTo(PO(e, outerR), new Size(outerR, outerR), 0, large,
                    SweepDirection.Clockwise, true, false);
                ctx.LineTo(PO(e, innerR), true, false);
                ctx.ArcTo(PO(s, innerR), new Size(innerR, innerR), 0, large,
                    SweepDirection.Counterclockwise, true, false);
            }
        }
        geo.Freeze();
        return geo;
    }

    // ── 仪表盘 ────────────────────────────────────────────────

    private void DrawGauge(DrawingContext dc, ChartDataSet data, double w, double h)
    {
        double value = data.Series.FirstOrDefault()?.Points.FirstOrDefault()?.Value ?? 0;
        double target = GaugeTarget <= 0 ? Math.Max(value, 1) : GaugeTarget;
        double ratio = Math.Max(0, Math.Min(1.4, value / target));

        double radius = Math.Max(24, Math.Min(w / 2 - 20, h - 46));
        var center = new Point(w / 2, h - 26);
        double thickness = Math.Max(10, radius * 0.22);
        double inner = radius - thickness;

        // 底环 180°
        dc.DrawGeometry(TrackBrush, null, BuildSlice(center, radius, inner, 180, 180));

        double sweep = 180 * Math.Min(ratio, 1.0) * Progress;
        string color = ratio >= 1 ? "#27AE60" : ratio >= 0.7 ? "#F2C94C" : "#EB5757";
        if (sweep > 0.01)
            dc.DrawGeometry(BrushOf(color), null, BuildSlice(center, radius, inner, 180, sweep));

        var vt = Text(ValueFormatter.ToCompactNumber(value), 22, StrongTextBrush, FontWeights.SemiBold);
        dc.DrawText(vt, new Point(center.X - vt.Width / 2, center.Y - radius * 0.52));

        var pt = Text($"{ratio * 100:0.#}% / 目标 {ValueFormatter.ToCompactNumber(target)}", 11, TextBrush);
        dc.DrawText(pt, new Point(center.X - pt.Width / 2, center.Y - radius * 0.52 + vt.Height + 2));

        AddHit(new Rect(0, 0, w, h),
            $"{data.Title}\n当前：{ValueFormatter.ToCompactNumber(value)}\n目标：{ValueFormatter.ToCompactNumber(target)}\n完成率：{ratio * 100:0.##}%");
    }
}
