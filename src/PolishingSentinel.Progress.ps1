if (-not ('SentinelGradientProgress' -as [type])) {
    $references = @([Windows.Forms.Control].Assembly.Location, [Drawing.Graphics].Assembly.Location, [Drawing.Color].Assembly.Location) | Select-Object -Unique
    if ($PSVersionTable.PSEdition -eq 'Core') { $references += @('System.Runtime', 'System.ComponentModel.Primitives', 'System.ComponentModel.TypeConverter') }
    Add-Type -ReferencedAssemblies $references -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
public sealed class SentinelGradientProgress : Control {
    private int current, maximum = 100, speed, offset;
    private ProgressBarStyle style = ProgressBarStyle.Continuous;
    private readonly Timer timer = new Timer();
    public SentinelGradientProgress() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        Height = 22; TabStop = false; AccessibleName = "Processing progress";
        timer.Tick += delegate { offset = (offset + 8) % Math.Max(1, Width + Width / 4); Invalidate(); };
    }
    public int Maximum { get { return maximum; } set { maximum = Math.Max(1, value); Invalidate(); } }
    public int Value { get { return current; } set { int next = Math.Max(0, Math.Min(maximum, value)); if (next != current) { current = next; Invalidate(); } } }
    public ProgressBarStyle Style { get { return style; } set { if (style != value) { style = value; UpdateTimer(); Invalidate(); } } }
    public int MarqueeAnimationSpeed { get { return speed; } set { speed = value; UpdateTimer(); } }
    private void UpdateTimer() {
        timer.Stop();
        if (style == ProgressBarStyle.Marquee && speed > 0 && Visible) { timer.Interval = Math.Max(16, speed); timer.Start(); }
    }
    protected override void OnVisibleChanged(EventArgs e) { base.OnVisibleChanged(e); UpdateTimer(); }
    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        e.Graphics.Clear(Color.FromArgb(224, 233, 245));
        int width = style == ProgressBarStyle.Marquee ? Math.Max(24, Width / 4) : (int)((long)Width * current / maximum);
        int x = style == ProgressBarStyle.Marquee ? offset - width : 0;
        if (width > 0 && Height > 0) {
            using (var brush = new LinearGradientBrush(new Rectangle(x, 0, width, Height), Color.FromArgb(29, 78, 216), Color.FromArgb(56, 189, 248), LinearGradientMode.Horizontal))
                e.Graphics.FillRectangle(brush, x, 0, width, Height);
        }
    }
    protected override void Dispose(bool disposing) { if (disposing) timer.Dispose(); base.Dispose(disposing); }
}
'@
}
