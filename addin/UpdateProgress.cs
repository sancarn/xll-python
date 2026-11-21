using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PythonExcel
{
    internal sealed class UpdateProgress : Form
    {
        readonly Timer timer = new Timer { Interval = 50 };
        int angle;

        internal UpdateProgress()
        {
            Text = "Checking for updates";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(16);

            var status = new Label {
                Text = "Checking for updates\u2026", AutoSize = true,
                Padding = new Padding(Font.Height * 2, 4, 0, 4)
            };
            status.Paint += (sender, e) => {
                int size = status.Font.Height;
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(SystemColors.Highlight, Math.Max(2, size / 6)))
                    e.Graphics.DrawArc(pen, 2, (status.Height - size) / 2, size, size, angle, 270);
            };
            Controls.Add(status);
            status.Location = new Point(Padding.Left, Padding.Top);
            timer.Tick += (sender, e) => {
                angle = (angle + 18) % 360;
                status.Invalidate();
            };
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            CenterToParent();
            timer.Start();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) timer.Dispose();
            base.Dispose(disposing);
        }
    }
}
