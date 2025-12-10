using System;
using System.Windows.Forms;
using System.Drawing;

namespace SampleProject1
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();

            // ===== Window settings =====
            this.Text = "Processing...";
            this.StartPosition = FormStartPosition.CenterParent;
            this.ControlBox = false;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // ===== Status label (fixed position, perfect center) =====
            lblStatus.AutoSize = false;
            lblStatus.TextAlign = ContentAlignment.MiddleCenter;     // horizontal + vertical center
            lblStatus.Dock = DockStyle.Top;
            lblStatus.Height = 100;
            lblStatus.AutoEllipsis = true;
            lblStatus.Padding = new Padding(5);

            // ===== Progress bar (80% width, centered) =====
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.MarqueeAnimationSpeed = 30;

            progressBar1.Dock = DockStyle.None;   // remove full-width docking
            progressBar1.Height = 22;

            // Set width = 80% of form
            int barWidth = (int)(this.Width * 0.80);
            progressBar1.Width = barWidth;

            // Center horizontally
            progressBar1.Left = (this.ClientSize.Width - progressBar1.Width) / 2;

            // Vertical position (below label)
            progressBar1.Top = lblStatus.Bottom + 10;

            // Ensure it remains centered when resizing
            this.Resize += (s, e) =>
            {
                progressBar1.Left = (this.ClientSize.Width - progressBar1.Width) / 2;
            };

            lblStatus.Text = "Starting...";
        }

        public void UpdateStatus(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(UpdateStatus), message);
                return;
            }

            lblStatus.Text = message;
        }

        private void lblStatus_Click(object sender, EventArgs e)
        {
            // No click action needed.
        }
    }
}
