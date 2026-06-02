using System;
using System.Drawing;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using fyiReporting.RDL;

namespace fyiReporting.RdlDesign
{
    internal sealed class DialogExportProgress : Form
    {
        private readonly OutputPresentationType _type;
        private readonly Action _work;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly DateTime _started;
        private readonly Label _statusLabel;
        private readonly Label _elapsedLabel;
        private readonly Label _remainingLabel;
        private readonly ProgressBar _progress;
        private readonly Button _cancelButton;
        private bool _progressBarIsContinuous;
        private bool _cancelling;
        private DateTime _phaseStarted;
        private string _lastPhase;
        private Exception _error;

        public DialogExportProgress(OutputPresentationType type, Action work)
        {
            _type = type;
            _work = work ?? throw new ArgumentNullException(nameof(work));

            Text = string.Format("Export in {0}", FriendlyName(type));
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ControlBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(380, 168);

            _statusLabel = new Label
            {
                Text = "Prepare...",
                Location = new Point(12, 12),
                AutoSize = false,
                Size = new Size(356, 18),
                Font = new Font(Font, FontStyle.Bold)
            };

            _progress = new ProgressBar
            {
                Location = new Point(12, 40),
                Size = new Size(356, 22),
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Minimum = 0,
                Maximum = 100
            };

            _elapsedLabel = new Label
            {
                Text = "Passed: 0:00",
                Location = new Point(12, 74),
                AutoSize = true
            };

            _remainingLabel = new Label
            {
                Text = "Remains: —",
                Location = new Point(200, 74),
                AutoSize = true,
                ForeColor = SystemColors.GrayText
            };

            _cancelButton = new Button
            {
                Text = "Cancel",
                Size = new Size(100, 28),
                Location = new Point(268, 128),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _cancelButton.Click += OnCancelClick;

            Controls.Add(_statusLabel);
            Controls.Add(_progress);
            Controls.Add(_elapsedLabel);
            Controls.Add(_remainingLabel);
            Controls.Add(_cancelButton);

            _timer = new System.Windows.Forms.Timer { Interval = 250 };
            _timer.Tick += OnTick;

            _started = DateTime.UtcNow;
            _phaseStarted = _started;
            Shown += OnShown;
            FormClosed += (s, e) => _timer.Dispose();
        }

        public static void Run(IWin32Window owner, OutputPresentationType type, Action work)
        {
            using (var dlg = new DialogExportProgress(type, work))
            {
                dlg.ShowDialog(owner);
                if (dlg._error != null)
                    throw dlg._error;
            }
        }

        private void OnCancelClick(object sender, EventArgs e)
        {
            ExportProgress.RequestCancel();
            _cancelling = true;
            _cancelButton.Enabled = false;
            _statusLabel.Text = "Cancelation...";
        }

        private async void OnShown(object sender, EventArgs e)
        {
            ExportProgress.Reset();
            _timer.Start();
            try
            {
                await Task.Run(_work).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                // User pressed Cancel
            }
            catch (Exception ex)
            {
                _error = ex;
            }
            finally
            {
                _timer.Stop();
                Close();
            }
        }

        private void OnTick(object sender, EventArgs e)
        {
            DateTime now = DateTime.UtcNow;
            TimeSpan elapsed = now - _started;
            _elapsedLabel.Text = "Passed: " + Format(elapsed);

            if (_cancelling)
                return;

            int current = ExportProgress.Current;
            int total = ExportProgress.Total;
            string phase = ExportProgress.Phase;

            // Detect phase change
            if (phase != _lastPhase)
            {
                _lastPhase = phase;
                _phaseStarted = now;
            }
            TimeSpan phaseElapsed = now - _phaseStarted;

            bool wantContinuous = total > 0;
            if (wantContinuous && !_progressBarIsContinuous)
            {
                _progress.Style = ProgressBarStyle.Continuous;
                _progressBarIsContinuous = true;
            }
            else if (!wantContinuous && _progressBarIsContinuous)
            {
                _progress.Style = ProgressBarStyle.Marquee;
                _progress.MarqueeAnimationSpeed = 30;
                _progressBarIsContinuous = false;
            }

            string phasePrefix = string.IsNullOrEmpty(phase) ? "Processed" : phase;

            if (total > 0)
            {
                int pct = (int)Math.Min(100, (current * 100L) / total);
                _progress.Value = pct;
                _statusLabel.Text = string.Format(
                    "{0}: {1} / {2} ({3}%)",
                    phasePrefix, FormatNumber(current), FormatNumber(total), pct);

                if (current > 0 && phaseElapsed.TotalSeconds > 0.5)
                {
                    double remainingSeconds = phaseElapsed.TotalSeconds * (total - current) / current;
                    if (remainingSeconds < 0) remainingSeconds = 0;
                    _remainingLabel.Text = "Remains: ~" + Format(TimeSpan.FromSeconds(remainingSeconds));
                    _remainingLabel.ForeColor = SystemColors.ControlText;
                }
            }
            else if (current > 0)
            {
                _statusLabel.Text = string.Format("{0}: {1}", phasePrefix, FormatNumber(current));
                _remainingLabel.Text = "Remains: —";
                _remainingLabel.ForeColor = SystemColors.GrayText;
            }
            else if (!string.IsNullOrEmpty(phase))
            {
                _statusLabel.Text = phase;
                _remainingLabel.Text = "Remains: —";
                _remainingLabel.ForeColor = SystemColors.GrayText;
            }
            else
            {
                _statusLabel.Text = "Data preparation...";
            }
        }

        private static string Format(TimeSpan ts)
        {
            if (ts.TotalHours >= 1)
                return string.Format("{0}:{1:00}:{2:00}", (int)ts.TotalHours, ts.Minutes, ts.Seconds);
            return string.Format("{0}:{1:00}", (int)ts.TotalMinutes, ts.Seconds);
        }

        private static string FormatNumber(int n)
            => n.ToString("N0", CultureInfo.CurrentCulture);

        private static string FriendlyName(OutputPresentationType type)
        {
            switch (type)
            {
                case OutputPresentationType.ExcelTableOnly: return "Table Only";
                case OutputPresentationType.Excel2007ClosedXML: return "Optimized Without Shapes";
                case OutputPresentationType.Excel2007NPOI: return "Full";
                default: return type.ToString();
            }
        }
    }
}
