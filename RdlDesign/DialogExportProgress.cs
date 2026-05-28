/* ====================================================================
   Copyright (C) 2004-2008  fyiReporting Software, LLC
   Copyright (C) 2011  Peter Gill <peter@majorsilence.com>

   This file is part of the fyiReporting RDL project.

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0
*/
using System;
using System.Drawing;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using fyiReporting.RDL;

namespace fyiReporting.RdlDesign
{
    /// <summary>
    /// Modal progress dialog shown while a long export operation runs on a background thread
    /// </summary>
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
        private bool _progressBarIsContinuous;
        private DateTime _phaseStarted;
        private string _lastPhase;
        private Exception _error;

        public DialogExportProgress(OutputPresentationType type, Action work)
        {
            _type = type;
            _work = work ?? throw new ArgumentNullException(nameof(work));

            Text = string.Format("Экспорт в {0}", FriendlyName(type));
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ControlBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(380, 130);

            _statusLabel = new Label
            {
                Text = "Подготовка...",
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
                Text = "Прошло: 0:00",
                Location = new Point(12, 74),
                AutoSize = true
            };

            _remainingLabel = new Label
            {
                Text = "Осталось: —",
                Location = new Point(200, 74),
                AutoSize = true,
                ForeColor = SystemColors.GrayText
            };

            Controls.Add(_statusLabel);
            Controls.Add(_progress);
            Controls.Add(_elapsedLabel);
            Controls.Add(_remainingLabel);

            _timer = new System.Windows.Forms.Timer { Interval = 250 };
            _timer.Tick += OnTick;

            _started = DateTime.UtcNow;
            _phaseStarted = _started;
            Shown += OnShown;
            FormClosed += (s, e) => _timer.Dispose();
        }

        /// <summary>
        /// Shows the dialog modally and runs the work delegate on a background thread. Returns when the work completes
        /// </summary>
        public static void Run(IWin32Window owner, OutputPresentationType type, Action work)
        {
            using (var dlg = new DialogExportProgress(type, work))
            {
                dlg.ShowDialog(owner);
                if (dlg._error != null)
                    throw dlg._error;
            }
        }

        private async void OnShown(object sender, EventArgs e)
        {
            ExportProgress.Reset();
            _timer.Start();
            try
            {
                await Task.Run(_work).ConfigureAwait(true);
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
            _elapsedLabel.Text = "Прошло: " + Format(elapsed);

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

            string phasePrefix = string.IsNullOrEmpty(phase) ? "Обработано" : phase;

            if (total > 0)
            {
                int pct = (int)Math.Min(100, (current * 100L) / total);
                _progress.Value = pct;
                _statusLabel.Text = string.Format(
                    "{0}: {1} / {2} ({3}%)",
                    phasePrefix, FormatNumber(current), FormatNumber(total), pct);

                // ETA from observed rate within the current phase
                if (current > 0 && phaseElapsed.TotalSeconds > 0.5)
                {
                    double remainingSeconds = phaseElapsed.TotalSeconds * (total - current) / current;
                    if (remainingSeconds < 0) remainingSeconds = 0;
                    _remainingLabel.Text = "Осталось: ~" + Format(TimeSpan.FromSeconds(remainingSeconds));
                    _remainingLabel.ForeColor = SystemColors.ControlText;
                }
            }
            else if (current > 0)
            {
                _statusLabel.Text = string.Format("{0}: {1}", phasePrefix, FormatNumber(current));
                _remainingLabel.Text = "Осталось: —";
                _remainingLabel.ForeColor = SystemColors.GrayText;
            }
            else if (!string.IsNullOrEmpty(phase))
            {
                _statusLabel.Text = phase;
                _remainingLabel.Text = "Осталось: —";
                _remainingLabel.ForeColor = SystemColors.GrayText;
            }
            else
            {
                _statusLabel.Text = "Подготовка данных...";
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
                case OutputPresentationType.ExcelTableOnly: return "Excel (быстрый)";
                case OutputPresentationType.Excel2007ClosedXML: return "Excel (ClosedXML)";
                case OutputPresentationType.Excel2007NPOI: return "Excel (NPOI)";
                default: return type.ToString();
            }
        }
    }
}
