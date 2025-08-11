using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        int? requestedPt = PromptForFontSize(12);
        if (!requestedPt.HasValue) return;

        int appliedCount = 0;
        int targetFs = Math.Max(1, requestedPt.Value * 2); // RTF \fsN is half-points

        foreach (Track track in vegas.Project.Tracks)
        {
            var vtrack = track as VideoTrack;
            if (vtrack == null) continue;

            foreach (TrackEvent te in vtrack.Events)
            {
                if (!te.Selected) continue;

                var ve = te as VideoEvent;
                if (ve == null) continue;

                var take = ve.ActiveTake;
                if (take == null || take.Media == null || !take.Media.IsGenerated() || take.Media.Generator == null)
                    continue;

                string genName = (take.Media.Generator.PlugIn != null) ? take.Media.Generator.PlugIn.Name : "";
                if (!IsTitlesAndText(genName)) continue;

                OFXEffect ofx = take.Media.Generator.OFXEffect;
                if (ofx == null) continue;

                // The font "size" lives inside the RTF carried by the "Text" string parameter.
                var textParam = ofx.FindParameterByName("Text") as OFXStringParameter;
                if (textParam == null) continue;

                try
                {
                    string rtf = textParam.Value ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(rtf))
                        continue;

                    // Replace all \fsN occurrences (N is half-points)
                    string replaced = Regex.Replace(
                        rtf,
                        @"\\fs\d+",
                        @"\fs" + targetFs,
                        RegexOptions.CultureInvariant
                    );

                    // If no \fs was present, inject one near the start so it applies globally.
                    if (ReferenceEquals(replaced, rtf) || !Regex.IsMatch(rtf, @"\\fs\d+"))
                    {
                        int brace = replaced.IndexOf('{');
                        if (brace >= 0)
                        {
                            replaced = replaced.Insert(brace + 1, @"\fs" + targetFs + " ");
                        }
                    }

                    textParam.Value = replaced;
                    appliedCount++;
                }
                catch
                {
                    // ignore and continue
                }
            }
        }

        if (appliedCount == 0)
        {
            MessageBox.Show(
                "No selected VEGAS Titles & Text events found, or text parameter not accessible.",
                "Update Font Size",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }

    private static bool IsTitlesAndText(string name)
    {
        if (string.IsNullOrEmpty(name)) return false;
        return name.IndexOf("Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0
            || name.IndexOf("VEGAS Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0
            || name.IndexOf("Sony Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static int? PromptForFontSize(int defaultSize)
    {
        using (Form dialog = new Form())
        using (Label label = new Label())
        using (NumericUpDown numeric = new NumericUpDown())
        using (Button ok = new Button())
        using (Button cancel = new Button())
        {
            dialog.Text = "Update Font Size";
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.ClientSize = new Size(280, 120);
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;

            label.AutoSize = true;
            label.Text = "Enter font size (points):";
            label.Location = new Point(12, 15);

            numeric.Minimum = 1;
            numeric.Maximum = 1000;
            numeric.Value = Math.Max(1, Math.Min(defaultSize, 1000));
            numeric.Location = new Point(15, 45);
            numeric.Width = 120;
            numeric.DecimalPlaces = 0;

            ok.Text = "OK";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(dialog.ClientSize.Width - 170, 80);
            ok.Size = new Size(75, 25);

            cancel.Text = "Cancel";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(dialog.ClientSize.Width - 90, 80);
            cancel.Size = new Size(75, 25);

            dialog.Controls.Add(label);
            dialog.Controls.Add(numeric);
            dialog.Controls.Add(ok);
            dialog.Controls.Add(cancel);

            dialog.AcceptButton = ok;
            dialog.CancelButton = cancel;

            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
                return (int)numeric.Value;
            return null;
        }
    }
}
