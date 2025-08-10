using System;
using System.Drawing;
using System.Windows.Forms; // RichTextBox for RTF <-> plain text
using ScriptPortal.Vegas;

public class EntryPoint
{
    // Number of characters that fill the full width at Scale = 1
    private const double FullWidthCharacters = 42.0;

    // Toggle to append/update a debug line with the measured longest line length
    private const bool DebugMode = true;

    public void FromVegas(Vegas vegas)
    {
        const string PiP_UID = "{Svfx:com.vegascreativesoftware:pictureinpicture}";

        foreach (Track track in vegas.Project.Tracks)
        {
            VideoTrack videoTrack = track as VideoTrack;
            if (videoTrack == null)
                continue;

            foreach (TrackEvent trackEvent in videoTrack.Events)
            {
                VideoEvent videoEvent = trackEvent as VideoEvent;
                if (videoEvent == null)
                    continue;

                Take activeTake = videoEvent.ActiveTake;
                if (activeTake == null || activeTake.Media == null || !activeTake.Media.IsGenerated() || activeTake.Media.Generator == null)
                    continue;

                string generatorName = (activeTake.Media.Generator.PlugIn != null) ? activeTake.Media.Generator.PlugIn.Name : "";
                if (!IsTextGenerator(generatorName))
                    continue;

                // Read the Titles & Text "Text" parameter (RTF) and convert to plain text for measuring
                string rtf = GetGeneratedTextRtf(activeTake);
                if (string.IsNullOrEmpty(rtf))
                    continue;

                string textContent = RtfToPlainText(rtf);
                if (string.IsNullOrEmpty(textContent))
                    continue;

                int longestLineLength = GetLongestLineLength(textContent);
                if (longestLineLength <= 0)
                    continue;

                double targetScale = FullWidthCharacters / (double)longestLineLength;

                // --- Debug: annotate Titles & Text content by updating its RTF "Text" param ---
                if (DebugMode)
                {
                    try
                    {
                        string updatedRtf = UpsertDebugLineIntoRtf(rtf, longestLineLength);
                        SetGeneratedTextRtf(activeTake, updatedRtf);
                        // refresh local copies
                        rtf = updatedRtf;
                        textContent = RtfToPlainText(rtf);
                    }
                    catch
                    {
                        // best-effort debug annotation
                    }
                }
                // ---------------------------------------------------------------------------

                Effect pip = FindFirstPiP(videoEvent, PiP_UID);
                if (pip == null)
                    continue;

                OFXEffect ofx = pip.OFXEffect;
                if (ofx == null)
                    continue;

                OFXDoubleParameter scaleParam = ofx.FindParameterByName("Scale") as OFXDoubleParameter;
                if (scaleParam == null)
                    continue;

                // Set a single static value at the start of the event
                scaleParam.IsAnimated = false;
                scaleParam.SetValueAtTime(Timecode.FromFrames(0), targetScale);
            }
        }
    }

    // ----- Helpers -----

    private static bool IsTextGenerator(string name)
    {
        if (name == null) return false;
        return name.IndexOf("Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0
            || name.IndexOf("VEGAS Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0
            || name.IndexOf("Text", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static Effect FindFirstPiP(VideoEvent ev, string uid)
    {
        foreach (Effect fx in ev.Effects)
        {
            if (!fx.IsOFX) continue;
            if (fx.PlugIn == null) continue;

            bool uidMatch = (fx.PlugIn.UniqueID == uid);
            bool nameMatch = fx.PlugIn.Name != null &&
                             fx.PlugIn.Name.IndexOf("Picture in Picture", StringComparison.OrdinalIgnoreCase) >= 0;

            if (uidMatch || nameMatch)
                return fx;
        }
        return null;
    }

    private static int GetLongestLineLength(string text)
    {
        string normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
        string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);
        int longest = 0;
        foreach (string line in lines)
        {
            if (line.Length > longest)
                longest = line.Length;
        }
        return longest;
    }

    // ----- Titles & Text RTF helpers -----

    private static string GetGeneratedTextRtf(Take take)
    {
        try
        {
            if (take == null || take.Media == null || take.Media.Generator == null)
                return null;

            OFXEffect genOfx = take.Media.Generator.OFXEffect;
            if (genOfx != null)
            {
                // The Titles & Text "Text" parameter stores RTF, not plain text
                OFXStringParameter textParam = genOfx.FindParameterByName("Text") as OFXStringParameter;
                if (textParam != null)
                    return textParam.Value;
            }
        }
        catch { }
        return null;
    }

    private static void SetGeneratedTextRtf(Take take, string rtf)
    {
        if (take == null || take.Media == null || take.Media.Generator == null) return;
        OFXEffect genOfx = take.Media.Generator.OFXEffect;
        if (genOfx == null) return;

        OFXStringParameter textParam = genOfx.FindParameterByName("Text") as OFXStringParameter;
        if (textParam == null) return;

        textParam.Value = rtf;
    }

    private static string RtfToPlainText(string rtf)
    {
        if (string.IsNullOrEmpty(rtf)) return "";
        using (RichTextBox rbx = new RichTextBox())
        {
            try
            {
                rbx.Rtf = rtf;
                return rbx.Text ?? "";
            }
            catch
            {
                // If it wasn't valid RTF, treat it as plain
                return rtf;
            }
        }
    }

    // Appends or updates a trailing debug line like "[DEBUG LEN=NN]" into the RTF
    // Preserves overall font/alignment by snapshotting from the existing content (per VEGAS FAQ approach)
    private static string UpsertDebugLineIntoRtf(string existingRtf, int length)
    {
        using (RichTextBox rbx = new RichTextBox())
        {
            // Load existing content
            try { rbx.Rtf = existingRtf; }
            catch { rbx.Text = existingRtf ?? ""; }

            // Capture overall formatting
            rbx.SelectAll();
            Font savedFont = rbx.SelectionFont ?? rbx.Font;
            HorizontalAlignment savedAlignment = rbx.SelectionAlignment;

            // Work in plain text to manage the final line
            string plain = rbx.Text ?? "";
            string normalized = plain.Replace("\r\n", "\n").Replace("\r", "\n");
            string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);

            string prefix = "[DEBUG LEN=";
            string debugLine = prefix + length + "]";

            if (lines.Length == 0 || string.IsNullOrEmpty(normalized))
            {
                normalized = debugLine;
            }
            else if (lines[lines.Length - 1].StartsWith(prefix, StringComparison.Ordinal))
            {
                lines[lines.Length - 1] = debugLine;
                normalized = string.Join("\n", lines);
            }
            else
            {
                normalized += "\n" + debugLine;
            }

            // Replace content and restore captured formatting
            rbx.Text = normalized;
            rbx.SelectAll();
            rbx.SelectionFont = savedFont;
            rbx.SelectionAlignment = savedAlignment;

            return rbx.Rtf;
        }
    }
}
