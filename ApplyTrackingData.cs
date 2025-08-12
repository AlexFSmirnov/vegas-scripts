using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        // Known PiP UIDs seen across versions
        const string PiP_UID_VEGAS = "{Svfx:com.vegascreativesoftware:pictureinpicture}";
        const string PiP_UID_SONY  = "{Svfx:com.sonycreativesoftware:pictureinpicture}";

        OFXEffect mochaSource = null; // single Mocha source across all selected events
        VideoEvent mochaSourceEvent = null; // the event that hosts the mocha source
        List<TargetPiP> targetPiPs = new List<TargetPiP>(); // last PiP + its event per selected event

        int selectedCount = 0;
        int selectedWithPiP = 0;

        ProgressForm progress = null;
        int totalSteps = 0;
        int stepCounter = 0;

        try
        {
            foreach (Track track in vegas.Project.Tracks)
            {
                VideoTrack videoTrack = track as VideoTrack;
                if (videoTrack == null) continue;

                foreach (TrackEvent trackEvent in videoTrack.Events)
                {
                    if (!trackEvent.Selected) continue;
                    selectedCount++;

                    VideoEvent videoEvent = trackEvent as VideoEvent;
                    if (videoEvent == null) continue;

                    // Collect last PiP OFXEffect for this selected event (if any)
                    Effect lastPip = FindLastPiP(videoEvent, PiP_UID_VEGAS, PiP_UID_SONY);
                    if (lastPip != null && lastPip.OFXEffect != null)
                    {
                        targetPiPs.Add(new TargetPiP { Effect = lastPip.OFXEffect, Event = videoEvent });
                        selectedWithPiP++;
                    }

                    // Determine single Mocha OFXEffect source (throw if more than one selected event has Mocha)
                    OFXEffect mocha = FindMochaOFX(videoEvent);
                    if (mocha != null)
                    {
                        if (mochaSource == null)
                        {
                            mochaSource = mocha;
                            mochaSourceEvent = videoEvent;
                        }
                        else
                        {
                            throw new ApplicationException("Multiple selected events with Mocha detected. Select exactly one Mocha source event.");
                        }
                    }
                }
            }

            // === Copy Mocha corner pin keyframes -> all selected PiP effects ===
            if (mochaSource == null)
                throw new ApplicationException("No Mocha source found among the selected events.");

            if (targetPiPs.Count == 0)
                throw new ApplicationException("No Picture in Picture effects found on the selected events.");

            // Determine normalization scales from the first Top Right entry (width, height in pixels)
            double scaleX = 1.0;
            double scaleY = 1.0;
            try
            {
                var topRightParam = mochaSource.FindParameterByName("surfaceTopRight") as OFXDouble2DParameter;
                if (topRightParam != null)
                {
                    OFXDouble2D wh;
                    var keys = topRightParam.Keyframes;
                    if (keys != null && keys.Count > 0)
                    {
                        // Get value at the first keyframe time
                        var firstTime = keys[0].Time;
                        wh = topRightParam.GetValueAtTime(firstTime);
                    }
                    else
                    {
                        // Fallback to current value
                        wh = topRightParam.Value;
                    }

                    double widthPixels = wh.X;
                    double heightPixels = wh.Y;

                    if (widthPixels != 0.0)
                        scaleX = 1.0 / widthPixels;
                    if (heightPixels != 0.0)
                        scaleY = 1.0 / heightPixels;
                }
            }
            catch { }

            // Ensure PiP is in Free Form mode once
            foreach (var target in targetPiPs)
            {
                var modeParam = target.Effect.FindParameterByName("KeepProportions") as OFXChoiceParameter;
                if (modeParam != null)
                {
                    // Choices: 0=On, 1=Fill, 2=Free Form (empirical)
                    if (modeParam.Choices != null && modeParam.Choices.Length >= 3)
                    {
                        modeParam.Value = modeParam.Choices[2]; // Free Form
                        modeParam.ParameterChanged();
                    }
                }
            }

            var cornerMap = new System.Collections.Generic.KeyValuePair<string, string>[] {
                new KeyValuePair<string,string>("surfaceTopLeft", "CornerTL"),
                new KeyValuePair<string,string>("surfaceTopRight", "CornerTR"),
                new KeyValuePair<string,string>("surfaceBottomLeft", "CornerBL"),
                new KeyValuePair<string,string>("surfaceBottomRight", "CornerBR"),
            };

            // Pre-calc an approximate total step count for the progress bar
            totalSteps = 0;
            foreach (var entry in cornerMap)
            {
                var p = mochaSource.FindParameterByName(entry.Key) as OFXDouble2DParameter;
                if (p != null && p.Keyframes != null)
                {
                    // Each keyframe potentially copied to each target (filtered by time window later)
                    totalSteps += (p.Keyframes.Count * Math.Max(1, targetPiPs.Count));
                }
            }
            if (totalSteps <= 0) totalSteps = 1;

            // Show the progress window
            progress = new ProgressForm();
            progress.SetStatus("Preparing…");
            progress.SetMax(totalSteps);
            progress.Show();
            Application.DoEvents();

            double fps = vegas.Project.Video.FrameRate;
            long srcStartFrames = ConvertTimecodeToFrames(fps, mochaSourceEvent.Start);

            // Cache start/end frames per target for fast in-range checks
            foreach (var t in targetPiPs)
            {
                t.StartFrames = ConvertTimecodeToFrames(fps, t.Event.Start);
                t.EndFrames   = ConvertTimecodeToFrames(fps, t.Event.Start + t.Event.Length);
            }

            // Copy only the keyframes that fall inside each target event's active time window.
            int frames = 0;
            int targetIndex = 0;

            foreach (var target in targetPiPs)
            {
                targetIndex++;
                progress.SetStatus("Processing event " + targetIndex + "/" + targetPiPs.Count + "…");

                foreach (var entry in cornerMap)
                {
                    var srcParam = mochaSource.FindParameterByName(entry.Key) as OFXDouble2DParameter;
                    if (srcParam == null)
                        throw new ApplicationException("Mocha parameter " + entry.Key + " not found or not a 2D parameter.");

                    var dstParam = target.Effect.FindParameterByName(entry.Value) as OFXDouble2DParameter;
                    if (dstParam == null) 
                    {
                        // Still advance progress for src keyframes to keep bar moving consistently
                        if (srcParam.Keyframes != null)
                        {
                            foreach (OFXKeyframe _ in srcParam.Keyframes)
                                progress.Increment();
                        }
                        continue;
                    }

                    if (!dstParam.IsAnimated)
                        dstParam.IsAnimated = true;

                    foreach (OFXKeyframe srcKf in srcParam.Keyframes)
                    {
                        Timecode t = srcKf.Time;
                        OFXDouble2D pt = srcParam.GetValueAtTime(t);
                        long srcKeyFrames = ConvertTimecodeToFrames(fps, t);
                        long srcKeyFramesAbs = srcKeyFrames + srcStartFrames;

                        // Only write if the target event actually exists at this absolute time
                        if (srcKeyFramesAbs >= target.StartFrames && srcKeyFramesAbs <= target.EndFrames)
                        {
                            // Normalize coordinates from pixels to [0,1]
                            var ptScaled = new OFXDouble2D
                            {
                                X = pt.X * scaleX,
                                Y = pt.Y * scaleY
                            };

                            long dstKeyFrames = srcKeyFramesAbs - target.StartFrames;
                            dstParam.SetValueAtTime(Timecode.FromFrames(dstKeyFrames), ptScaled);
                            frames++;
                        }

                        // Advance progress for each examined keyframe (applies whether copied or skipped)
                        stepCounter++;
                        progress.Increment();
                    }
                }
            }

            progress.Complete("Done");
            Application.DoEvents();

            // Preserve PiP counter dialog for reference
            string message;
            if (selectedCount == 0)
            {
                message = "No selected video events.";
            }
            else
            {
                message = string.Format(
                    "Selected events: {0}\n- With Picture in Picture: {1}\n- Frames copied: {2}",
                    selectedCount, selectedWithPiP, frames);
            }

            MessageBox.Show(
                message,
                "Selected FX Counts",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (progress != null)
            {
                try { progress.SetStatus("Error"); progress.Complete("Error"); } catch { }
            }
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (progress != null)
            {
                try { progress.Close(); progress.Dispose(); } catch { }
            }
        }
    }

    private static long ConvertTimecodeToFrames(double framesPerSecond, Timecode timecode)
    {
        return (long) Math.Round(timecode.ToMilliseconds() / 1000.0 * framesPerSecond);
    }

    private class TargetPiP
    {
        public OFXEffect Effect;
        public VideoEvent Event;
        public long StartFrames;
        public long EndFrames;
    }

    private static Effect FindLastPiP(VideoEvent videoEvent, params string[] knownUids)
    {
        Effect last = null;
        foreach (Effect effect in videoEvent.Effects)
        {
            if (IsPiPEffect(effect, knownUids))
            {
                last = effect;
            }
        }
        return last;
    }

    private static bool IsPiPEffect(Effect effect, params string[] knownUids)
    {
        if (effect == null || !effect.IsOFX || effect.PlugIn == null) return false;

        string uid = effect.PlugIn.UniqueID ?? string.Empty;
        for (int i = 0; i < knownUids.Length; i++)
        {
            if (string.Equals(uid, knownUids[i], StringComparison.OrdinalIgnoreCase))
                return true;
        }

        string name = effect.PlugIn.Name ?? string.Empty;
        return name.IndexOf("Picture in Picture", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static OFXEffect FindMochaOFX(VideoEvent videoEvent)
    {
        foreach (Effect effect in videoEvent.Effects)
        {
            if (!effect.IsOFX || effect.PlugIn == null) continue;

            string uid = effect.PlugIn.UniqueID ?? string.Empty;
            string name = effect.PlugIn.Name ?? string.Empty;

            if (uid.IndexOf("mocha", StringComparison.OrdinalIgnoreCase) >= 0)
                return effect.OFXEffect;
            if (name.IndexOf("Mocha", StringComparison.OrdinalIgnoreCase) >= 0)
                return effect.OFXEffect;
        }
        return null;
    }
}

/// <summary>
/// Tiny modal-free progress window that stays responsive during long ops.
/// </summary>
internal sealed class ProgressForm : Form
{
    private readonly ProgressBar _bar = new ProgressBar { Dock = DockStyle.Top, Style = ProgressBarStyle.Continuous, Height = 22 };
    private readonly Label _label = new Label { Dock = DockStyle.Top, AutoSize = false, Height = 22, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };

    public ProgressForm()
    {
        Text = "Copying Mocha → PiP";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Width = 420;
        Height = 110;

        Padding = new Padding(10);
        _label.Margin = new Padding(0, 0, 0, 6);

        Controls.Add(_label);
        Controls.Add(_bar);
    }

    public void SetMax(int max)
    {
        if (max <= 0) max = 1;
        _bar.Minimum = 0;
        _bar.Maximum = max;
        _bar.Value = 0;
        Pump();
    }

    public void Increment()
    {
        if (_bar.Value < _bar.Maximum) _bar.Value++;
        Pump();
    }

    public void SetStatus(string text)
    {
        _label.Text = text;
        Pump();
    }

    public void Complete(string text = "Done")
    {
        _label.Text = text;
        _bar.Value = _bar.Maximum;
        Pump();
    }

    private static void Pump()
    {
        try { Application.DoEvents(); } catch { /* ignore */ }
    }
}
