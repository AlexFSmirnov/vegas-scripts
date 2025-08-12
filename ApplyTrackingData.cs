using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    // Known PiP UIDs seen across versions
    private const string PiP_UID_VEGAS = "{Svfx:com.vegascreativesoftware:pictureinpicture}";
    private const string PiP_UID_SONY  = "{Svfx:com.sonycreativesoftware:pictureinpicture}";

    // Mocha → PiP corner parameter map
    private static readonly KeyValuePair<string, string>[] CornerMap = new KeyValuePair<string, string>[] {
        new KeyValuePair<string,string>("surfaceTopLeft",     "CornerTL"),
        new KeyValuePair<string,string>("surfaceTopRight",    "CornerTR"),
        new KeyValuePair<string,string>("surfaceBottomLeft",  "CornerBL"),
        new KeyValuePair<string,string>("surfaceBottomRight", "CornerBR"),
    };

    public void FromVegas(Vegas vegas)
    {
        List<TargetPiP> targetPiPs = new List<TargetPiP>();   // last PiP + its event per selected event
        List<MochaSource> mochaSources = new List<MochaSource>(); // all Mocha sources across selected events

        int selectedCount = 0;

        ProgressForm progress = null;
        int totalSteps = 0;

        try
        {
            // Collect selected events, their last PiP (if any), and Mocha sources
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

                    Effect lastPip = FindLastPiP(videoEvent, PiP_UID_VEGAS, PiP_UID_SONY);
                    if (lastPip != null && lastPip.OFXEffect != null)
                    {
                        targetPiPs.Add(new TargetPiP { Effect = lastPip.OFXEffect, Event = videoEvent });
                    }

                    OFXEffect mocha = FindMochaOFX(videoEvent);
                    if (mocha != null)
                    {
                        mochaSources.Add(new MochaSource { Effect = mocha, Event = videoEvent });
                    }
                }
            }

            if (mochaSources.Count == 0)
                throw new ApplicationException("No Mocha source found among the selected events.");

            if (targetPiPs.Count == 0)
                throw new ApplicationException("No Picture in Picture effects found on the selected events.");

            // Ensure PiP is in Free Form mode
            EnsurePiPFreeForm(targetPiPs);

            double fps = vegas.Project.Video.FrameRate;

            // Cache start/end frames per target for fast in-range checks
            for (int i = 0; i < targetPiPs.Count; i++)
            {
                TargetPiP t = targetPiPs[i];
                t.StartFrames = ConvertTimecodeToFrames(fps, t.Event.Start);
                t.EndFrames   = ConvertTimecodeToFrames(fps, t.Event.Start + t.Event.Length);
                targetPiPs[i] = t;
            }

            // Pre-calc approximate total steps across ALL mocha sources (for progress)
            totalSteps = EstimateTotalSteps(mochaSources, targetPiPs, CornerMap);
            if (totalSteps <= 0) totalSteps = 1;

            // Show the progress window
            progress = new ProgressForm();
            progress.SetStatus("Preparing…");
            progress.SetMax(totalSteps);
            progress.Show();
            Application.DoEvents();

            // Copy keyframes from every Mocha source to all target PiPs
            int framesCopied = CopyMochaToPiP(mochaSources, targetPiPs, fps, progress);

            progress.Complete("Done");
            Application.DoEvents();

            // Summary
            string message;
            if (selectedCount == 0)
            {
                message = "No selected video events.";
            }
            else
            {
                message = string.Format(
                    "PiP targets: {0}\nMocha sources: {1}\nFrames copied: {2}",
                    targetPiPs.Count, mochaSources.Count, framesCopied);
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

    // === Core operations ===

    private static void EnsurePiPFreeForm(List<TargetPiP> targetPiPs)
    {
        for (int i = 0; i < targetPiPs.Count; i++)
        {
            OFXEffect fx = targetPiPs[i].Effect;
            OFXChoiceParameter modeParam = fx.FindParameterByName("KeepProportions") as OFXChoiceParameter;
            if (modeParam == null) continue;

            // Choices: 0=On, 1=Fill, 2=Free Form (empirical)
            if (modeParam.Choices != null && modeParam.Choices.Length >= 3)
            {
                modeParam.Value = modeParam.Choices[2]; // Free Form
                modeParam.ParameterChanged();
            }
        }
    }

    private static int EstimateTotalSteps(List<MochaSource> mochaSources, List<TargetPiP> targetPiPs, KeyValuePair<string, string>[] cornerMap)
    {
        int total = 0;
        for (int i = 0; i < mochaSources.Count; i++)
        {
            OFXEffect srcFx = mochaSources[i].Effect;
            for (int j = 0; j < cornerMap.Length; j++)
            {
                OFXDouble2DParameter p = srcFx.FindParameterByName(cornerMap[j].Key) as OFXDouble2DParameter;
                if (p != null && p.Keyframes != null)
                {
                    total += (p.Keyframes.Count * Math.Max(1, targetPiPs.Count));
                }
            }
        }
        return total;
    }

    private static int CopyMochaToPiP(List<MochaSource> mochaSources, List<TargetPiP> targetPiPs, double fps, ProgressForm progress)
    {
        int framesCopied = 0;

        int sourceIndex = 0;
        foreach (MochaSource ms in mochaSources)
        {
            sourceIndex++;

            // Determine normalization scales from the first Top Right entry (width, height in pixels) PER SOURCE
            double scaleX = 1.0;
            double scaleY = 1.0;
            try
            {
                OFXDouble2DParameter topRightParam = ms.Effect.FindParameterByName("surfaceTopRight") as OFXDouble2DParameter;
                if (topRightParam != null)
                {
                    OFXDouble2D wh;
                    if (topRightParam.Keyframes != null && topRightParam.Keyframes.Count > 0)
                    {
                        // Get value at the first keyframe time
                        Timecode firstTime = topRightParam.Keyframes[0].Time;
                        wh = topRightParam.GetValueAtTime(firstTime);
                    }
                    else
                    {
                        // Fallback to current value
                        wh = topRightParam.Value;
                    }

                    double widthPixels = wh.X;
                    double heightPixels = wh.Y;

                    if (widthPixels != 0.0) scaleX = 1.0 / widthPixels;
                    if (heightPixels != 0.0) scaleY = 1.0 / heightPixels;
                }
            }
            catch { }

            long srcStartFrames = ConvertTimecodeToFrames(fps, ms.Event.Start);

            int targetIndex = 0;
            foreach (TargetPiP target in targetPiPs)
            {
                targetIndex++;
                progress.SetStatus("Processing source " + sourceIndex + "/" + mochaSources.Count + " • event " + targetIndex + "/" + targetPiPs.Count + "…");

                for (int k = 0; k < CornerMap.Length; k++)
                {
                    KeyValuePair<string, string> map = CornerMap[k];

                    OFXDouble2DParameter srcParam = ms.Effect.FindParameterByName(map.Key) as OFXDouble2DParameter;
                    if (srcParam == null)
                        throw new ApplicationException("Mocha parameter " + map.Key + " not found or not a 2D parameter.");

                    OFXDouble2DParameter dstParam = target.Effect.FindParameterByName(map.Value) as OFXDouble2DParameter;
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
                            OFXDouble2D ptScaled = new OFXDouble2D();
                            ptScaled.X = pt.X * scaleX;
                            ptScaled.Y = pt.Y * scaleY;

                            long dstKeyFrames = srcKeyFramesAbs - target.StartFrames;
                            dstParam.SetValueAtTime(Timecode.FromFrames(dstKeyFrames), ptScaled);
                            framesCopied++;
                        }

                        // Advance progress for each examined keyframe (applies whether copied or skipped)
                        progress.Increment();
                    }
                }
            }
        }

        return framesCopied;
    }

    // === Helpers & data containers ===

    private static long ConvertTimecodeToFrames(double framesPerSecond, Timecode timecode)
    {
        return (long)Math.Round(timecode.ToMilliseconds() / 1000.0 * framesPerSecond);
    }

    private class TargetPiP
    {
        public OFXEffect Effect;
        public VideoEvent Event;
        public long StartFrames;
        public long EndFrames;
    }

    private class MochaSource
    {
        public OFXEffect Effect;
        public VideoEvent Event;
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
