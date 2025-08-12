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
        
        int frames = 0;
        double fps = vegas.Project.Video.FrameRate;
        long srcStartFrames = ConvertTimecodeToFrames(fps, mochaSourceEvent.Start);

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
        
        foreach (var target in targetPiPs)
        {
            var modeParam = target.Effect.FindParameterByName("KeepProportions") as OFXChoiceParameter;
            if (modeParam != null)
            {
                modeParam.Value = modeParam.Choices[2]; // Free Form
                modeParam.ParameterChanged();
            }
        }
        
        var cornerMap = new System.Collections.Generic.KeyValuePair<string, string>[] {
            new KeyValuePair<string,string>("surfaceTopLeft", "CornerTL"),
            new KeyValuePair<string,string>("surfaceTopRight", "CornerTR"),
            new KeyValuePair<string,string>("surfaceBottomLeft", "CornerBL"),
            new KeyValuePair<string,string>("surfaceBottomRight", "CornerBR"),
        };

        foreach (var entry in cornerMap)
        {
            var srcParam = mochaSource.FindParameterByName(entry.Key) as OFXDouble2DParameter;
            if (srcParam == null)
                throw new ApplicationException("Mocha parameter " + entry.Key + " not found or not a 2D parameter.");

            foreach (OFXKeyframe srcKf in srcParam.Keyframes)
            {
                Timecode t = srcKf.Time;
                OFXDouble2D pt = srcParam.GetValueAtTime(t);
                long srcKeyFrames = ConvertTimecodeToFrames(fps, t);

                foreach (var target in targetPiPs)
                {
                    var dstParam = target.Effect.FindParameterByName(entry.Value) as OFXDouble2DParameter;
                    if (dstParam == null) continue;

                    if (!dstParam.IsAnimated)
                        dstParam.IsAnimated = true;

                    // Normalize coordinates from pixels to [0,1]
                    var ptScaled = new OFXDouble2D();
                    ptScaled.X = pt.X * scaleX;
                    ptScaled.Y = pt.Y * scaleY;

                    long dstStartFrames = ConvertTimecodeToFrames(fps, target.Event.Start);
                    long offsetFrames = srcStartFrames - dstStartFrames;
                    long dstFrames = srcKeyFrames + offsetFrames;
                    if (dstFrames < 0)
                        continue; // skip negative local times

                    dstParam.SetValueAtTime(Timecode.FromFrames(dstFrames), ptScaled);
                    frames++;
                }
            }
        }

        // Preserve PiP counter dialog for reference
        string message;
        if (selectedCount == 0)
        {
            message = "No selected video events.";
        }
        else
        {
            message = string.Format(
                "Selected events: {0}\n- With Picture in Picture: {1}\n- Frames: {2}",
                selectedCount, selectedWithPiP, frames);
        }

        MessageBox.Show(
            message,
            "Selected FX Counts",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
    
    private static long ConvertTimecodeToFrames(double framesPerSecond, Timecode timecode)
    {
        return (long) Math.Round(timecode.ToMilliseconds() / 1000.0 * framesPerSecond);
    }
    
    private class TargetPiP
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

