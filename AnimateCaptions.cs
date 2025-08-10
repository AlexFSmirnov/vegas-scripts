using System;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    // Width estimation borrowed from ResizeToFullWidth.cs
    private const double FullWidthCharacters = 42.0; // characters at Scale=1 that fill full width
    private const double DefaultMaxScale = 1.2;      // 120%
    private const double DefaultMinScale = 0.8;      // 50%
    private const double MaxAbsExpansionFraction = 0.04; // max +10% of full width in absolute terms
    

    public long popInFramesA = 4;
    public long popInFramesB = 6;
    
    public long popOutFramesA = 4;
    public long popOutFramesB = 3;
    

    public long fullPopOutMinFrameBuffer = 15;
    public long halfPopOutMinFrameBuffer = 8;


    public void FromVegas(Vegas vegas)
    {
        // Known UID for VEGAS Picture In Picture (fallback to name check as well)
        const string PiP_UID = "{Svfx:com.vegascreativesoftware:pictureinpicture}";

        foreach (Track t in vegas.Project.Tracks)
        {
            // Only video tracks
            VideoTrack vtrack = t as VideoTrack;
            if (vtrack == null)
                continue;

            foreach (TrackEvent te in vtrack.Events)
            {
                VideoEvent ev = te as VideoEvent;
                if (ev == null)
                    continue;

                // Only generated text media (Titles & Text etc.)
                Take take = ev.ActiveTake;
                if (take == null || take.Media == null || !take.Media.IsGenerated() || take.Media.Generator == null)
                    continue;

                string genName = (take.Media.Generator.PlugIn != null) ? take.Media.Generator.PlugIn.Name : "";
                if (!IsTextGenerator(genName))
                    continue;

                // Find first Picture in Picture effect on this event
                Effect pip = FindFirstPiP(ev, PiP_UID);
                if (pip == null)
                    continue;

                // Get OFX and a parameter to keyframe (Scale is simple & safe)
                OFXEffect ofx = pip.OFXEffect;
                if (ofx == null)
                    continue;

                OFXDoubleParameter scale = ofx.FindParameterByName("Scale") as OFXDoubleParameter;
                if (scale == null)
                    continue;

                // Compute per-event scales
                double minScaleValue = ComputeMinScaleForEvent(take);
                double maxScaleValue = ComputeMaxScaleForEvent(take);

                // Remove any existing Scale keyframes, then enable animation
                scale.IsAnimated = false;
                scale.IsAnimated = true;
                
                // Add pop-in keyframes at the start of the effect
                SetKey(scale, Timecode.FromFrames(0), minScaleValue);   
                SetKey(scale, Timecode.FromFrames(popInFramesA), maxScaleValue);   
                SetKey(scale, Timecode.FromFrames(popInFramesA + popInFramesB), 1);  

                // Add pop-out keyframes at the end of the effect unless another text event follows immediately
                // or the event ends together with an audio event (clip change on audio)
                double fps = vegas.Project.Video.FrameRate;
                bool skipPopOut = HasImmediatelyFollowingTextEvent(vtrack, ev, fps) || EndsTogetherWithAudioEvent(vegas, ev, fps);
                if (!skipPopOut)
                {
                    long durationFrames = (long)Math.Round(ev.Length.ToMilliseconds() / 1000.0 * fps);
                    if (durationFrames - popInFramesA - popInFramesB - popOutFramesA - popOutFramesB > fullPopOutMinFrameBuffer) {
                        SetKey(scale, Timecode.FromFrames(durationFrames - popOutFramesA - popOutFramesB), 1);   
                        SetKey(scale, Timecode.FromFrames(durationFrames - popOutFramesB), maxScaleValue);  
                        SetKey(scale, Timecode.FromFrames(durationFrames), minScaleValue);  
                    } else if (durationFrames - popInFramesA - popInFramesB - popOutFramesB > halfPopOutMinFrameBuffer) {
                        SetKey(scale, Timecode.FromFrames(durationFrames - popOutFramesB), 1);   
                        SetKey(scale, Timecode.FromFrames(durationFrames), minScaleValue);  
                    }
                }
            }
        }
    }

    // ----- Helpers -----

    private static bool IsTextGenerator(string name)
    {
        if (name == null) return false;
        // Case-insensitive checks without modern overloads
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

    private static void SetKey(OFXDoubleParameter p, Timecode tc, double value)
    {
        // VEGAS will create or update a keyframe at 'tc' automatically.
        p.IsAnimated = true;
        p.SetValueAtTime(tc, value);
    }

    private static long ConvertTimecodeToFrames(double framesPerSecond, Timecode timecode)
    {
        return (long) Math.Round(timecode.ToMilliseconds() / 1000.0 * framesPerSecond);
    }

    private static bool HasImmediatelyFollowingTextEvent(VideoTrack track, VideoEvent currentEvent, double framesPerSecond)
    {
        long currentEndFrames = ConvertTimecodeToFrames(framesPerSecond, currentEvent.Start)
                                + ConvertTimecodeToFrames(framesPerSecond, currentEvent.Length);

        foreach (TrackEvent teNext in track.Events)
        {
            VideoEvent nextEvent = teNext as VideoEvent;
            if (nextEvent == null || nextEvent == currentEvent) continue;

            // Only consider generated text media events
            Take nextTake = nextEvent.ActiveTake;
            if (nextTake == null || nextTake.Media == null || !nextTake.Media.IsGenerated() || nextTake.Media.Generator == null)
                continue;

            string nextGenName = (nextTake.Media.Generator.PlugIn != null) ? nextTake.Media.Generator.PlugIn.Name : "";
            if (!IsTextGenerator(nextGenName))
                continue;

            long nextStartFrames = ConvertTimecodeToFrames(framesPerSecond, nextEvent.Start);
            if (nextStartFrames == currentEndFrames)
                return true;
        }

        return false;
    }

    private static bool EndsTogetherWithAudioEvent(Vegas vegas, VideoEvent currentEvent, double framesPerSecond)
    {
        long currentEndFrames = ConvertTimecodeToFrames(framesPerSecond, currentEvent.Start)
                                + ConvertTimecodeToFrames(framesPerSecond, currentEvent.Length);

        foreach (Track track in vegas.Project.Tracks)
        {
            AudioTrack audioTrack = track as AudioTrack;
            if (audioTrack == null) continue;

            foreach (TrackEvent te in audioTrack.Events)
            {
                AudioEvent otherEvent = te as AudioEvent;
                if (otherEvent == null) continue;

                long otherEndFrames = ConvertTimecodeToFrames(framesPerSecond, otherEvent.Start)
                                      + ConvertTimecodeToFrames(framesPerSecond, otherEvent.Length);
                if (otherEndFrames != currentEndFrames)
                    continue;

                // Any audio event ending together counts
                return true;
            }
        }

        return false;
    }

    // ----- Text width helpers (borrowed from ResizeToFullWidth.cs) -----

    private static string GetGeneratedTextRtf(Take take)
    {
        try
        {
            if (take == null || take.Media == null || take.Media.Generator == null)
                return null;

            OFXEffect genOfx = take.Media.Generator.OFXEffect;
            if (genOfx != null)
            {
                OFXStringParameter textParam = genOfx.FindParameterByName("Text") as OFXStringParameter;
                if (textParam != null)
                    return textParam.Value;
            }
        }
        catch { }
        return null;
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
                return rtf;
            }
        }
    }

    private static int GetLongestLineLength(string text)
    {
        string normalized = (text ?? "").Replace("\r\n", "\n").Replace("\r", "\n");
        string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);
        int longest = 0;
        foreach (string line in lines)
        {
            if (line.Length > longest)
                longest = line.Length;
        }
        return longest;
    }

    private static double ComputeMaxScaleForEvent(Take take)
    {
        // Default behavior if we cannot estimate width
        double fallback = DefaultMaxScale;

        string rtf = GetGeneratedTextRtf(take);
        if (string.IsNullOrEmpty(rtf))
            return fallback;

        string plain = RtfToPlainText(rtf);
        int longest = GetLongestLineLength(plain);
        if (longest <= 0)
            return fallback;

        double baseFractionOfFullWidth = longest / FullWidthCharacters; // width at Scale=1 as fraction of full width
        if (baseFractionOfFullWidth <= 0)
            return fallback;

        // Desired width after max scale (fraction of full width)
        double desiredWidth = baseFractionOfFullWidth * DefaultMaxScale;
        double maxAllowedWidth = baseFractionOfFullWidth + MaxAbsExpansionFraction; // no more than +10% of full width
        double clampedWidth = Math.Min(desiredWidth, maxAllowedWidth);

        double computedScale = clampedWidth / baseFractionOfFullWidth;
        if (computedScale < 1.0)
            return 1.0; // do not shrink for max
        return computedScale;
    }

    private static double ComputeMinScaleForEvent(Take take)
    {
        // Default behavior if we cannot estimate width
        double fallback = DefaultMinScale;

        string rtf = GetGeneratedTextRtf(take);
        if (string.IsNullOrEmpty(rtf))
            return fallback;

        string plain = RtfToPlainText(rtf);
        int longest = GetLongestLineLength(plain);
        if (longest <= 0)
            return fallback;

        double baseFractionOfFullWidth = longest / FullWidthCharacters; // width at Scale=1 as fraction of full width
        if (baseFractionOfFullWidth <= 0)
            return fallback;

        // Desired width after min scale (fraction of full width)
        double desiredWidth = baseFractionOfFullWidth * DefaultMinScale;
        double minAllowedWidth = Math.Max(baseFractionOfFullWidth - MaxAbsExpansionFraction, 0.0); // no more than -10% of full width
        double clampedWidth = Math.Max(desiredWidth, minAllowedWidth);

        double computedScale = clampedWidth / baseFractionOfFullWidth;
        if (computedScale > 1.0)
            return 1.0; // do not grow for min
        return computedScale;
    }
}
