// AnimateCaptions.cs
// Finds all text media events, locates the first VEGAS Picture In Picture FX,
// and adds 3 Scale keyframes at the start.

using System;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public float minScale = 0.5f;
    public float maxScale = 1.5f;
    

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

                // Remove any existing Scale keyframes, then enable animation
                scale.IsAnimated = false;
                scale.IsAnimated = true;
                
                // Add pop-in keyframes at the start of the effect
                SetKey(scale, Timecode.FromFrames(0), minScale);   
                SetKey(scale, Timecode.FromFrames(popInFramesA), maxScale);   
                SetKey(scale, Timecode.FromFrames(popInFramesA + popInFramesB), 1);  

                // Add pop-out keyframes at the end of the effect
                double fps = vegas.Project.Video.FrameRate;
                long durationFrames = (long)Math.Round(ev.Length.ToMilliseconds() / 1000.0 * fps);
                if (durationFrames - popInFramesA - popInFramesB - popOutFramesA - popOutFramesB > fullPopOutMinFrameBuffer) {
                    SetKey(scale, Timecode.FromFrames(durationFrames - popOutFramesA - popOutFramesB), 1);   
                    SetKey(scale, Timecode.FromFrames(durationFrames - popOutFramesB), maxScale);  
                    SetKey(scale, Timecode.FromFrames(durationFrames), minScale);  
                } else if (durationFrames - popInFramesA - popInFramesB - popOutFramesB > halfPopOutMinFrameBuffer) {
                    SetKey(scale, Timecode.FromFrames(durationFrames - popOutFramesB), 1);   
                    SetKey(scale, Timecode.FromFrames(durationFrames), minScale);  
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
}
