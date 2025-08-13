using System;
using ScriptPortal.Vegas;

public class EntryPoint
{
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

                // Find last Picture in Picture effect with Fixed Shape mode on this event
                Effect pip = FindLastPiPWithFixedShape(ev, PiP_UID);
                if (pip == null)
                    continue;

                // Get OFX and a parameter to keyframe (Scale is simple & safe)
                OFXEffect ofx = pip.OFXEffect;
                if (ofx == null)
                    continue;

                OFXDoubleParameter scale = ofx.FindParameterByName("Scale") as OFXDoubleParameter;
                if (scale == null)
                    continue;

                scale.IsAnimated = false;
                scale.SetValueAtTime(Timecode.FromFrames(0), 1);
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

    private static Effect FindLastPiPWithFixedShape(VideoEvent ev, string uid)
    {
        Effect lastFixedShapePiP = null;
        
        foreach (Effect fx in ev.Effects)
        {
            if (!fx.IsOFX) continue;
            if (fx.PlugIn == null) continue;

            bool uidMatch = (fx.PlugIn.UniqueID == uid);
            bool nameMatch = fx.PlugIn.Name != null &&
                             fx.PlugIn.Name.IndexOf("Picture in Picture", StringComparison.OrdinalIgnoreCase) >= 0;

            if (uidMatch || nameMatch)
            {
                // Check if this PiP has Fixed Shape mode (choice index 1)
                if (IsPiPInFixedShapeMode(fx))
                {
                    lastFixedShapePiP = fx;
                }
            }
        }
        return lastFixedShapePiP;
    }

    private static bool IsPiPInFixedShapeMode(Effect fx)
    {
        if (fx == null || !fx.IsOFX || fx.OFXEffect == null)
            return false;

        OFXChoiceParameter modeParam = fx.OFXEffect.FindParameterByName("KeepProportions") as OFXChoiceParameter;
        if (modeParam == null || modeParam.Choices == null)
            return false;

        // Choice index 0 should be "Fixed Shape" mode
        return modeParam.Value != null && 
               modeParam.Choices.Length > 1 && 
               modeParam.Value.Index == 0;
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
}
