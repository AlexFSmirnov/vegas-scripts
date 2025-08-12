using System;
using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        // Known PiP UIDs seen across versions
        const string PiP_UID_VEGAS = "{Svfx:com.vegascreativesoftware:pictureinpicture}";
        const string PiP_UID_SONY  = "{Svfx:com.sonycreativesoftware:pictureinpicture}";

        // Collect all selected text video events that have a PiP with accessible Location
        var selectedWithPiP = new System.Collections.Generic.List<EventWithPiPLocation>();

        foreach (Track track in vegas.Project.Tracks)
        {
            var videoTrack = track as VideoTrack;
            if (videoTrack == null) continue;

            foreach (TrackEvent te in videoTrack.Events)
            {
                if (!te.Selected) continue;

                var videoEvent = te as VideoEvent;
                if (videoEvent == null) continue;

                if (videoEvent.Takes.Count == 0) continue;
                var take = videoEvent.ActiveTake;
                if (take == null || !take.IsGenerator) continue;

                if (!IsTextGenerator(take)) continue;

                Effect pip = FindPiP(videoEvent, PiP_UID_VEGAS, PiP_UID_SONY);
                if (pip == null || !pip.IsOFX || pip.OFXEffect == null) continue;

                OFXDouble2DParameter loc2D;
                OFXDoubleParameter xParam, yParam;
                if (!TryGetLocation(pip.OFXEffect, out loc2D, out xParam, out yParam))
                    continue;

                selectedWithPiP.Add(new EventWithPiPLocation
                {
                    Event = videoEvent,
                    Location2D = loc2D,
                    X = xParam,
                    Y = yParam
                });
            }
        }

        if (selectedWithPiP.Count == 0)
        {
            MessageBox.Show(
                "No selected text events with Picture in Picture 'Location' found.",
                "Copy PiP Location",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        // Find earliest by start time
        EventWithPiPLocation earliest = null;
        foreach (var item in selectedWithPiP)
        {
            if (earliest == null || item.Event.Start < earliest.Event.Start)
                earliest = item;
        }
        if (earliest == null) return;

        // Read reference location as doubles
        double refX, refY;
        if (earliest.Location2D != null)
        {
            var v = earliest.Location2D.Value;
            refX = v.X;
            refY = v.Y;
        }
        else
        {
            refX = (earliest.X != null) ? earliest.X.Value : 0.0;
            refY = (earliest.Y != null) ? earliest.Y.Value : 0.0;
        }

        // Apply to all
        int applied = 0;
        foreach (var item in selectedWithPiP)
        {
            try
            {
                if (item.Location2D != null)
                {
                    item.Location2D.IsAnimated = false;
                    var v = item.Location2D.Value;
                    v.X = refX;
                    v.Y = refY;
                    item.Location2D.Value = v;
                }
                else if (item.X != null && item.Y != null)
                {
                    item.X.IsAnimated = false;
                    item.Y.IsAnimated = false;
                    item.X.Value = refX;
                    item.Y.Value = refY;
                }
                else
                {
                    continue;
                }
                applied++;
            }
            catch
            {
                // Skip on error
            }
        }

        if (applied == 0)
        {
            MessageBox.Show(
                "Location could not be applied to any selected events.",
                "Copy PiP Location",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show(
                string.Format("Location applied to {0} selected event(s).", applied),
                "Copy PiP Location",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }

    private static bool IsTextGenerator(Take take)
    {
        try
        {
            var gen = take.Media.Generator;
            if (gen == null || gen.PlugIn == null) return false;

            string uid = gen.PlugIn.UniqueID ?? string.Empty;
            string name = gen.PlugIn.Name ?? string.Empty;

            if (uid.IndexOf("titlesandtext", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            if (name.IndexOf("Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0
             || name.IndexOf("VEGAS Titles & Text", StringComparison.OrdinalIgnoreCase) >= 0
             || name.IndexOf("Text", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }
        catch { }
        return false;
    }

    private static Effect FindPiP(VideoEvent ev, params string[] knownUids)
    {
        foreach (Effect fx in ev.Effects)
        {
            if (!fx.IsOFX || fx.PlugIn == null) continue;

            string uid = fx.PlugIn.UniqueID ?? string.Empty;
            foreach (var known in knownUids)
            {
                if (string.Equals(uid, known, StringComparison.OrdinalIgnoreCase))
                    return fx;
            }

            string name = fx.PlugIn.Name ?? string.Empty;
            if (name.IndexOf("Picture in Picture", StringComparison.OrdinalIgnoreCase) >= 0)
                return fx;
        }
        return null;
    }

    private static bool TryGetLocation(OFXEffect ofx,
        out OFXDouble2DParameter loc2D,
        out OFXDoubleParameter x,
        out OFXDoubleParameter y)
    {
        loc2D = ofx.FindParameterByName("Location") as OFXDouble2DParameter;
        if (loc2D != null)
        {
            x = null;
            y = null;
            return true;
        }

        x = ofx.FindParameterByName("Location X") as OFXDoubleParameter
          ?? ofx.FindParameterByName("Position X") as OFXDoubleParameter
          ?? ofx.FindParameterByName("Center X") as OFXDoubleParameter;

        y = ofx.FindParameterByName("Location Y") as OFXDoubleParameter
          ?? ofx.FindParameterByName("Position Y") as OFXDoubleParameter
          ?? ofx.FindParameterByName("Center Y") as OFXDoubleParameter;

        return (x != null && y != null);
    }

    private class EventWithPiPLocation
    {
        public VideoEvent Event;
        public OFXDouble2DParameter Location2D;
        public OFXDoubleParameter X;
        public OFXDoubleParameter Y;
    }
}
