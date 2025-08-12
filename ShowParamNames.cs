using System.Windows.Forms;
using ScriptPortal.Vegas;

public class EntryPoint
{
    public void FromVegas(Vegas vegas)
    {
        // Find the first selected VideoEvent
        VideoEvent firstSelected = null;
        foreach (Track track in vegas.Project.Tracks)
        {
            var vtrack = track as VideoTrack;
            if (vtrack == null) continue;

            foreach (TrackEvent te in vtrack.Events)
            {
                if (!te.Selected) continue;

                VideoEvent videoEvent = te as VideoEvent;
                if (videoEvent == null) continue;
                
                firstSelected = videoEvent;
                break;
            }
            if (firstSelected != null) break;
        }

        if (firstSelected == null)
        {
            MessageBox.Show("No selected video event.");
            return;
        }

        // For each FX on that event, show a dialog with parameter names
        foreach (Effect fx in firstSelected.Effects)
        {
            string fxName = (fx.PlugIn != null && !string.IsNullOrEmpty(fx.PlugIn.Name))
                ? fx.PlugIn.Name
                : "(unknown FX)";

            if (!fx.IsOFX || fx.OFXEffect == null)
            {
                MessageBox.Show(fxName + "\n(not an OFX effect)");
                continue;
            }

            var ofx = fx.OFXEffect;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(fxName);
            sb.AppendLine("OFX Effect parameter dump");

            int i = 0;
            foreach (OFXParameter p in ofx.Parameters)
            {
                string type = p.GetType().Name;
                string name = p.Name != null ? p.Name : "(null)";
                string val = GetParamValueString(p);

                string anim;
                if (p is OFXDouble2DParameter && ((OFXDouble2DParameter)p).IsAnimated) anim = "animated";
                else if (p is OFXDoubleParameter && ((OFXDoubleParameter)p).IsAnimated) anim = "animated";
                else if (p is OFXBooleanParameter && ((OFXBooleanParameter)p).IsAnimated) anim = "animated";
                else if (p is OFXStringParameter && ((OFXStringParameter)p).IsAnimated) anim = "animated";
                else if (p is OFXRGBAParameter && ((OFXRGBAParameter)p).IsAnimated) anim = "animated";
                else if (p is OFXChoiceParameter && ((OFXChoiceParameter)p).IsAnimated) anim = "animated";
                else anim = "static";

                sb.AppendLine(i.ToString("D3") + " | " + type + " | " + anim + " | " + name + " = " + val);
                i++;
            }

            string dump = sb.ToString();
            if (dump.Length > 1400)
                MessageBox.Show(dump.Substring(0, 1400) + "\n... (truncated)", "FX Parameters");
            else
                MessageBox.Show(dump, "FX Parameters");
        }
    }

    private static string GetParamValueString(OFXParameter p)
    {
        try
        {
            if (p is OFXDouble2DParameter)
            {
                OFXDouble2D v = ((OFXDouble2DParameter)p).Value;
                return v.X + ", " + v.Y;
            }
            if (p is OFXDoubleParameter)
                return ((OFXDoubleParameter)p).Value.ToString();
            if (p is OFXBooleanParameter)
                return ((OFXBooleanParameter)p).Value.ToString();
            if (p is OFXStringParameter)
                return ((OFXStringParameter)p).Value ?? "";
            if (p is OFXChoiceParameter)
                return ((OFXChoiceParameter)p).Value.ToString();

            // OFXRGBAParameter exists, but its Value type isn't exposed as OFXRGBA here.
            if (p is OFXRGBAParameter)
                return "(color value not displayed)";
        }
        catch
        {
            // some params may throw on Value read
        }
        return "(unreadable)";
    }
}