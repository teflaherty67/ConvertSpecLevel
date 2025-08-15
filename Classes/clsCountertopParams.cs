using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertSpecLevel.Classes
{
    internal class clsCountertopParams
    {
        public double MitreLeftOther { get; set; }
        public bool MitreLeft45Out { get; set; }
        public bool MitreLeft45In { get; set; }
        public bool MitreLeft22_5In { get; set; }
        public double MitreRightOther { get; set; }
        public bool MitreRight45Out { get; set; }
        public bool MitreRight45In { get; set; }
        public bool MitreRight22_5In { get; set; }       
        public bool Sink { get; set; }
        public bool BacksplashLeft { get; set; }
        public bool BacksplashRight { get; set; }
        public bool BacksplashBack { get; set; }
        public double AltSinkLoc { get; set; }

        // Constructor to capture from existing element
        public clsCountertopParams(FamilyInstance countertop)
        {
            // Get the MitreLeftOther parameter
            Parameter paramMitreLeftOther = countertop.LookupParameter("Mitre Left Other");
            MitreLeftOther = paramMitreLeftOther?.AsDouble() ?? 0.0;

            // Get the MitreLeft45Out parameter  
            Parameter paramMitreLeft45Out = countertop.LookupParameter("Mitre Left 45 Out");
            MitreLeft45Out = paramMitreLeft45Out?.AsInteger() == 1;

            // Get the MitreLeft45In parameter  
            Parameter paramMitreLeft45In = countertop.LookupParameter("Mitre Left 45 In");
            MitreLeft45In = paramMitreLeft45In?.AsInteger() == 1;

            // Get the MitreLeft22_5In parameter  
            Parameter paramMitreLeft22_5In = countertop.LookupParameter("Mitre Left 22.5 In");
            MitreLeft22_5In = paramMitreLeft22_5In?.AsInteger() == 1;

            // Get the MitreRightOther parameter
            Parameter paramMitreRightOther = countertop.LookupParameter("Mitre Right Other");
            MitreRightOther = paramMitreRightOther?.AsDouble() ?? 0.0;

            // Get the MitreRight45Out parameter  
            Parameter paramMitreRight45Out = countertop.LookupParameter("Mitre Right 45 Out");
            MitreRight45Out = paramMitreRight45Out?.AsInteger() == 1;

            // Get the MitreRight45In parameter  
            Parameter paramMitreRight45In = countertop.LookupParameter("Mitre Right 45 In");
            MitreRight45In = paramMitreRight45In?.AsInteger() == 1;

            // Get the MitreRight22_5In parameter  
            Parameter paramMitreRight22_5In = countertop.LookupParameter("Mitre Right 22.5 In");
            MitreRight22_5In = paramMitreRight22_5In?.AsInteger() == 1;

            // Get the Sink parameter  
            Parameter paramSinkYesNo = countertop.LookupParameter("Sink");
            Sink = paramSinkYesNo?.AsInteger() == 1;

            // Get the BacksplashLeft parameter  
            Parameter paramBacksplashLeft = countertop.LookupParameter("Backsplash Left");
            BacksplashLeft = paramBacksplashLeft?.AsInteger() == 1;

            // Get the BacksplashRight parameter  
            Parameter paramBacksplashRight = countertop.LookupParameter("Backsplash Right");
            BacksplashRight = paramBacksplashRight?.AsInteger() == 1;

            // Get the BacksplashBack parameter  
            Parameter paramBacksplashBack = countertop.LookupParameter("Backsplash Back");
            BacksplashBack = paramBacksplashBack?.AsInteger() == 1;

            // Get the AltSinkLoc parameter
            Parameter paramAltSinkLoc = countertop.LookupParameter("Alternate Sink Location");
            AltSinkLoc = paramAltSinkLoc?.AsDouble() ?? 0.0;
        }

        // Method to restore to new element  
        public void RestoreToElement(FamilyInstance countertop)
        {
            // Restore MitreLeftOther parameter
            Parameter paramMitreLeftOther = countertop.LookupParameter("Mitre Left Other");
            paramMitreLeftOther?.Set(MitreLeftOther);

            // Restore MitreLeft45Out parameter
            Parameter paramMitreLeft45Out = countertop.LookupParameter("Mitre Left 45 Out");
            paramMitreLeft45Out?.Set(MitreLeft45Out ? 1 : 0);

            // Restore MitreLeft45In parameter
            Parameter paramMitreLeft45In = countertop.LookupParameter("Mitre Left 45 In");
            paramMitreLeft45In?.Set(MitreLeft45In ? 1 : 0);

            // Restore MitreLeft22_5In parameter  
            Parameter paramMitreLeft22_5In = countertop.LookupParameter("Mitre Left 22.5 In");
            paramMitreLeft22_5In?.Set(MitreLeft22_5In ? 1 : 0);

            // Restore MitreRightOther parameter
            Parameter paramMitreRightOther = countertop.LookupParameter("Mitre Right Other");
            paramMitreRightOther?.Set(MitreRightOther);

            // Restore MitreRight45Out parameter  
            Parameter paramMitreRight45Out = countertop.LookupParameter("Mitre Right 45 Out");
            paramMitreRight45Out?.Set(MitreRight45Out ? 1 : 0);

            // Restore MitreRight45In parameter  
            Parameter paramMitreRight45In = countertop.LookupParameter("Mitre Right 45 In");
            paramMitreRight45In?.Set(MitreRight45In ? 1 : 0);

            // Restore MitreRight22_5In parameter  
            Parameter paramMitreRight22_5In = countertop.LookupParameter("Mitre Right 22.5 In");
            paramMitreRight22_5In?.Set(MitreRight22_5In ? 1 : 0);

            // Restore Sink parameter  
            Parameter paramSinkYesNo = countertop.LookupParameter("Sink");
            paramSinkYesNo?.Set(Sink ? 1 : 0);

            // Restore BacksplashLeft parameter  
            Parameter paramBacksplashLeft = countertop.LookupParameter("Backsplash Left");
           paramBacksplashLeft?.Set(BacksplashLeft ? 1 : 0);

            // Restore BacksplashRight parameter  
            Parameter paramBacksplashRight = countertop.LookupParameter("Backsplash Right");
            paramBacksplashRight?.Set(BacksplashRight ? 1 : 0);

            // Restore BacksplashBack parameter  
            Parameter paramBacksplashBack = countertop.LookupParameter("Backsplash Back");
            paramBacksplashBack?.Set(BacksplashBack ? 1 : 0);

            // Restore AltSinkLoc parameter
            Parameter paramAltSinkLoc = countertop.LookupParameter("Alternate Sink Location");
            paramAltSinkLoc?.Set(AltSinkLoc);
        }
    }
}
