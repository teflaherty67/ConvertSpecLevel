using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertSpecLevel.Classes
{
    internal class clsWindowData
    {
        public FamilyInstance WindowInstance { get; set; }
        public double CurrentHeadHeight { get; set; }
        public double CurrentWindowHeight { get; set; }
        public Parameter HeadHeightParam { get; set; }
        public Parameter WindowHeightParam { get; set; }

        public clsWindowData(FamilyInstance window)
        {
            WindowInstance = window;
            // Get the parameters you need to modify
            HeadHeightParam = window.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM); // or whatever the parameter name is
            WindowHeightParam = window.LookupParameter("Height"); // or sill height, etc.

            // Store current values
            CurrentHeadHeight = HeadHeightParam?.AsDouble() ?? 0.0;
            CurrentWindowHeight = WindowHeightParam?.AsDouble() ?? 0.0;
        }
    }
}
