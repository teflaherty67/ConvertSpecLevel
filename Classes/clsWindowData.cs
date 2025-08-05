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
        public double CurHeadHeight { get; set; }
        public double CurWindowHeight { get; set; }
        public Parameter HeadHeightParam { get; set; }
        public Parameter WindowHeightParam { get; set; }

        public clsWindowData(FamilyInstance window)
        {
            WindowInstance = window;
            // Get the parameters you need to modify
            HeadHeightParam = window.get_Parameter(BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM); // get the current head height
            WindowHeightParam = window.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM); // get the current height

            // Store current values
            CurHeadHeight = HeadHeightParam?.AsDouble() ?? 0.0;
            CurWindowHeight = WindowHeightParam?.AsDouble() ?? 0.0;
        }
    }
}
