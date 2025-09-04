using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertSpecLevel.Classes
{
    internal class clsCabSpecMap
    {
        internal class RefSpSettings
        {
            public bool IsVisible;
            public string TypeName;
            public double HeightOffsetFromLevel;
        }

        public static string GetMWCabHeight(string client, string specLevel)
        {
            string key = $"{client}-{specLevel}";
            return _mwHeights.TryGetValue(key, out string height) ? height : null;
        }

        public static RefSpSettings GetRefSpSettings(string client, string specLevel)
        {
            string key = $"{client}-{specLevel}";
            return _refSpSettings.TryGetValue(key, out RefSpSettings settings) ? settings : null;
        }

        private static readonly Dictionary<string, string> _mwHeights = new Dictionary<string, string>
        {
            // Complete Home mappings
            { "Central Texas-Complete Home", "21\"" },
            { "Dallas/Ft Worth-Complete Home", "24\"" },
            { "Florida-Complete Home", "18\"" },
            { "Houston-Complete Home", "21\"" },
            { "Maryland-Complete Home", "24\"" },
            { "Minnesota-Complete Home", "21\"" },
            { "Oklahoma-Complete Home", "21\"" },
            { "Pennsylvania-Complete Home", "21\"" },
            { "Southeast-Complete Home", "21\"" },           
            { "Virginia-Complete Home", "24\"" },
            { "West Virginia-Complete Home", "24\"" },
        
            // Complete Home Plus mappings  
            { "Central Texas-Complete Home Plus", "30\"" },
            { "Dallas/Ft Worth-Complete Home Plus", "30\"" },
            { "Florida-Complete Home Plus", "30\"" },
            { "Houston-Complete Home Plus", "30\"" },
            { "Maryland-Complete Home Plus", "30\"" },
            { "Oklahoma-Complete Home Plus", "30\"" },
            { "Pennsylvania-Complete Home Plus", "30\"" },
            { "Southeast-Complete Home Plus", "30\"" },
            { "Virginia-Complete Home Plus", "30\"" },
            { "West Virginia-Complete Home Plus", "30\"" },

        };

        private static readonly Dictionary<string, RefSpSettings> _refSpSettings = new Dictionary<string, RefSpSettings>
        {
            // Complete Home = RefSp cabinet hidden
            { "Central Texas-Complete Home", new RefSpSettings { IsVisible = true, TypeName = "39\"x18\"x15\"", HeightOffsetFromLevel = 6.0 } },
            { "Dallas/Ft Worth-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Florida-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Houston-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Maryland-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Minnesota-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Oklahoma-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Pennsylvania-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Southeast-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Virginia-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "West Virginia-Complete Home", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
        
            // Complete Home Plus = RefSp cabinet visible with client-specific settings
            { "Central Texas-Complete Home Plus", new RefSpSettings { IsVisible = true, TypeName = "39\"x27\"x15\"", HeightOffsetFromLevel = 6.25 } },
            { "Dallas/Ft Worth-Complete Home Plus", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Florida-Complete Home Plus", new RefSpSettings { IsVisible = true, TypeName = "39\"x24\"x15\"", HeightOffsetFromLevel = 6.0 } },
            { "Houston-Complete Home Plus", new RefSpSettings { IsVisible = true, TypeName = "39\"x24\"x15\"", HeightOffsetFromLevel = 6.5 } },
            { "Maryland-Complete Home Plus", new RefSpSettings {IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0} },
            { "Minnesota-Complete Home Plus", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Oklahoma-Complete Home Plus", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "Pennsylvania-Complete Home Plus", new RefSpSettings {IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0} },
            { "Southeast-Complete Home Plus", new RefSpSettings { IsVisible = true, TypeName = "39\"x24\"x15\"", HeightOffsetFromLevel = 6.5 } },
            { "Virginia-Complete Home Plus", new RefSpSettings { IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0 } },
            { "West Virginia-Complete Home Plus", new RefSpSettings {IsVisible = false, TypeName = "", HeightOffsetFromLevel = 0.0} },
            
        };
    }
}
