

namespace ConvertSpecLevel.Common
{
    internal static class Utils
    {
        #region Ribbon Panel
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        internal static View GetViewByNameContainsAndAssociatedLevel(Document curDoc, string v1, string v2)
        {
            throw new NotImplementedException();
        }

        internal static List<string> UpdateFloorFinishInActiveView(Document curDoc, string selectedSpecLevel)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
