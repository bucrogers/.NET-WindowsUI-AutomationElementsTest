using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ItemSystemTests.Utilities;

namespace ItemSystemTests
{
    public class ExpectedButtonStates
    {
        public static SortedDictionary<AddinRibbonButton, RibbonButtonState> LoggedInButtonStates = new SortedDictionary<AddinRibbonButton, RibbonButtonState>() 
        { 
            {AddinRibbonButton.About, RibbonButtonState.Enabled},
            {AddinRibbonButton.Create, RibbonButtonState.Enabled},
            {AddinRibbonButton.EditDetails, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Help, RibbonButtonState.Enabled},
            {AddinRibbonButton.History, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.InsertItem, RibbonButtonState.Enabled},
            {AddinRibbonButton.Login, RibbonButtonState.NonVisible},
            {AddinRibbonButton.Logout, RibbonButtonState.Enabled},
            {AddinRibbonButton.ManageItems, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Publish, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.PublishItems, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Refresh, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.RefreshItems, RibbonButtonState.VisibleNotEnabled}
        };

        public static SortedDictionary<AddinRibbonButton, RibbonButtonState> NotLoggedInButtonStates = new SortedDictionary<AddinRibbonButton, RibbonButtonState>() 
        { 
            {AddinRibbonButton.About, RibbonButtonState.Enabled},
            {AddinRibbonButton.Create, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.EditDetails, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Help, RibbonButtonState.Enabled},
            {AddinRibbonButton.History, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.InsertItem, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Login, RibbonButtonState.Enabled},
            {AddinRibbonButton.Logout, RibbonButtonState.NonVisible},
            {AddinRibbonButton.ManageItems, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Publish, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.PublishItems, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Refresh, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.RefreshItems, RibbonButtonState.VisibleNotEnabled}
        };

        public static SortedDictionary<AddinRibbonButton, RibbonButtonState> OnExistingAllowPublishItemSheetOffItemCell = new SortedDictionary<AddinRibbonButton, RibbonButtonState>() 
        { 
            {AddinRibbonButton.About, RibbonButtonState.Enabled},
            {AddinRibbonButton.Create, RibbonButtonState.Enabled},
            {AddinRibbonButton.EditDetails, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Help, RibbonButtonState.Enabled},
            {AddinRibbonButton.History, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.InsertItem, RibbonButtonState.Enabled},
            {AddinRibbonButton.Login, RibbonButtonState.NonVisible},
            {AddinRibbonButton.Logout, RibbonButtonState.Enabled},
            {AddinRibbonButton.ManageItems, RibbonButtonState.Enabled},
            {AddinRibbonButton.Publish, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.PublishItems, RibbonButtonState.Enabled},
            {AddinRibbonButton.Refresh, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.RefreshItems, RibbonButtonState.VisibleNotEnabled}
        };

        public static SortedDictionary<AddinRibbonButton, RibbonButtonState> OnNonPublishableItemCell = new SortedDictionary<AddinRibbonButton, RibbonButtonState>() 
        { 
            {AddinRibbonButton.About, RibbonButtonState.Enabled},
            {AddinRibbonButton.Create, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.EditDetails, RibbonButtonState.Enabled},
            {AddinRibbonButton.Help, RibbonButtonState.Enabled},
            {AddinRibbonButton.History, RibbonButtonState.Enabled},
            {AddinRibbonButton.InsertItem, RibbonButtonState.Enabled},
            {AddinRibbonButton.Login, RibbonButtonState.NonVisible},
            {AddinRibbonButton.Logout, RibbonButtonState.Enabled},
            {AddinRibbonButton.ManageItems, RibbonButtonState.Enabled},
            {AddinRibbonButton.Publish, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.PublishItems, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.Refresh, RibbonButtonState.Enabled},
            {AddinRibbonButton.RefreshItems, RibbonButtonState.Enabled}
        };

        public static SortedDictionary<AddinRibbonButton, RibbonButtonState> OnPublishableItemCell = new SortedDictionary<AddinRibbonButton, RibbonButtonState>() 
        { 
            {AddinRibbonButton.About, RibbonButtonState.Enabled},
            {AddinRibbonButton.Create, RibbonButtonState.VisibleNotEnabled},
            {AddinRibbonButton.EditDetails, RibbonButtonState.Enabled},
            {AddinRibbonButton.Help, RibbonButtonState.Enabled},
            {AddinRibbonButton.History, RibbonButtonState.Enabled},
            {AddinRibbonButton.InsertItem, RibbonButtonState.Enabled},
            {AddinRibbonButton.Login, RibbonButtonState.NonVisible},
            {AddinRibbonButton.Logout, RibbonButtonState.Enabled},
            {AddinRibbonButton.ManageItems, RibbonButtonState.Enabled},
            {AddinRibbonButton.Publish, RibbonButtonState.Enabled},
            {AddinRibbonButton.PublishItems, RibbonButtonState.Enabled},
            {AddinRibbonButton.Refresh, RibbonButtonState.Enabled},
            {AddinRibbonButton.RefreshItems, RibbonButtonState.VisibleNotEnabled}
        };
    }
}
