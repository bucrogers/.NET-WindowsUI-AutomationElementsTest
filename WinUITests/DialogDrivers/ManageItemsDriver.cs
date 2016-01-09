using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Threading;

using ItemSystemTests.Utilities;

namespace ItemSystemTests.DialogDrivers
{
    public sealed class ManageItemsDriver : DriverBase
    {
        private AutomationElement _itemDataGridCtl;
        private AutomationElement _refreshCtl;
        private AutomationElement _editCtl;
        private AutomationElement _unlinkCtl;
        private AutomationElement _closeCtl;
        private AutomationElement _closeIconCtl;
        
        private ManageItemsDriver(AutomationElement parentElement)
        {
            _parentElement = parentElement;
            _dialog = UIAUtility.FindElementByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _itemDataGridCtl = UIAUtility.FindElementById(_dialog,
                        "ItemListView");

            _refreshCtl = UIAUtility.FindElementById(_dialog,
                        "RefreshButton");
            _editCtl = UIAUtility.FindElementById(_dialog,
                        "EditButton");
            _unlinkCtl = UIAUtility.FindElementById(_dialog,
                        "UnlinkButton");
            _closeCtl = UIAUtility.FindElementById(_dialog,
                        "CloseButton");
            _closeIconCtl = UIAUtility.FindElementById(_dialog,
                        "Close", TreeScope.Descendants);

            WaitForDialogReady();
        }

        /// <summary>
        /// Get Dialog UI driver instance
        /// </summary>
        /// <param name="parent"></param>
        /// <returns></returns>
        /// <exception cref="TimeoutException">dialog not found</exception>
        /// <exception cref="ApplicationException">If any dialog control not found by its automation id</exception>
        /// <exception cref="ApplicationException">If there is not at least one item (after waiting for grid to populate)</exception>
        public static ManageItemsDriver FindFromParent(AutomationElement parent)
        {
            return new ManageItemsDriver(parent);
        }

        public string SelectedItem
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                UpdatePropBagValue("selecteditem", value);

                //Find first occurrence of sought item in grid
                const int ItemNameCol = 1;
                var selectedRow = UIAUtility.SelectGridRow(_itemDataGridCtl, value, ItemNameCol);

                //Wait for indication that post-select dialog events have been handled
                UIAUtility.WaitForElementEnabledWithTimeout(_refreshCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);
                UIAUtility.WaitForElementEnabledWithTimeout(_editCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);
                UIAUtility.WaitForElementEnabledWithTimeout(_unlinkCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

                //Additional delay to allow for dialog events to possibly reset scrolling to make not visible
                Thread.Sleep(AddinTestUtility.DialogControlEventStateUpdateDelay);

                //Ensure selected row is visible after select and dialog events processed
                UIAUtility.WaitForElementVisibleWithTimeout(selectedRow, AddinTestUtility.DialogControlEventStateUpdateTimeout, value);
            }
        }

        /// <summary>
        /// Select Close button to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectCloseButtonVerifySuccess()
        {
            EnsureDialogEventsFired();

            UIAUtility.WaitForElementEnabledWithTimeout(_closeCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            UIAUtility.PressButton(_closeCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);
        }

        protected override string DialogId { get { return "ManageItemsDialog"; } }

        protected override void EnsureDialogEventsFired()
        {
        }

        protected override void WaitForDialogReady()
        {
            var itemDataGridPattern = (GridPattern)_itemDataGridCtl.GetCurrentPattern(GridPattern.Pattern);
            UIAUtility.WaitForPopulatedDatagrid(itemDataGridPattern, AddinTestUtility.WebServiceResponsePopulateTimeout);
        }
    }
}
