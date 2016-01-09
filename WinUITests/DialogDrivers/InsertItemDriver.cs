using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Automation;
using System.Threading;

using ItemSystemTests.Utilities;

namespace ItemSystemTests.DialogDrivers
{
    public sealed class InsertItemDriver : DriverBase
    {
        private AutomationElement _searchBoxCtl;
        private AutomationElement _itemDataGridCtl;
        private AutomationElement _allowPublishCheckboxCtl;
        private AutomationElement _insertCtl;
        private AutomationElement _cancelCtl;
        private AutomationElement _closeIconCtl;

        private InsertItemDriver(AutomationElement parentElement)
        {
            _parentElement = parentElement;
            _dialog = UIAUtility.FindElementByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _searchBoxCtl = UIAUtility.FindElementById(_dialog,
                    "SearchTextBox", TreeScope.Descendants);
            _itemDataGridCtl = UIAUtility.FindElementById(_dialog, 
                    "FindListView");
            _allowPublishCheckboxCtl = UIAUtility.FindElementById(_dialog, 
                    "AsOutputCheckBox");
            _insertCtl = UIAUtility.FindElementById(_dialog, 
                    "InsertButton");
            _cancelCtl = UIAUtility.FindElementById(_dialog, 
                    "CancelButton");
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
        public static InsertItemDriver FindFromParent(AutomationElement parent)
        {
            return new InsertItemDriver(parent);
        }

        public string Search
        {
            get
            {
                return UIAUtility.GetText(_searchBoxCtl);
            }
            set
            {
                UIAUtility.InsertText(_searchBoxCtl, value);
            }
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

                //Use search to narrow down the scroll area, for more reliable finding
                Search = value;

                //Find first occurrence of sought item in grid
                const int ItemNameCol = 1;
                var selectedRow = UIAUtility.SelectGridRow(_itemDataGridCtl, value, ItemNameCol);

                //Wait for indication that post-select dialog events have been handled
                 UIAUtility.WaitForElementEnabledWithTimeout(_insertCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);
                 UIAUtility.WaitForElementEnabledWithTimeout(_allowPublishCheckboxCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

                //Additional delay to allow for dialog events to possibly reset scrolling to make not visible
                 Thread.Sleep(AddinTestUtility.DialogControlEventStateUpdateDelay);

                //Ensure selected row is visible after select and dialog events processed
                UIAUtility.WaitForElementVisibleWithTimeout(selectedRow, AddinTestUtility.DialogControlEventStateUpdateTimeout, value);
            }
        }

        public bool AllowPublish
        {
            get 
            {
                return UIAUtility.GetCheckboxValue(_allowPublishCheckboxCtl);
            }
            set
            {
                UpdatePropBagValue("allowpublish", value.ToString());
                UIAUtility.SetCheckbox(_allowPublishCheckboxCtl, value);
            }
        }

        /// <summary>
        /// Select Insert button to invoke insert item - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="ApplicationException">Insert item failed</exception>
        public void SelectInsertVerifySuccess()
        {
            EnsureDialogEventsFired();

            UIAUtility.WaitForElementEnabledWithTimeout(_insertCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            UIAUtility.PressButton(_insertCtl);

            try
            {
                UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);
            }
            catch (TimeoutException)
            {
                throw new ApplicationException("Insert item failed for item '" + GetPropBagValue("selecteditem") + "'");
            }
        }

        public void SelectCancelVerifySuccess()
        {
            UIAUtility.PressButton(_cancelCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        protected override string DialogId { get { return "InsertItemDialog"; } }

        protected override void EnsureDialogEventsFired()
        {
            //Ensure any dialog events are fired by switching focus between two input controls
            _itemDataGridCtl.SetFocus();
            _allowPublishCheckboxCtl.SetFocus();
        }

        protected override void WaitForDialogReady()
        {
            var itemDataGridPattern = (GridPattern)_itemDataGridCtl.GetCurrentPattern(GridPattern.Pattern);
            UIAUtility.WaitForPopulatedDatagrid(itemDataGridPattern, AddinTestUtility.WebServiceResponsePopulateTimeout);
        }
    }
}
