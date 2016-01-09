using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;

using Microsoft.Test.Input;

using ItemSystemTests.Utilities;

namespace ItemSystemTests.DialogDrivers
{
    public sealed class PublishDriver : DriverBase
    {
        private AutomationElement _closeIconCtl;
        private AutomationElement _publishCtl;
        private AutomationElement _cancelCtl;

        private PublishDriver(AutomationElement parentElement)
        {
            _parentElement = parentElement;
            _dialog = UIAUtility.FindElementByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _publishCtl = UIAUtility.FindElementById(_dialog,
                    "PublishButton");
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
        public static PublishDriver FindFromParent(AutomationElement parent)
        {
            return new PublishDriver(parent);
        }

        /// <summary>
        /// Select Publish button to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectPublishVerifySuccess()
        {
            EnsureDialogEventsFired();

            UIAUtility.WaitForElementEnabledWithTimeout(_publishCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            UIAUtility.PressButton(_publishCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        /// <summary>
        /// Select Close Icon to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectCloseIconVerifySuccess()
        {
            UIAUtility.PressButton(_closeIconCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        /// <summary>
        /// Select Esc key to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectEscVerifySuccess()
        {
            Keyboard.Press(Key.Escape);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        protected override string DialogId { get { return "PublishItemsDialog"; } }

        protected override void EnsureDialogEventsFired()
        {
        }

        protected override void WaitForDialogReady()
        {
            UIAUtility.WaitForElementEnabledWithTimeout(_publishCtl, AddinTestUtility.FloorDialogInitDelay);
        }
    }
}
