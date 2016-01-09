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
    public sealed class MessageBoxDriver : DriverBase
    {
        private AutomationElement _closeIconCtl;
        private AutomationElement _okCtl;

        private MessageBoxDriver(AutomationElement parentElement)
        {
            _parentElement = parentElement;

            _dialog = UIAUtility.FindModalDialogWithTimeout(_parentElement, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _okCtl = UIAUtility.FindElementById(_dialog,
                    "1");

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
        public static MessageBoxDriver FindFromParent(AutomationElement parent)
        {
            return new MessageBoxDriver(parent);
        }

        /// <summary>
        /// Select OK button to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectOKVerifySuccess()
        {
            UIAUtility.PressButton(_okCtl);

            UIAUtility.WaitForModalDialogToDisappearWithTimeout(_parentElement, AddinTestUtility.DialogCancelTimeout);
        }

        /// <summary>
        /// Select Close Icon to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectCloseIconVerifySuccess()
        {
            UIAUtility.PressButton(_closeIconCtl);

            UIAUtility.WaitForModalDialogToDisappearWithTimeout(_parentElement, AddinTestUtility.DialogCancelTimeout);
        }

        /// <summary>
        /// Select Esc key to close dialog - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="TimeoutException">dialog did not close</exception>
        public void SelectEscVerifySuccess()
        {
            Keyboard.Press(Key.Escape);

            UIAUtility.WaitForModalDialogToDisappearWithTimeout(_parentElement, AddinTestUtility.DialogCancelTimeout);
        }

        protected override string DialogId 
        { 
            get 
            {
                throw new NotImplementedException("use _dialogName");
            } 
        }

        protected override void EnsureDialogEventsFired()
        {
        }

        protected override void WaitForDialogReady()
        {
            UIAUtility.WaitForElementEnabledWithTimeout(_okCtl, AddinTestUtility.FloorDialogInitDelay);
        }

    }
}
