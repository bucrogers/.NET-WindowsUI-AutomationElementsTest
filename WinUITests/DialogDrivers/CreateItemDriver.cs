using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

using ItemSystemTests.Utilities;

namespace ItemSystemTests.DialogDrivers
{
    public sealed class CreateItemDriver : DriverBase
    {
        private AutomationElement _nameCtl;
        private AutomationElement _descCtl;
        private AutomationElement _createCtl;
        private AutomationElement _cancelCtl;

        private CreateItemDriver(AutomationElement parentElement)
        {
            _parentElement = parentElement;
            _dialog = UIAUtility.FindElementByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _nameCtl = UIAUtility.FindElementById(_dialog, 
                    "NameTextBox");
            _descCtl = UIAUtility.FindElementById(_dialog, 
                    "DescTextBox");
            _createCtl = UIAUtility.FindElementById(_dialog, 
                    "OKButton");
            _cancelCtl = UIAUtility.FindElementById(_dialog, 
                    "CancelButton");

            WaitForDialogReady();
        }

        /// <summary>
        /// Get UI driver instance
        /// </summary>
        /// <param name="parent"></param>
        /// <returns></returns>
        /// <exception cref="TimeoutException">dialog not found</exception>
        /// <exception cref="ApplicationException">If any dialog control not found by its automation id</exception>
        public static CreateItemDriver FindFromParent(AutomationElement parent)
        {
            return new CreateItemDriver(parent);
        }

        public string Name
        {
            get
            {
                return UIAUtility.GetText(_nameCtl);
            }
            set
            {
                UpdatePropBagValue("name", value);
                UIAUtility.InsertText(_nameCtl, value);
            }
        }

        public string Desc
        {
            get
            {
                return UIAUtility.GetText(_descCtl);
            }
            set
            {
                UpdatePropBagValue("desc", value);
                UIAUtility.InsertText(_descCtl, value);
            }
        }

        /// <summary>
        /// Select Create button to invoke create item - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="ApplicationException">Create item failed</exception>
        public void SelectCreateVerifySuccess()
        {
            EnsureDialogEventsFired();

            UIAUtility.WaitForElementEnabledWithTimeout(_createCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            UIAUtility.PressButton(_createCtl);

            try
            {
                UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);
            }
            catch (TimeoutException)
            {
                throw new ApplicationException("Create item failed using name '" + GetPropBagValue("name") + "'");
            }
        }

        public void SelectCancelVerifySuccess()
        {
            UIAUtility.PressButton(_cancelCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        protected override string DialogId { get { return "CreateItemDialog"; } }

        protected override void EnsureDialogEventsFired()
        {
            //Ensure any dialog events are fired by switching focus between two input controls
            _nameCtl.SetFocus();
            _descCtl.SetFocus();
        }

        protected override void WaitForDialogReady()
        {
        }
    }
}
