using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Threading;
using System.Diagnostics;

using ItemSystemTests.Utilities;

namespace ItemSystemTests.DialogDrivers
{
    public sealed class CreateAccountDriver : DriverBase
    {
        public enum ErrorMessageType
        {
            None,
            GenericServer
        }
        private AutomationElement _nameCtl;
        private AutomationElement _emailCtl;
        private AutomationElement _passwordCtl;
        private AutomationElement _termsCtl;
        private AutomationElement _rememberMeCtl;
        private AutomationElement _registerCtl;
        private AutomationElement _cancelCtl;
        private AutomationElement _genericServerErrorErrMsgCtl;

        private AutomationElement _closeIconCtl;

        private CreateAccountDriver(AutomationElement parentElement)
        {
            _parentElement = parentElement;
            _dialog = UIAUtility.FindElementByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _nameCtl = UIAUtility.FindElementById(_dialog,
                    "NameTextBox");
            _emailCtl = UIAUtility.FindElementById(_dialog,
                    "EmailTextBox");
            _passwordCtl = UIAUtility.FindElementById(_dialog,
                    "PasswordTextBox");
            _termsCtl = UIAUtility.FindElementById(_dialog,
                    "TermsCheckbox");
            _rememberMeCtl = UIAUtility.FindElementById(_dialog,
                    "RememberMeCheckbox");
            _registerCtl = UIAUtility.FindElementById(_dialog,
                    "OKButton");
            _cancelCtl = UIAUtility.FindElementById(_dialog,
                    "CancelButton");
            _genericServerErrorErrMsgCtl = UIAUtility.FindElementByIdChildByClassname(_dialog,
                    "GenericServerErrorErrorMessage", "TextBlock", TreeScope.Descendants);
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
        public static CreateAccountDriver FindFromParent(AutomationElement parent)
        {
            return new CreateAccountDriver(parent);
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

        public string Email
        {
            get
            {
                return UIAUtility.GetText(_emailCtl);
            }
            set
            {
                UpdatePropBagValue("email", value);
                UIAUtility.InsertText(_emailCtl, value);
            }
        }

        public string Password
        {
            get
            {
                return UIAUtility.GetText(_passwordCtl);
            }
            set
            {
                UpdatePropBagValue("password", value);
                UIAUtility.InsertText(_passwordCtl, value);
            }
        }

        public bool Terms
        {
            get
            {
                return UIAUtility.GetCheckboxValue(_termsCtl);
            }
            set
            {
                UpdatePropBagValue("terms", value.ToString());
                UIAUtility.SetCheckbox(_termsCtl, value);
            }
        }

        public bool RememberMe
        {
            get
            {
                return UIAUtility.GetCheckboxValue(_rememberMeCtl);
            }
            set
            {
                UpdatePropBagValue("rememberme", value.ToString());
                UIAUtility.SetCheckbox(_rememberMeCtl, value);
            }
        }

        public KeyValuePair<ErrorMessageType, string> CurrentErrorMessage
        {
            get
            {
                if ( ! _genericServerErrorErrMsgCtl.Current.IsOffscreen)
                {
                    return new KeyValuePair<ErrorMessageType,string>(ErrorMessageType.GenericServer, _genericServerErrorErrMsgCtl.Current.Name);
                }

                return new KeyValuePair<ErrorMessageType,string>(ErrorMessageType.None, string.Empty);
            }
        }

        public bool RegisterEnabled
        {
            get { return _registerCtl.Current.IsEnabled;  }
        }

        /// <summary>
        /// Select Register button - validate success by ensuring dialog closes, returning the modal messagebox
        /// </summary>
        /// <exception cref="ApplicationException">Register failed</exception>
        /// <returns>Modal message box driver</returns>
        public MessageBoxDriver SelectRegister()
        {
            EnsureDialogEventsFired();
    
            UIAUtility.WaitForElementEnabledWithTimeout(_registerCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            var priorErrMsg = CurrentErrorMessage;
            UIAUtility.PressButton(_registerCtl);

            var sw = Stopwatch.StartNew();
            do
            {
                //Determine if window closed (meaning authenticated)
                var win = _parentElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty,
                        DialogId, PropertyConditionFlags.IgnoreCase));
                if (win == null)
                {
                    //Login dialog closed, means successfully authenticated
                    return MessageBoxDriver.FindFromParent(_parentElement);
                }

                if (CurrentErrorMessage.Key != priorErrMsg.Key
                        && CurrentErrorMessage.Key != ErrorMessageType.None)
                {
                    //ErrorMessage has appeared
                    throw new ApplicationException("Register failed using email '" + GetPropBagValue("email") + "' with error " + CurrentErrorMessage.Key + ": '" + CurrentErrorMessage.Value + "'");
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < AddinTestUtility.AuthenticationTimeout);

            //Timed out waiting

            throw new ApplicationException("Register failed using email '" + GetPropBagValue("email") + "' with error " + CurrentErrorMessage.Key + ": '" + CurrentErrorMessage.Value + "'");
        }

        public void SelectCancelVerifySuccess()
        {
            UIAUtility.PressButton(_cancelCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        protected override string DialogId { get { return "NewUserDialog"; } }

        protected override void EnsureDialogEventsFired()
        {
            //Ensure any dialog events are fired by switching focus between two input controls
            _emailCtl.SetFocus();
            _passwordCtl.SetFocus();
        }

        protected override void WaitForDialogReady()
        {
        }
    }
}
