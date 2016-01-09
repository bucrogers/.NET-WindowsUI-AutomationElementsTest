using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;

using ItemSystemTests.Utilities;

namespace ItemSystemTests.DialogDrivers
{
    public sealed class LoginDriver : DriverBase
    {
        public enum ErrorMessageType
        { 
            None,
            Credentials,
            ServerUnavailable,
            GenericServer
        }

        private AutomationElement _emailCtl;
        private AutomationElement _passwordCtl;
        private AutomationElement _rememberMeCtl;
        private AutomationElement _okCtl;
        private AutomationElement _cancelCtl;
        private AutomationElement _passwordChangeCtl;
        private AutomationElement _newAccountCtl;
        private AutomationElement _credentialsErrMsgCtl;
        private AutomationElement _genericServiceUnavailableErrMsgCtl;
        private AutomationElement _genericServerErrorErrMsgCtl;

        private AutomationElement _closeIconCtl;

        private LoginDriver(AutomationElement parentElement) 
        {
            _parentElement = parentElement;
            _dialog = UIAUtility.FindElementByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.WebServiceResponsePopulateTimeout);

            _emailCtl = UIAUtility.FindElementById(_dialog, 
                    "EmailTextBox");
            _passwordCtl = UIAUtility.FindElementById(_dialog, 
                    "PasswordTextBox");
            _rememberMeCtl = UIAUtility.FindElementById(_dialog, 
                    "RememberMeCheckBox");
            _okCtl = UIAUtility.FindElementById(_dialog, 
                    "OKButton");
            _cancelCtl = UIAUtility.FindElementById(_dialog,
                    "CancelButton");
            _passwordChangeCtl = UIAUtility.FindElementById(_dialog,
                    "PasswordChangeLink", TreeScope.Descendants);
            _newAccountCtl = UIAUtility.FindElementById(_dialog,
                    "NewAccountLink", TreeScope.Descendants);
            _credentialsErrMsgCtl = UIAUtility.FindElementByIdChildByClassname(_dialog,
                    "CredentialsErrorMessage", "TextBlock", TreeScope.Descendants);
            _genericServiceUnavailableErrMsgCtl = UIAUtility.FindElementByIdChildByClassname(_dialog,
                    "GenericServiceUnavailableErrorMessage", "TextBlock", TreeScope.Descendants);
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
        public static LoginDriver FindFromParent(AutomationElement parent)
        {
            return new LoginDriver(parent);
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
                if ( ! _credentialsErrMsgCtl.Current.IsOffscreen)
                {
                    return new KeyValuePair<ErrorMessageType, string>(ErrorMessageType.Credentials, _credentialsErrMsgCtl.Current.Name);
                }
                if ( ! _genericServiceUnavailableErrMsgCtl.Current.IsOffscreen)
                {
                    return new KeyValuePair<ErrorMessageType, string>(ErrorMessageType.ServerUnavailable, _genericServiceUnavailableErrMsgCtl.Current.Name);
                }
                if ( ! _genericServerErrorErrMsgCtl.Current.IsOffscreen)
                {
                    return new KeyValuePair<ErrorMessageType, string>(ErrorMessageType.GenericServer, _genericServerErrorErrMsgCtl.Current.Name);
                }

                return new KeyValuePair<ErrorMessageType, string>(ErrorMessageType.None, string.Empty);

            }
        }

        /// <summary>
        /// Select OK button to invoke login - validate success by ensuring dialog closes
        /// </summary>
        /// <exception cref="ApplicationException">Login failed</exception>
        public void SelectOkVerifySuccess()
        {
            EnsureDialogEventsFired();

            UIAUtility.WaitForElementEnabledWithTimeout(_okCtl, TimeSpan.FromSeconds(5));

            var priorErrMsg = CurrentErrorMessage;
            UIAUtility.PressButton(_okCtl);

            var sw = Stopwatch.StartNew();
            do
            {
                //Determine if window closed (meaning authenticated)
                var win = _parentElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty,
                        DialogId, PropertyConditionFlags.IgnoreCase));
                if (win == null)
                {
                    //Login dialog closed, means successfully authenticated
                    return;
                }

                if (CurrentErrorMessage.Key != priorErrMsg.Key
                        && CurrentErrorMessage.Key != ErrorMessageType.None)
                {
                    //ErrorMessage has appeared
                    throw new ApplicationException("Login failed using email '" + GetPropBagValue("email") + "' with error " + CurrentErrorMessage.Key + ": '" + CurrentErrorMessage.Value + "'");
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < AddinTestUtility.AuthenticationTimeout);

            //Timed out waiting
            throw new ApplicationException("Login failed using email '" + GetPropBagValue("email") + "' with error " + CurrentErrorMessage.Key + ": '" + CurrentErrorMessage.Value + "' after waiting " + AddinTestUtility.AuthenticationTimeout.TotalSeconds + "sec");
        }

        /// <summary>
        /// Select OK button to invoke login - validate failure by ensuring dialog stays open
        /// </summary>
        /// <exception cref="ApplicationException">Login failed</exception>
        public void SelectOkVerifyFailure(ErrorMessageType expectedErrorMessageType)
        {
            EnsureDialogEventsFired();

            UIAUtility.WaitForElementEnabledWithTimeout(_okCtl, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            var priorErrMsg = CurrentErrorMessage;

            UIAUtility.PressButton(_okCtl);

            var sw = Stopwatch.StartNew();
            do
            {
                //Determine if window closed (meaning authenticated)
                var win = _parentElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty,
                        DialogId, PropertyConditionFlags.IgnoreCase));
                if (win == null)
                {
                    throw new ApplicationException("Authentication completed when it was not expected to for Email '" + GetPropBagValue("email") + "'");
                }

                if (CurrentErrorMessage.Key != priorErrMsg.Key)
                {
                    //ErrorMessageType has changed - examine
                    if (CurrentErrorMessage.Key != expectedErrorMessageType)
                    {
                        throw new ApplicationException("Expected '" + expectedErrorMessageType + "' error but got " + CurrentErrorMessage.Key + ": '" + CurrentErrorMessage.Value + "'");
                    }

                    //Got expected Credentials error
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < AddinTestUtility.AuthenticationTimeout);

            if (CurrentErrorMessage.Key != expectedErrorMessageType)
            {
                throw new ApplicationException("Expected '" + expectedErrorMessageType + "' error but got " + CurrentErrorMessage.Key + ": '" + CurrentErrorMessage.Value + "'");
            }
        }

        public void SelectCancelVerifySuccess()
        {
            UIAUtility.PressButton(_cancelCtl);

            UIAUtility.WaitForWindowToDisappearByIdWithTimeout(_parentElement, DialogId, AddinTestUtility.DialogCancelTimeout);
        }

        public CreateAccountDriver SelectCreateAccount()
        {
            UIAUtility.PressButton(_newAccountCtl);

            return CreateAccountDriver.FindFromParent(_parentElement);
        }

        protected override string DialogId { get { return "LoginDialog"; } }

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
