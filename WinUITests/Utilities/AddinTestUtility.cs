using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

using ItemSystemTests.DialogDrivers;

namespace ItemSystemTests.Utilities
{
    public class AddinTestUtility
    {
        public const string DemoUrl = "https://theapp-demo.herokuapp.com/";

        //Credentials
        public const string LoginEmail = "testuser@testdomain.com";
        public const string LoginPsw = "uay3phoLai4fief";
        public const string KnownInvalidLoginEmail = "invaliduser@invaliddomain.com";
        public const string KnownInvalidLoginPassword = "password";

        //Items
        public const string ReadonlyItemName = "Census, Small";
        public const string ModifiableItemName = "UITEST-ModifiableItem";

        //Timeouts
        public static TimeSpan AuthenticationTimeout = TimeSpan.FromSeconds(60);
        public static TimeSpan FindRibbonButtonsTimeout = TimeSpan.FromSeconds(2);
        public static TimeSpan RibbonMouseMoveToClickDelay = TimeSpan.FromSeconds(1);
        public static TimeSpan RibbonMouseMoveToClickDelayAllowForTooltip = TimeSpan.FromSeconds(4);
        public static TimeSpan RibbonButtonsBecomeActivatedTimeout = TimeSpan.FromSeconds(20);
        public static TimeSpan WebServiceResponsePopulateTimeout = TimeSpan.FromSeconds(30);
        public static TimeSpan DialogCancelTimeout = TimeSpan.FromSeconds(5);
        public static TimeSpan WaitForSelectedWebLinkDelay = TimeSpan.FromSeconds(5);
        public static TimeSpan FloorDialogInitDelay = TimeSpan.FromMilliseconds(500);
        public static TimeSpan DialogControlEventStateUpdateTimeout = TimeSpan.FromSeconds(5);
        public static TimeSpan DialogControlEventStateUpdateDelay = TimeSpan.FromSeconds(1);

        public static void SetEnvironmentVars(string serviceUrl)
        {
            //Use non-prod URL by default for all tests in this fixture
            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("TURBO_SPICE_SERVER")))
            {
                Environment.SetEnvironmentVariable("TURBO_SPICE_SERVER", serviceUrl);
            }
        }

        /// <summary>
        /// Open login dialog, first closing a remember-me logged-in scenario if necessary
        /// </summary>
        public static LoginDriver OpenLoginDialog(AddinRibbonController ribbon)
        {
            ribbon.InvokeButton(AddinRibbonButton.Login);

            return LoginDriver.FindFromParent(ribbon.ExcelElement);
        }

        /// <summary>
        /// Login success scenario
        /// </summary>
        /// <remarks>
        /// If already-logged in session exists, preventing login dialog from showing, this will logout and login with 'remember me' unchecked
        /// </remarks>
        /// <param name="deRibbonTabElement"></param>
        /// <param name="loginEmail"></param>
        /// <param name="loginPassword"></param>
        /// <exception cref="ApplicationException">Login failed for any reason</exception>
        public static void LoginExpectSuccess(AddinRibbonController ribbon, string loginEmail, string loginPassword)
        {
            var loginDialog = OpenLoginDialog(ribbon);

            loginDialog.Email = loginEmail;
            loginDialog.Password = loginPassword;
            loginDialog.RememberMe = false;

            loginDialog.SelectOkVerifySuccess();
        }

        public static AutomationElement GetAddInRibbonTabElement(AutomationElement excelElement, ExcelAppWrapper app)
        {
            const string TheAddInTabControlName = "The AddIn";

            var ribbonTabs = UIAUtility.FindElementByNameWithTimeout(excelElement, "Ribbon Tabs", AddinTestUtility.RibbonButtonsBecomeActivatedTimeout,
                    TreeScope.Descendants);

            var deTab = UIAUtility.FindElementByNameWithTimeout(ribbonTabs, TheAddInTabControlName, TimeSpan.FromSeconds(5),
                    TreeScope.Descendants); //Note Descendants needed here only for excel 2007

            if (app.IsVersion2010OrAbove())
            {
                UIAUtility.SelectMenu(deTab);
            }
            else
            {
                UIAUtility.PressButton(deTab);
            }

            var lowerRibbon = UIAUtility.FindElementByNameWithTimeout(excelElement, "Lower Ribbon", AddinTestUtility.RibbonButtonsBecomeActivatedTimeout,
                    TreeScope.Descendants);

            return UIAUtility.FindElementByNameWithTimeout(lowerRibbon, TheAddInTabControlName, AddinTestUtility.RibbonButtonsBecomeActivatedTimeout,
                    TreeScope.Descendants); //Note Descendants needed here only for excel 2007
        }

        public static string GetUniqueSuffix()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmssfff");
        }
    }
}
