using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;

using ItemSystemTests.DialogDrivers;

namespace ItemSystemTests.Utilities
{
    public class AddinRibbonController
    {
        private Dictionary<AddinRibbonButton, string> _buttonNameLookup = new Dictionary<AddinRibbonButton, string>() 
        { 
            {AddinRibbonButton.About, "About"},
            {AddinRibbonButton.Create, "Create"},
            {AddinRibbonButton.EditDetails, "Edit Details"},
            {AddinRibbonButton.Help, "Help"},
            {AddinRibbonButton.History, "History"},
            {AddinRibbonButton.InsertItem, "Insert item"},
            {AddinRibbonButton.Login, "Login"},
            {AddinRibbonButton.Logout, "Logout"},
            {AddinRibbonButton.ManageItems, "Manage Items"},
            {AddinRibbonButton.Publish, "Publish"},
            {AddinRibbonButton.PublishItems, "Publish Items"},
            {AddinRibbonButton.Refresh, "Refresh"},
            {AddinRibbonButton.RefreshItems, "Refresh Items"}
        };

        private ExcelAppWrapper _app;
        private AutomationElement _excel;
        private AutomationElement _addinRibbonCtl;

        private AddinRibbonController(ExcelAppWrapper app, bool disableWaitForRibbonReady)
        {
            _app = app;
            _excel = ExcelAutoUtility.GetAppElementFromApp(_app.App.Application);

            _addinRibbonCtl = AddinTestUtility.GetAddInRibbonTabElement(_excel, app);

            if ( ! disableWaitForRibbonReady)
            {
                WaitForRibbonReady();
            }
        }

        public static AddinRibbonController Initialize(ExcelAppWrapper app)
        {
            return Initialize(app,
                    false); //disableWaitForRibbonReady
        }

        public static AddinRibbonController Initialize(ExcelAppWrapper app, bool disableWaitForRibbonReady)
        {
            return new AddinRibbonController(app, disableWaitForRibbonReady);
        }

        public AutomationElement ExcelElement
        {
            get
            {
                return _excel;
            }
        }

        public void PostInitializeSpecialCaseAutologgedIn()
        {
            //With Excel 2010 and earlier, the below is necessary to refresh the ribbon bar buttons
            if (!_app.IsVersion2013OrAbove())
            {
                InvokeButton(AddinRibbonButton.About);

                var aboutDlg = AboutDriver.FindFromParent(_excel);
                aboutDlg.SelectOKVerifySuccess();
            }
        }

        public void InvokeButton(AddinRibbonButton button)
        {
            if ( (button == AddinRibbonButton.Create && ! _app.IsVersion2010OrAbove() ) //2007 create button does not support invoke pattern
                || ! WindowsUtility.WindowsVerSupportsModalInvokeWithoutHang() )
            {
                //Special-cases which lead to blocking modal dialog when using InvokePattern - must instead use mouse click
                UIAUtility.FindElementByNameFilteredByControlTypeAndMouseClick(_addinRibbonCtl,
                        _buttonNameLookup[button], ControlType.Button, ControlType.Custom, AddinTestUtility.RibbonMouseMoveToClickDelayAllowForTooltip,
                        TreeScope.Descendants);
            }
            else
            {
                var el = UIAUtility.FindElementByControlTypeAndNameWithTimeout(_addinRibbonCtl, ControlType.Button, ControlType.Custom,
                        _buttonNameLookup[button], 
                        AddinTestUtility.FindRibbonButtonsTimeout, TreeScope.Descendants);

                UIAUtility.WaitForElementEnabledWithTimeout(el, AddinTestUtility.RibbonButtonsBecomeActivatedTimeout);

                UIAUtility.PressButton(el);

                if (button == AddinRibbonButton.Logout)
                {
                    //Special case force ribbon refresh when logout button invoked
                    ForceRibbonRefresh();
                }
            }
        }

        public string GetButtonName(AddinRibbonButton button)
        {
            return _buttonNameLookup[button];
        }

        public RibbonButtonState GetButtonState(AddinRibbonButton button)
        {
            RefreshRibbonPreRead();

            return GetButtonStateRaw(button);
        }

        private RibbonButtonState GetButtonStateRaw(AddinRibbonButton button)
        {
            var el = _addinRibbonCtl.FindFirst(TreeScope.Descendants,
                    new AndCondition(
                            new OrCondition(
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)),
                            new PropertyCondition(AutomationElement.NameProperty, _buttonNameLookup[button])));

            if (el == null)
            {
                return RibbonButtonState.NonVisible;
            }
            if (el.Current.IsEnabled)
            {
                if (el.Current.IsOffscreen)
                {
                    return RibbonButtonState.EnabledNotVisible;
                }

                return RibbonButtonState.Enabled;
            }
            if (el.Current.IsOffscreen)
            {
                return RibbonButtonState.NonVisible;
            }

            return RibbonButtonState.VisibleNotEnabled;
        }

        /// <summary>
        /// Refresh ribbon AutomationElement - required when app-under-test re-instantiates its controls
        /// </summary>
        public void Refresh()
        {
            //TODO: Determine if this is still necessary, now that improved downstream retries are in place
            _addinRibbonCtl = AddinTestUtility.GetAddInRibbonTabElement(_excel, _app);
        }

        public void ValidateSingleButtonState(AddinRibbonButton button, RibbonButtonState expectedState)
        {
            ValidateSingleButtonState(button, expectedState, AddinTestUtility.RibbonButtonsBecomeActivatedTimeout);
        }

        /// <summary>
        /// Determine if specified button states match, using timeout to allow for event updating
        /// </summary>
        /// <param name="button"></param>
        /// <param name="expectedState"></param>
        /// <param name="timeOutToUse"></param>
        /// <exception cref="">if state still does not match after timeout</exception>
        public void ValidateSingleButtonState(AddinRibbonButton button, RibbonButtonState expectedState, TimeSpan timeOutToUse)
        {
            RefreshRibbonPreRead();

            var sw = Stopwatch.StartNew();
            do
            {
                if (GetButtonStateRaw(button) == expectedState)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeOutToUse);

            throw new ApplicationException("Button '" + button + "' did not get to expected state '" + expectedState + "' within " + AddinTestUtility.DialogControlEventStateUpdateTimeout.TotalSeconds + "sec");
        }

        /// <summary>
        /// Determine if expected button states match, using timeout to allow for event updating
        /// </summary>
        /// <param name="expectedStates"></param>
        /// <returns>string.Empty if ok - otherwise newline-delimited list of diffs</returns>
        public string ValidateExpectedButtonStates(SortedDictionary<AddinRibbonButton, RibbonButtonState> expectedStates)
        {
            RefreshRibbonPreRead();

            var sw = Stopwatch.StartNew();
            StringBuilder mismatches = null;
            do
            {
                mismatches = new StringBuilder();
                foreach (var pair in expectedStates)
                {
                    if (GetButtonStateRaw(pair.Key) != pair.Value)
                    {
                        mismatches.Append("Button " + pair.Key + " expected '" + pair.Value + " but got " + GetButtonStateRaw(pair.Key) + Environment.NewLine);
                    }
                }

                if (mismatches.Length == 0)
                {
                    //No mismatches, done
                    return string.Empty;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < AddinTestUtility.DialogControlEventStateUpdateTimeout);

            //have at least one mismatch
            return mismatches.ToString();
        }

        private void RefreshRibbonPreRead()
        {
            //Nop at present - probably can be eliminated after proving the downstream retry timeouts are doing the job
        }

        /// <summary>
        /// Forces ribbon refresh by clicking about then closing the modal dialog
        /// </summary>
        /// <remarks>
        /// Should be used sparingly, only when necessary to force refresh under automation
        /// </remarks>
        public void ForceRibbonRefresh()
        {
            if (WindowsUtility.WindowsVerSupportsModalInvokeWithoutHang())
            {
                //Ensure event updating of ribbon in excel by invoking a dialog then closing it
                InvokeButton(AddinRibbonButton.About);
                var aboutDlg = AboutDriver.FindFromParent(_excel);
                aboutDlg.SelectOKVerifySuccess();
            }
            else
            {
                const int retries = 2;
                for (int i = 0; i < retries; i++)
                {
                    try
                    {
                        //Ensure event updating of ribbon in excel by invoking a dialog then closing it
                        InvokeButton(AddinRibbonButton.About);
                        var aboutDlg = AboutDriver.FindFromParent(_excel);
                        aboutDlg.SelectOKVerifySuccess();

                        break;
                    }
                    catch (Exception)
                    {
                        //Go to retries
                    }
                }
            }
        }

        private void WaitForRibbonReady()
        {
            //Wait for Login to show as enabled, as indication ribbon is in ready state
            var sw = Stopwatch.StartNew();
            do
            {
                if (GetButtonStateRaw(AddinRibbonButton.Login) == RibbonButtonState.Enabled)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < AddinTestUtility.RibbonButtonsBecomeActivatedTimeout);

            //Force ribbon refresh, for the scenario where an initialize facade logout is not yet "seen"
            ForceRibbonRefresh();

            ValidateSingleButtonState(AddinRibbonButton.Login, RibbonButtonState.Enabled);
        }
    }
}
