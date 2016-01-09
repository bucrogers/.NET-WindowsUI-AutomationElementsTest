using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;

using Microsoft.Test.Input;

using ItemSystemTests.Utilities;
using ItemSystemTests.DialogDrivers;

namespace ItemSystemTests
{
    [TestClass]
    public class AddinUITests
    {
        private static ExcelAppWrapper _app;
        private ExcelWorkbookWrapper _wb;
        private static dynamic _facade;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            AddinTestUtility.SetEnvironmentVars(AddinTestUtility.DemoUrl);

            if (Debugger.IsAttached) //use excel-per-instance only when running from VS - blows up running automated from console
            {
                _app = ExcelAutoUtility.OpenAndInstallAddin("TheAddIn", out _facade);
            }
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            if (Debugger.IsAttached) //use excel-per-instance only when running from VS - blows up running automated from console
            {
                if (_app != null)
                {
                    _app.Dispose();
                }
            }
        }

        [TestInitialize]
        public void Initialize()
        {
            if (!Debugger.IsAttached) //use excel-per-instance only when running from VS - blows up running automated from console
            {
                _app = ExcelAutoUtility.OpenAndInstallAddin("TheAddIn", out _facade);
            }
            _wb = new ExcelWorkbookWrapper(_app.App.Workbooks.Add());

            //Ensure logged out at start
            _facade.Logout();
        }

        [TestCleanup]
        public void Cleanup()
        {
            try
            {
                //Close modal dialog if any (to avoid excel addin corruption when process later killed)
                if (_wb != null && _wb.Wb != null)
                {
                    var excel = ExcelAutoUtility.GetAppElementFromApp(_wb.Wb.Application);
                    UIAUtility.FindModalDialogIfAnyAndClose(excel);

                    _wb.Dispose();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error closing workbook (will fall through to excel app kill): " + ex.Message);
            }

            if (!Debugger.IsAttached) //use excel-per-instance only when running from VS - blows up running automated from console
            {
                if (_app != null)
                {
                    _app.Dispose();
                }
            }
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void AboutDialog_CloseViaOKButton_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Open About dialog
            ribbon.InvokeButton(AddinRibbonButton.About);

            var aboutDlg = AboutDriver.FindFromParent(ribbon.ExcelElement);
            aboutDlg.SelectOKVerifySuccess();
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void AboutDialog_CloseViaXIcon_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Open About dialog
            ribbon.InvokeButton(AddinRibbonButton.About);

            var aboutDlg = AboutDriver.FindFromParent(ribbon.ExcelElement);
            aboutDlg.SelectCloseIconVerifySuccess();
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void AboutDialog_CloseViaEscKey_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Open About dialog
            ribbon.InvokeButton(AddinRibbonButton.About);

            var aboutDlg = AboutDriver.FindFromParent(ribbon.ExcelElement);
            aboutDlg.SelectEscVerifySuccess();
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void CreateAccountDialog_CreateUnconfirmedAccount_FailsLogin()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Open Login dialog
            var loginDialog = AddinTestUtility.OpenLoginDialog(ribbon);

            //Open Create Account dialog and create a unique one
            var createAccountDialog = loginDialog.SelectCreateAccount();
            createAccountDialog.Name = "non unique test account name";
            var uniqueAccountName = "testacct" + AddinTestUtility.GetUniqueSuffix() + "@testdomain.com";
            createAccountDialog.Email = uniqueAccountName;
            var accountPassword = "password";
            createAccountDialog.Password = accountPassword;
            
            //Ensure terms is required
            Assert.IsFalse(createAccountDialog.RegisterEnabled);

            //Accept terms then register
            createAccountDialog.Terms = true;

            var messageBoxDialog = createAccountDialog.SelectRegister();

            messageBoxDialog.SelectOKVerifySuccess();

            //Attempt to login with above unconfirmed account, expect failure
            loginDialog = LoginDriver.FindFromParent(ribbon.ExcelElement);
            loginDialog.Email = uniqueAccountName;
            loginDialog.Password = accountPassword;

            loginDialog.SelectOkVerifyFailure(LoginDriver.ErrorMessageType.Credentials);

            //Close login dialog
            loginDialog.SelectCancelVerifySuccess();
        }

        /// <summary>
        /// Basic happy-path end-to-end insert scenario of existing item
        /// </summary>
        [TestMethod]
        [TestCategory("Integration")]
        public void CreateItemDialog_CreateSecondItemSameSheet_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login
            AddinTestUtility.LoginExpectSuccess(ribbon, AddinTestUtility.LoginEmail, AddinTestUtility.LoginPsw);

            //Insert item
            ribbon.InvokeButton(AddinRibbonButton.InsertItem);

            var itemDlg = InsertItemDriver.FindFromParent(ribbon.ExcelElement);
            itemDlg.SelectedItem = AddinTestUtility.ReadonlyItemName;
            itemDlg.AllowPublish = true;
            itemDlg.SelectInsertVerifySuccess();

            //Wait for item to be inserted in item-named worksheet
            var newlyCreatedItemSheet = ExcelAutoUtility.GetWorkSheetWithTimeout(_wb.Wb, AddinTestUtility.ReadonlyItemName, AddinTestUtility.WebServiceResponsePopulateTimeout);

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnPublishableItemCell));

            ExcelAutoUtility.SetCellValue(newlyCreatedItemSheet, 2, 48, "singleCellValue");
            var range = newlyCreatedItemSheet.get_Range("AV2", "AV2");
            range.Select();

            //Force ribbon refresh in this special case of selecting range programmatically via excel automation
            //which does not result in ribbon refresh
            ribbon.ForceRibbonRefresh();

            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnExistingAllowPublishItemSheetOffItemCell));

            //Create the second item
            var newItemName = "uitest" + AddinTestUtility.GetUniqueSuffix();
            var newItemDescr = "sample description #2";

            ribbon.Refresh();
            ribbon.InvokeButton(AddinRibbonButton.Create);

            var createItemDlg = CreateItemDriver.FindFromParent(ribbon.ExcelElement);
            createItemDlg.Name = newItemName;
            createItemDlg.Desc = newItemDescr;

            createItemDlg.SelectCreateVerifySuccess();

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnPublishableItemCell));
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void CreateItemDialog_CreateSimpleItem_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            var newItemName = "uitest" + AddinTestUtility.GetUniqueSuffix();
            var newItemDescr = "sample description";

            CreateItemDialogCreateSimpleItem(ribbon, newItemName, newItemDescr);

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnPublishableItemCell));

            //Validate item got created by finding in Manage Items

            //Get ribbon control again after enabled states changed from above
            ribbon.Refresh();

            ribbon.InvokeButton(AddinRibbonButton.ManageItems);

            var manageItemsDlg = ManageItemsDriver.FindFromParent(ribbon.ExcelElement);
            manageItemsDlg.SelectedItem = newItemName; //no ex means success

            manageItemsDlg.SelectCloseButtonVerifySuccess();
        }

        private void CreateItemDialogCreateSimpleItem(AddinRibbonController ribbon, string newItemName, string newItemDescr)
        {
            //Login
            AddinTestUtility.LoginExpectSuccess(ribbon, AddinTestUtility.LoginEmail, AddinTestUtility.LoginPsw);

            //Create a data range from which to create item
            var sheet = _wb.Wb.ActiveSheet as Excel.Worksheet;

            ExcelAutoUtility.SetCellValue(sheet, 1, 1, "colA");
            ExcelAutoUtility.SetCellValue(sheet, 1, 2, "colB");
            ExcelAutoUtility.SetCellValue(sheet, 2, 1, 1);
            ExcelAutoUtility.SetCellValue(sheet, 2, 2, 2);
            ExcelAutoUtility.SetCellValue(sheet, 3, 1, 3);
            ExcelAutoUtility.SetCellValue(sheet, 3, 2, 4);

            var range = sheet.get_Range("A1", "B3");
            range.Select();

            //Create item
            ribbon.InvokeButton(AddinRibbonButton.Create);

            var createItemDlg = CreateItemDriver.FindFromParent(ribbon.ExcelElement);
            createItemDlg.Name = newItemName;
            createItemDlg.Desc = newItemDescr;

            createItemDlg.SelectCreateVerifySuccess();
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void EditDetailsDialog_ModifyNameAndDescription_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            var newItemName = "uitest" + DateTime.Now.ToString("yyyyMMddHHmmssfff");
            var newItemDescr = "sample description";

            CreateItemDialogCreateSimpleItem(ribbon, newItemName, newItemDescr);

            //Test edit item details

            string modName = newItemName + "-mod";
            string modDesc = newItemDescr + "-mod";

            //Modify
            {
                //Get ribbon control again after enabled states changed from above
                ribbon.Refresh();

                ribbon.InvokeButton(AddinRibbonButton.EditDetails);

                var editDetailsDlg = EditDetailsDriver.FindFromParent(ribbon.ExcelElement);

                //Validate start values
                Assert.AreEqual(newItemName, editDetailsDlg.Name);
                Assert.AreEqual(newItemDescr, editDetailsDlg.Desc);

                //Modify
                editDetailsDlg.Name = modName;
                editDetailsDlg.Desc = modDesc;
                editDetailsDlg.SelectOKVerifySuccess();
            }

            //Verify mods
            {
                ribbon.InvokeButton(AddinRibbonButton.EditDetails);

                var editDetailsDlg = EditDetailsDriver.FindFromParent(ribbon.ExcelElement);

                //Validate
                Assert.AreEqual(modName, editDetailsDlg.Name);
                Assert.AreEqual(modDesc, editDetailsDlg.Desc);
            }
        }

        /// <summary>
        /// Basic happy-path end-to-end insert scenario of existing item
        /// </summary>
        [TestMethod]
        [TestCategory("Integration")]
        public void InsertItemDialog_InsertViaOK_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login
            AddinTestUtility.LoginExpectSuccess(ribbon, AddinTestUtility.LoginEmail, AddinTestUtility.LoginPsw);

            //Insert item
            ribbon.InvokeButton(AddinRibbonButton.InsertItem);

            var itemDlg = InsertItemDriver.FindFromParent(ribbon.ExcelElement);
            itemDlg.SelectedItem = AddinTestUtility.ReadonlyItemName;
            itemDlg.SelectInsertVerifySuccess();

            //Wait for item to be inserted in item-named worksheet
            var newlyCreatedItemSheet = ExcelAutoUtility.GetWorkSheetWithTimeout(_wb.Wb, AddinTestUtility.ReadonlyItemName, AddinTestUtility.WebServiceResponsePopulateTimeout);

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnNonPublishableItemCell));

            //Validate cell content from inserted item
            Assert.AreEqual("Sumlev", ExcelAutoUtility.GetCellValue(newlyCreatedItemSheet, 1, 1));
            Assert.AreEqual(-10.42333312, ExcelAutoUtility.GetCellValue(newlyCreatedItemSheet, 58, 46));
        }

        /// <summary>
        /// Invoke help dialog. Implementation note: Currently a simple fire-and-forget 
        /// - Success defined as: Lack of error here, along with visual observation of appropriate link coming  up
        /// </summary>
        [TestMethod]
        [TestCategory("Integration")]
        public void HelpDialog_Select_BrowserPageLaunched()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Open Help link
            ribbon.InvokeButton(AddinRibbonButton.Help);

            Thread.Sleep(AddinTestUtility.WaitForSelectedWebLinkDelay);
        }

        /// <summary>
        /// Invoke history dialog. Implementation note: Currently a simple fire-and-forget 
        /// - Success defined as: Lack of error here, along with visual observation of appropriate link coming  up
        /// </summary>
        [TestMethod]
        [TestCategory("Integration")]
        public void HistoryDialog_Select_BrowserPageLaunched()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login
            AddinTestUtility.LoginExpectSuccess(ribbon, AddinTestUtility.LoginEmail, AddinTestUtility.LoginPsw);

            //Insert item
            ribbon.InvokeButton(AddinRibbonButton.InsertItem);

            var itemDlg = InsertItemDriver.FindFromParent(ribbon.ExcelElement);
            itemDlg.SelectedItem = AddinTestUtility.ReadonlyItemName;
            itemDlg.SelectInsertVerifySuccess();

            //Buttons including History now enabled
            //Open History link

            //Get ribbon tab automation element again to get its newly-enabled instances of descendant buttons
            ribbon.Refresh();

            ribbon.InvokeButton(AddinRibbonButton.History);

            Thread.Sleep(AddinTestUtility.WaitForSelectedWebLinkDelay);
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void RibbonBar_StartupLoggedIn_SelectedItemButtonsEnabled()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login
            AddinTestUtility.LoginExpectSuccess(ribbon, AddinTestUtility.LoginEmail, AddinTestUtility.LoginPsw);

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.LoggedInButtonStates));
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void RibbonBar_StartupLoggedOut_ItemButtonsDisabled()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.NotLoggedInButtonStates));
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void LoginDialog_InvalidCredentials_ExpectedError()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Attempt login with invalid credential
            var loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            loginDlg.Email = AddinTestUtility.KnownInvalidLoginEmail;
            loginDlg.Password = AddinTestUtility.KnownInvalidLoginPassword;
            loginDlg.RememberMe = false;
            loginDlg.SelectOkVerifyFailure(LoginDriver.ErrorMessageType.Credentials);

            //Attempt login with valid credential (after the failed attempt above)
            loginDlg.Email = AddinTestUtility.LoginEmail;
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = false;
            loginDlg.SelectOkVerifySuccess();
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void LoginDialog_RememberMeSelectedLogout_EmailOnlyRemembered()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login with Remember Me selected
            var loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            loginDlg.Email = AddinTestUtility.LoginEmail;
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = true;
            loginDlg.SelectOkVerifySuccess();

            //Logout
            ribbon.InvokeButton(AddinRibbonButton.Logout);
            ribbon.ValidateSingleButtonState(AddinRibbonButton.Logout, RibbonButtonState.NonVisible);

            //Open login and expect only email is remembered
            ribbon.Refresh();
            ribbon.InvokeButton(AddinRibbonButton.Login);

            loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            Assert.AreEqual(AddinTestUtility.LoginEmail, loginDlg.Email);
            Assert.AreEqual(string.Empty, loginDlg.Password);
            Assert.IsTrue(loginDlg.RememberMe);

            //End test logged-out / remember-me deselected
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = false;
            loginDlg.SelectOkVerifySuccess();

            ribbon.InvokeButton(AddinRibbonButton.Logout);
            ribbon.ValidateSingleButtonState(AddinRibbonButton.Logout, RibbonButtonState.NonVisible);
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void LoginDialog_RememberMeDeselected_NotLoggedInOnStartup()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login and select Remember Me
            var loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            loginDlg.Email = AddinTestUtility.LoginEmail;
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = false;
            loginDlg.SelectOkVerifySuccess();

            //Close excel
            _wb.Dispose();
            _app.Dispose();

            //Re-open excel
            _app = ExcelAutoUtility.OpenAndInstallAddin("TheAddIn", out _facade);
            _wb = new ExcelWorkbookWrapper(_app.App.Workbooks.Add());

            //Validate not auto-logged-in
            ribbon = AddinRibbonController.Initialize(_app);

            Assert.AreEqual(RibbonButtonState.Enabled, ribbon.GetButtonState(AddinRibbonButton.Login));
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void LoginDialog_RememberMeSelected_LoggedInOnStartup()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login and select Remember Me
            var loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            loginDlg.Email = AddinTestUtility.LoginEmail;
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = true;
            loginDlg.SelectOkVerifySuccess();

            //Close excel
            _wb.Dispose();
            _app.Dispose();

            //Re-open excel
            _app = ExcelAutoUtility.OpenAndInstallAddin("TheAddIn", out _facade);
            _wb = new ExcelWorkbookWrapper(_app.App.Workbooks.Add());

            //Validate auto-logged-in
            ribbon = AddinRibbonController.Initialize(_app,
                    true); //disableWaitForRibbonReady

            ribbon.PostInitializeSpecialCaseAutologgedIn();

            ribbon.ValidateSingleButtonState(AddinRibbonButton.Logout, RibbonButtonState.Enabled, AddinTestUtility.AuthenticationTimeout);

            //Leave test in a logged-out / non-remember-me state
            ribbon.InvokeButton(AddinRibbonButton.Logout);
            ribbon.ValidateSingleButtonState(AddinRibbonButton.Logout, RibbonButtonState.NonVisible);

            loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            loginDlg.Email = AddinTestUtility.LoginEmail;
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = false;
            loginDlg.SelectOkVerifySuccess();

            ribbon.InvokeButton(AddinRibbonButton.Logout);
            ribbon.ValidateSingleButtonState(AddinRibbonButton.Logout, RibbonButtonState.NonVisible);
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void LoginDialog_RememberMeSelectedLogout_RememberEmailOnlyOnStartup()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Login and select Remember Me
            var loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);
            loginDlg.Email = AddinTestUtility.LoginEmail;
            loginDlg.Password = AddinTestUtility.LoginPsw;
            loginDlg.RememberMe = true;
            loginDlg.SelectOkVerifySuccess();

            //Logout
            ribbon.InvokeButton(AddinRibbonButton.Logout);
            ribbon.ValidateSingleButtonState(AddinRibbonButton.Logout, RibbonButtonState.NonVisible);

            //Close excel
            _wb.Dispose();
            _app.Dispose();

            //Re-open excel
            _app = ExcelAutoUtility.OpenAndInstallAddin("TheAddIn", out _facade);
            _wb = new ExcelWorkbookWrapper(_app.App.Workbooks.Add());

            ribbon = AddinRibbonController.Initialize(_app);

            //Validate not auto-logged-in, but email addr remembered
            loginDlg = AddinTestUtility.OpenLoginDialog(ribbon);

            Assert.AreEqual(AddinTestUtility.LoginEmail, loginDlg.Email);
            Assert.AreEqual(string.Empty, loginDlg.Password);

            //Close dialog
            loginDlg.SelectCancelVerifySuccess();
        }


        [TestMethod]
        [TestCategory("Integration")]
        public void PublishDialog_ViaPublishItems_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Insert an item to modify
            var ws = InsertOKToModifyItemData(ribbon, AddinTestUtility.ModifiableItemName,
                    false); //getSecondInstance

            //Modify a cell
            var modValRow = 3;
            var modValCol = 2;
            int newVal = Convert.ToInt32(ExcelAutoUtility.GetCellValue(ws, modValRow, modValCol)) + 1;
            ExcelAutoUtility.SetCellValue(ws, modValRow, modValCol, newVal);

            //Publish
            
            //Get ribbon tab automation element again to get its newly-enabled instances of descendant buttons
            ribbon.Refresh();

            ribbon.InvokeButton(AddinRibbonButton.PublishItems);

            var pubDlg = PublishDriver.FindFromParent(ribbon.ExcelElement);
            pubDlg.SelectPublishVerifySuccess();

            //Verify
            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnPublishableItemCell));

            ws = InsertOKToModifyItemData(ribbon, AddinTestUtility.ModifiableItemName,
                    true); //getSecondInstance

            Assert.AreEqual(newVal, Convert.ToInt32(ExcelAutoUtility.GetCellValue(ws, modValRow, modValCol)));
        }

        [TestMethod]
        [TestCategory("Integration")]
        public void PublishDialog_ViaPublish_Success()
        {
            var ribbon = AddinRibbonController.Initialize(_app);

            //Insert an item to modify
            var ws = InsertOKToModifyItemData(ribbon, AddinTestUtility.ModifiableItemName,
                    false); //getSecondInstance

            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnPublishableItemCell));

            //Modify a cell
            var modValRow = 3;
            var modValCol = 2;
            int newVal = Convert.ToInt32(ExcelAutoUtility.GetCellValue(ws, modValRow, modValCol)) + 1;
            ExcelAutoUtility.SetCellValue(ws, modValRow, modValCol, newVal);

            //Publish

            //Get ribbon tab automation element again to get its newly-enabled instances of descendant buttons
            ribbon.Refresh();

            ribbon.InvokeButton(AddinRibbonButton.Publish);

            var pubDlg = PublishDriver.FindFromParent(ribbon.ExcelElement);
            pubDlg.SelectPublishVerifySuccess();

            //Verify
            ribbon.Refresh();
            Assert.AreEqual(string.Empty, ribbon.ValidateExpectedButtonStates(ExpectedButtonStates.OnPublishableItemCell));

            ws = InsertOKToModifyItemData(ribbon, AddinTestUtility.ModifiableItemName,
                    true); //getSecondInstance

            Assert.AreEqual(newVal, Convert.ToInt32(ExcelAutoUtility.GetCellValue(ws, modValRow, modValCol)));
        }

        private Excel.Worksheet InsertOKToModifyItemData(AddinRibbonController ribbon, string itemName, bool getSecondInstance)
        {
            //Login
            if ( ! getSecondInstance)
            {
                AddinTestUtility.LoginExpectSuccess(ribbon, AddinTestUtility.LoginEmail, AddinTestUtility.LoginPsw);
            }

            //Insert item
            ribbon.InvokeButton(AddinRibbonButton.InsertItem);

            var itemDlg = InsertItemDriver.FindFromParent(ribbon.ExcelElement);
            itemDlg.SelectedItem = itemName;
            itemDlg.AllowPublish = true;
            itemDlg.SelectInsertVerifySuccess();

            //Wait for item to be inserted in item-named worksheet
            var sheetName = itemName;
            if (getSecondInstance)
            {
                sheetName += " 1";
            }
            var ws = ExcelAutoUtility.GetWorkSheetWithTimeout(_wb.Wb, sheetName, AddinTestUtility.WebServiceResponsePopulateTimeout);

            //Wait for access to worksheet to be stable
            ExcelAutoUtility.WaitForNewWorksheetToBeAccessible(ws);

            return ws;
        }
    }
}
