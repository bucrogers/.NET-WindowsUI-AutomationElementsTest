# WinUITests

This Windows UI Automated Test project uses the Windows UIA API (which is what CUIT sits on top of) to test an Excel Addin.

It targets Excel platforms 2007, 2010 and 2013.

## Approach

The tests are all in the AddinUITests TestClass. Each test is succinct and short, most fit in a 
single screen view.

To minimize test brittleness, AutomationIds and UIA invoke / select are used throughout 
(rather than positional mouse clicks etc.).

The Page Object pattern is used to encapsulate controlling and interrogating each of the Addin's 
modal dialogs (including login, selecting a "feed" from a list etc.) - these are 
found in the DialogDrivers folder.

The Wrapper / Dispose pattern is used to ensure that Excel automation instances are cleaned-up properly
without accumulating zombie Excel processes (ExcelAppWrapper).

The AddinRibbonController encapsulates all Ribbon button operations, including invoke and 
interrogating enabled states.

Error messages and types are detected by interrogating visible state of error-specific controls
(rather than matching on text which is brittle), and including any error message in 
the test assert failure message - e.g., LoginDriver.CurrentErrorMessage returns this info.

Timeouts and retries are used throughout, rather than delays, to ensure the tests are both
speedy and reliable / repeatable.