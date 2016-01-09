using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Diagnostics;

using Microsoft.Test.Input;

namespace ItemSystemTests.Utilities
{
    public class UIAUtility
    {
        public static void WaitForPopulatedDatagrid(GridPattern grid, TimeSpan timeout)
        {
            var stableTimeToConsiderPopulated = TimeSpan.FromSeconds(1);
            int lastRowCount = 0;
            Stopwatch swStable = new Stopwatch();
            var sw = Stopwatch.StartNew();
            do
            {
                if (grid.Current.RowCount > 0)
                {
                    if (swStable.IsRunning)
                    {
                        if (grid.Current.RowCount != lastRowCount)
                        {
                            swStable.Restart();
                        }
                        else if (swStable.Elapsed >= stableTimeToConsiderPopulated)
                        {
                            return;
                        }
                    }
                    else
                    {
                        swStable.Start();
                    }

                    lastRowCount = grid.Current.RowCount;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out waiting for populated datagrid after " + timeout.TotalSeconds + "sec");
        }

        public static void FindElementByNameFilteredByControlTypeAndMouseClick(AutomationElement parent, string automationName, ControlType controlType, ControlType controlType2, TimeSpan mouseClickDelay)
        {
            FindElementByNameFilteredByControlTypeAndMouseClick(parent, automationName, controlType, controlType2, mouseClickDelay, TreeScope.Children);
        }

        /// <summary>
        /// Get automationelement by name, filtered by classname - mouse click if found
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="automationName">case-insensitive automation name</param>
        /// <param name="controlType"></param>
        /// <param name="controlType2"></param>
        /// <exception cref="ApplicationException">if matching element not found</exception>
        /// <exception cref="ApplicationException">if specified element is not enabled</exception>
        public static void FindElementByNameFilteredByControlTypeAndMouseClick(AutomationElement parent, string automationName, ControlType controlType, ControlType controlType2, TimeSpan mouseClickDelay, TreeScope treeScope)
        {
            FindElementByNameFilteredByControlTypeWithTimeoutAndMouseClick(parent, automationName, controlType, controlType2,
                    AddinTestUtility.FindRibbonButtonsTimeout, //findDelay
                    mouseClickDelay,
                    treeScope);
        }

        public static void FindElementByNameFilteredByControlTypeWithTimeoutAndMouseClick(AutomationElement parent, string automationName, ControlType controlType, ControlType controlType2, TimeSpan findDelay, TimeSpan mouseClickDelay)
        {
            FindElementByNameFilteredByControlTypeWithTimeoutAndMouseClick(parent, automationName, controlType, controlType2, findDelay, mouseClickDelay,
                    TreeScope.Children);
        }

        /// <summary>
        /// Get automationelement by name, filtered by controltype - mouse click if found
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="automationName">case-insensitive automation name</param>
        /// <param name="controlType"></param>
        /// <param name="controlType2"></param>
        /// <param name="findDelay"></param>
        /// <param name="mouseClickDelay"></param>
        /// <param name="treeScope"></param>
        /// <exception cref="ApplicationException">if matching element not found</exception>
        public static void FindElementByNameFilteredByControlTypeWithTimeoutAndMouseClick(AutomationElement parent, string automationName, ControlType controlType, ControlType controlType2, TimeSpan findDelay, TimeSpan mouseClickDelay, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            bool foundButIsDisabled = false;
            bool foundButIsNotVisible = false;
            do
            {
                foreach (AutomationElement foundEl in parent.FindAll(treeScope, new PropertyCondition(AutomationElement.NameProperty,
                        automationName, PropertyConditionFlags.IgnoreCase)))
                {
                    if ( (ControlType) foundEl.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty) == controlType
                            || ((ControlType)foundEl.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty) == controlType2)) 
                    {
                        if ( ! foundEl.Current.IsEnabled)
                        {
                            foundButIsDisabled = true;
                        }
                        else if (foundEl.Current.IsOffscreen)
                        {
                            foundButIsNotVisible = true;
                        }
                        else
                        {
                            UIAUtility.MouseClickInferPoint(foundEl, mouseClickDelay);

                            return;
                        }
                    }
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < findDelay);

            if (foundButIsDisabled)
            {
                throw new ApplicationException("Specified element named '" + automationName + "' with controltype '" + controlType + "': IsEnabled=false");
            }

            if (foundButIsNotVisible)
            {
                throw new ApplicationException("Specified element named '" + automationName + "' with controltype '" + controlType + "': IsOffscreen=true");
            }

            throw new ApplicationException("Could not find element named '" + automationName + "' with controltype '" + controlType + "'");
        }

        public static AutomationElement FindModalDialogWithTimeout(AutomationElement parent, TimeSpan timeout)
        {
            return FindModalDialogWithTimeout(parent, timeout,
                    TreeScope.Children);
        }

        public static AutomationElement FindModalDialogWithTimeout(AutomationElement parent, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                foreach (AutomationElement foundEl in parent.FindAll(treeScope, new PropertyCondition(AutomationElement.IsEnabledProperty, true)))
                {
                    object windowPattern;
                    if (foundEl.TryGetCurrentPattern(WindowPattern.Pattern, out windowPattern))
                    {
                        if (((WindowPattern)windowPattern).Current.IsModal)
                        {
                            return foundEl;
                        }
                    }
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new ApplicationException("Did not find a modal dialog under element '" + GetIdOrName(parent) + "'");
        }

        public static void WaitForModalDialogToDisappearWithTimeout(AutomationElement parent, TimeSpan timeout)
        {
            WaitForModalDialogToDisappearWithTimeout(parent, timeout,
                    TreeScope.Children);
        }

        public static void WaitForModalDialogToDisappearWithTimeout(AutomationElement parent, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            AutomationElement modalDialog = null;
            do
            {
                foreach (AutomationElement foundEl in parent.FindAll(treeScope, new PropertyCondition(AutomationElement.IsEnabledProperty, true)))
                {
                    object windowPattern;
                    if (foundEl.TryGetCurrentPattern(WindowPattern.Pattern, out windowPattern))
                    {
                        if (((WindowPattern)windowPattern).Current.IsModal)
                        {
                            modalDialog = foundEl;
                        }
                    }
                }

                if (modalDialog == null)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new ApplicationException("Modal dialog '" + GetIdOrName(modalDialog) + "' did not disappear within " + timeout.TotalSeconds + "sec");
        }

        public static void FindModalDialogIfAnyAndClose(AutomationElement parent)
        {
            FindModalDialogIfAnyAndClose(parent,
                    TreeScope.Children);
        }

        public static void FindModalDialogIfAnyAndClose(AutomationElement parent, TreeScope treeScope)
        {
            foreach (AutomationElement foundEl in parent.FindAll(treeScope, new PropertyCondition(AutomationElement.IsEnabledProperty, true)))
            {
                object windowPattern;
                if (foundEl.TryGetCurrentPattern(WindowPattern.Pattern, out windowPattern))
                {
                    if ( ((WindowPattern)windowPattern).Current.IsModal)
                    {
                        var closeIconCtl = foundEl.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty,
                                "Close"));
                        if (closeIconCtl == null)
                        {
                            throw new ApplicationException("Modal dialog '" + foundEl.Current.Name + "' does not implement close icon having automation id 'Close'");
                        }

                        UIAUtility.PressButton(closeIconCtl);

                        //Done since there can be only one modal dialog
                        return;
                    }
                }
            }
        }

        public static AutomationElement FindElementByName(AutomationElement root, string automationName)
        {
            return FindElementByName(root, automationName, TreeScope.Children);
        }

        /// <summary>
        /// Find first element by name, throwing if not found
        /// </summary>
        /// <param name="root"></param>
        /// <param name="automationName"></param>
        /// <param name="treeScope"></param>
        /// <returns></returns>
        /// <exception cref="AutomationException">If not found</exception>
        public static AutomationElement FindElementByName(AutomationElement root, string automationName, TreeScope treeScope)
        {
            var el = root.FindFirst(treeScope, new PropertyCondition(AutomationElement.NameProperty,
                    automationName, PropertyConditionFlags.IgnoreCase));

            if (el == null)
            {
                throw new ApplicationException("Could not find element named '" + automationName + "'");
            }

            return el;
        }

        public static AutomationElement FindElementById(AutomationElement root, string automationId)
        {
            return FindElementById(root, automationId, TreeScope.Children);
        }

        public static AutomationElement FindElementByIdChildByClassname(AutomationElement root, string automationId, string childClassName, TreeScope treeScope)
        {
            var automationIdEl = FindElementById(root, automationId, treeScope);

            var classNameChildEl = automationIdEl.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ClassNameProperty,
                    childClassName));

            if (classNameChildEl == null)
            {
                throw new ApplicationException("Could not find child element having classname '" + childClassName + "' of element having id '" + automationId + "'");
            }

            return classNameChildEl;
        }

        /// <summary>
        /// Find first element by name, throwing if not found
        /// </summary>
        /// <param name="root"></param>
        /// <param name="automationName"></param>
        /// <param name="treeScope"></param>
        /// <returns></returns>
        /// <exception cref="AutomationException">If not found</exception>
        public static AutomationElement FindElementById(AutomationElement root, string automationId, TreeScope treeScope)
        {
            var el = root.FindFirst(treeScope, new PropertyCondition(AutomationElement.AutomationIdProperty,
                    automationId, PropertyConditionFlags.IgnoreCase));

            if (el == null)
            {
                throw new ApplicationException("Could not find element having id '" + automationId + "'");
            }

            return el;
        }

        public static AutomationElement FindElementByClassNameAndNameWithTimeout(AutomationElement parent, string className, string automationName, TimeSpan timeout)
        {
            return FindElementByClassNameAndNameWithTimeout(parent, className, automationName, timeout,
                    TreeScope.Children);
        }

        public static AutomationElement FindElementByClassNameAndNameWithTimeout(AutomationElement parent, string className, string automationName, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                
                var win = parent.FindFirst(treeScope, 
                        new AndCondition(
                                    new PropertyCondition(AutomationElement.ClassNameProperty, className),
                                    new PropertyCondition(AutomationElement.NameProperty, automationName, PropertyConditionFlags.IgnoreCase)));

                if (win != null)
                {
                    return win;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out finding element with classname '" + className + "' and name '" + automationName + "' after " + timeout.TotalSeconds + "sec");
        }

        public static AutomationElement FindElementByControlTypeAndNameWithTimeout(AutomationElement parent, ControlType controlType, ControlType controlType2, string automationName, TimeSpan timeout)
        {
            return FindElementByControlTypeAndNameWithTimeout(parent, controlType, controlType2, automationName, timeout,
                    TreeScope.Children);
        }

        public static AutomationElement FindElementByControlTypeAndNameWithTimeout(AutomationElement parent, ControlType controlType, ControlType controlType2, string automationName, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {

                var win = parent.FindFirst(treeScope,
                        new AndCondition(
                                    new OrCondition(
                                                new PropertyCondition(AutomationElement.ControlTypeProperty, controlType),
                                                new PropertyCondition(AutomationElement.ControlTypeProperty, controlType2)),
                                    new PropertyCondition(AutomationElement.NameProperty, automationName, PropertyConditionFlags.IgnoreCase)));

                if (win != null)
                {
                    return win;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out finding element with ControlType " + controlType + " or " + controlType2 + " and name '" + automationName + "' after " + timeout.TotalSeconds + "sec");
        }

        public static AutomationElement FindElementByNameWithTimeout(AutomationElement parent, string automationName, TimeSpan timeout)
        {
            return FindElementByNameWithTimeout(parent, automationName, timeout,
                    TreeScope.Children);
        }

        public static AutomationElement FindElementByNameWithTimeout(AutomationElement parent, string automationName, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                var win = parent.FindFirst(treeScope, new PropertyCondition(AutomationElement.NameProperty,
                        automationName, PropertyConditionFlags.IgnoreCase));

                if (win != null)
                {
                    return win;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out finding element with name '" + automationName + "' after " + timeout.TotalSeconds + "sec");
        }

        public static void WaitForWindowToDisappearByIdWithTimeout(AutomationElement parent, string automationId, TimeSpan timeout)
        {
            WaitForWindowToDisappearByIdWithTimeout(parent, automationId, timeout,
                    TreeScope.Children);
        }

        public static void WaitForWindowToDisappearByIdWithTimeout(AutomationElement parent, string automationId, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                var win = parent.FindFirst(treeScope, new PropertyCondition(AutomationElement.AutomationIdProperty,
                        automationId, PropertyConditionFlags.IgnoreCase));

                if (win == null)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out waiting for window with id '" + automationId + "' to disappear after " + timeout.TotalSeconds + "sec");
        }

        public static void WaitForWindowToDisappearByNameWithTimeout(AutomationElement parent, string automationName, TimeSpan timeout)
        {
            WaitForWindowToDisappearByNameWithTimeout(parent, automationName, timeout,
                    TreeScope.Children);
        }

        public static void WaitForWindowToDisappearByNameWithTimeout(AutomationElement parent, string automationName, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                var win = parent.FindFirst(treeScope, new PropertyCondition(AutomationElement.NameProperty,
                        automationName, PropertyConditionFlags.IgnoreCase));

                if (win == null)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out waiting for window with name '" + automationName + "' to disappear after " + timeout.TotalSeconds + "sec");
        }

        public static AutomationElement FindElementByIdWithTimeout(AutomationElement parent, string automationId, TimeSpan timeout)
        {
            return FindElementByIdWithTimeout(parent, automationId, timeout,
                    TreeScope.Children);
        }

        public static AutomationElement FindElementByIdWithTimeout(AutomationElement parent, string automationId, TimeSpan timeout, TreeScope treeScope)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                var win = parent.FindFirst(treeScope, new PropertyCondition(AutomationElement.AutomationIdProperty,
                        automationId, PropertyConditionFlags.IgnoreCase));

                if (win != null)
                {
                    return win;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out finding element with id '" + automationId + "' after " + timeout.TotalSeconds + "sec");
        }

        public static void WaitForElementEnabledWithTimeout(AutomationElement el, TimeSpan timeout)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                if (el.Current.IsEnabled)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            throw new TimeoutException("Timed out waiting for element '" + GetIdOrName(el) + "' to become enabled");
        }

        public static void WaitForElementVisibleWithTimeout(AutomationElement el, TimeSpan timeout, string nameOverride)
        {
            var sw = Stopwatch.StartNew();
            do
            {
                if ( ! el.Current.IsOffscreen)
                {
                    return;
                }

                Thread.Sleep(TimeSpan.FromMilliseconds(100)); //be cpu-friendly
            }
            while (sw.Elapsed < timeout);

            string nameToUse = GetIdOrName(el);
            if ( ! string.IsNullOrEmpty(nameOverride))
            {
                nameToUse = nameOverride;
            }
            throw new TimeoutException("Timed out waiting for element '" + nameToUse + "' to become visible");
        }

        public static void SelectMenu(AutomationElement elem)
        {
            var patt = elem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
            patt.Select();
        }

        public static void MouseInvokeFromClickablePoint(AutomationElement elem)
        {
            var clickablePoint = elem.GetClickablePoint();

            Mouse.MoveTo(new System.Drawing.Point((int)clickablePoint.X, (int)clickablePoint.Y));

            if ( ! elem.Current.IsEnabled)
            {
                throw new InvalidOperationException("Element '" + GetIdOrName(elem) + "' cannot be clicked since IsEnabled=false");
            }

            Mouse.Click(MouseButton.Left);
        }

        public static void MouseClickInferPoint(AutomationElement element, TimeSpan delayBeforeClick, int offsetX = -1, int offsetY = -1)
        {
            var elementTopLeft = element.Current.BoundingRectangle.TopLeft;
            var elementBottomRight = element.Current.BoundingRectangle.BottomRight;

            var offsetXToUse = offsetX;
            if (offsetXToUse == -1)
            {
                offsetXToUse = Convert.ToInt32(elementBottomRight.X - elementTopLeft.X) / 2;
            }

            var offsetYToUse = offsetY;
            if (offsetYToUse == -1)
            {
                offsetYToUse = Convert.ToInt32(elementBottomRight.Y - elementTopLeft.Y) / 2;
            }

            Mouse.MoveTo(new System.Drawing.Point((int)elementTopLeft.X + offsetXToUse, (int)elementTopLeft.Y + offsetYToUse));

            //delay before click
            Thread.Sleep(delayBeforeClick);

            if ( ! element.Current.IsEnabled)
            {
                throw new InvalidOperationException("Element '" + GetIdOrName(element) + "' cannot be clicked since IsEnabled=false");
            }

            Mouse.Click(MouseButton.Left);
        }

        public static void PressButton(AutomationElement elem)
        {
            if (!elem.Current.IsEnabled)
            {
                throw new InvalidOperationException("Element '" + GetIdOrName(elem) + "' cannot be invoked since IsEnabled=false");
            }

            var patt = elem.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            patt.Invoke();
        }

        public static bool GetCheckboxValue(AutomationElement elem)
        {
            var patt = elem.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;

            ToggleState state = patt.Current.ToggleState;

            if (state == ToggleState.On)
                return true;

            if (state == ToggleState.Off)
                return false;

            throw new InvalidOperationException("Unsupported toggle state on checkbox: " + state);
        }

        public static void SetCheckbox(AutomationElement elem, bool val)
        {
            if ( ! elem.Current.IsEnabled)
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + elem.Current.AutomationId.ToString()
                    + " is not enabled");
            }

            var patt = elem.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
 
            ToggleState state = patt.Current.ToggleState;

            if (val && state != ToggleState.On)
            {
                patt.Toggle();
            }
            else if ( ! val && state != ToggleState.Off)
            {
                patt.Toggle();
            }
        }

        /// <summary>
        /// Select datagrid row and ensure visible after select
        /// </summary>
        /// <param name="el"></param>
        /// <param name="gridValueName"></param>
        /// <param name="nameCol"></param>
        /// <exception cref="ApplicationException">Element does not support GridPattern</exception>
        /// <exception cref="ApplicationException">Specified row not found</exception>
        /// <returns>Selected row element</returns>
        public static AutomationElement SelectGridRow(AutomationElement el, string gridValueName, int nameCol)
        {
            //TODO: Improve current brittleness here when used with custom control SortedListView
            //Doesn't always find item depending on where it is in the list
            //Using a backward search for now since that works with the current testuser@testdomain.com / DemoUrl data set

            object gridPatternObj;
            if (!el.TryGetCurrentPattern(GridPattern.Pattern, out gridPatternObj))
            {
                throw new ApplicationException("Specified element '" + UIAUtility.GetIdOrName(el) + "' does not support GridPattern");
            }

            var gridPattern = (GridPattern) gridPatternObj;
            
            AutomationElement matchRow = null;

            for (int i = gridPattern.Current.RowCount - 1; i >= 0; i--)
            {
                var elem = gridPattern.GetItem(i, nameCol);
                if (elem == null)
                {
                    elem = gridPattern.GetItem(i, nameCol);
                }
                if (elem != null)
                {
                    if (elem.Current.Name == gridValueName)
                    {
                        var tw = TreeWalker.ContentViewWalker;
                        matchRow = tw.GetParent(elem);
                        break;
                    }
                }
            }

            if (matchRow == null)
            {
                throw new ApplicationException("Did not find specified row '" + gridValueName + "'");
            }

            WaitForElementEnabledWithTimeout(matchRow, AddinTestUtility.DialogControlEventStateUpdateTimeout);

            //Select grid row
            object selectionPattern;
            if ( ! matchRow.TryGetCurrentPattern(SelectionItemPattern.Pattern, out selectionPattern))
            {
                throw new ApplicationException("Specified element '" + UIAUtility.GetIdOrName(matchRow) + "' does not support SelectionItemPattern");
            }
            ((SelectionItemPattern) selectionPattern).Select();

            return matchRow;
        }

        public static string GetText(AutomationElement element)
        {
            object textPattern = null;

            if (!element.TryGetCurrentPattern(
                TextPattern.Pattern, out textPattern))
            {
                throw new InvalidOperationException(
                        "The control with an AutomationID of "
                        + element.Current.AutomationId.ToString()
                        + " does not support TextPattern.");
            }

            return ((TextPattern)textPattern).DocumentRange.GetText(-1);
        }

        /// <summary> 
        /// Inserts a string into textbox control
        /// </summary> 
        /// <param name="element">A text control.</param>
        /// <param name="value">The string to be inserted.</param>
        public static void InsertText(AutomationElement element,
                                            string value)
        {
            // Validate arguments / initial setup 
            if (value == null)
                throw new ArgumentNullException(
                    "String parameter must not be null.");

            if (element == null)
                throw new ArgumentNullException(
                    "AutomationElement parameter must not be null");

            // A series of basic checks prior to attempting an insertion. 
            // 
            // Check #1: Is control enabled? 
            // An alternative to testing for static or read-only controls  
            // is to filter using  
            // PropertyCondition(AutomationElement.IsEnabledProperty, true)  
            // and exclude all read-only text controls from the collection. 
            if ( ! element.Current.IsEnabled)
            {
                throw new InvalidOperationException(
                    "The control with an AutomationID of "
                    + element.Current.AutomationId.ToString()
                    + " is not enabled");
            }

            // Once you have an instance of an AutomationElement,   
            // check if it supports the ValuePattern pattern. 
            object valuePattern = null;

            // Control does not support the ValuePattern pattern  
            // so use keyboard input to insert content. 
            // 
            // NOTE: Elements that support TextPattern  
            //       do not support ValuePattern and TextPattern 
            //       does not support setting the text of  
            //       multi-line edit or document controls. 
            //       For this reason, text input must be simulated 
            //       using one of the following methods. 
            //        
            if (!element.TryGetCurrentPattern(
                ValuePattern.Pattern, out valuePattern))
            {
                throw new InvalidOperationException(
                        "The control with an AutomationID of "
                        + element.Current.AutomationId.ToString()
                        + " does not support ValuePattern.");
            }

            // Control supports the ValuePattern pattern so we can  
            // use the SetValue method to insert content. 
            // Set focus for input functionality and begin.
            element.SetFocus();

            ((ValuePattern)valuePattern).SetValue(value);
        }

        public static string GetIdOrName(AutomationElement el)
        {
            if (string.IsNullOrEmpty(el.Current.AutomationId))
            {
                return el.Current.Name;
            }
            else
            {
                return el.Current.AutomationId;
            }
        }
    }
}
