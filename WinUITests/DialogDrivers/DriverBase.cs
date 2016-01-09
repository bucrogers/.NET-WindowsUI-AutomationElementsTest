using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace ItemSystemTests.DialogDrivers
{
    /// <summary>
    /// Base for all dialog automation drivers
    /// </summary>
    public abstract class DriverBase
    {
        protected AutomationElement _parentElement;
        protected AutomationElement _dialog;

        private Dictionary<string, string> _propBag = new Dictionary<string, string>();

        protected abstract string DialogId { get; }

        protected abstract void WaitForDialogReady();

        protected abstract void EnsureDialogEventsFired();

        protected void UpdatePropBagValue(string key, string val)
        {
            if (_propBag.ContainsKey(key))
            {
                _propBag[key] = val;
            }
            else
            {
                _propBag.Add(key, val);
            }
        }

        protected string GetPropBagValue(string key)
        {
            if (_propBag.ContainsKey(key))
            {
                return _propBag[key];
            }
            else
            {
                //Either was never set or a typeo in the referring name - no error here
                return string.Empty;
            }
        }
    }
}
