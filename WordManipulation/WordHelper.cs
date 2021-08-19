using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System;
using System.Threading;

namespace WinAppManipulator
{
    public class WordHelper
    {
        public bool OpenAndCloseWord()
        {
            var application = Application.Launch(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");

            application.CloseTimeout = TimeSpan.FromSeconds(5);

            return application.Close();
        }

    }
}
