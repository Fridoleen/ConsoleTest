using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System;
using System.Threading;

namespace WinAppManipulator
{
    public class NotepadHelper
    {
        public void CreateMessageFileWithHotkeys()
        {
            var application = Application.Launch("notepad.exe");
            var mainWindowNotepad = application.GetMainWindow(new UIA3Automation(), TimeSpan.FromSeconds(2));
            ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());

            var textEditor = mainWindowNotepad.FindFirstDescendant(cf.ByName("Text Editor"));
            textEditor.Focus();

            Keyboard.Type("I'm alive!!! (c) Skynet");

            using (Keyboard.Pressing(VirtualKeyShort.CONTROL))
            {
                Keyboard.Press(VirtualKeyShort.KEY_S);               
            }

            //TODO: replace this with checking for availability of filename textbox
            Thread.Sleep(500);

            Keyboard.Type("Message_from_your_PC");
            Keyboard.Press(VirtualKeyShort.ENTER);

            application.CloseTimeout = TimeSpan.FromSeconds(2);
            application.Close();
        }

        public void CreateMessageFileWithMenuInteraction()
        {
            var application = Application.Launch("notepad.exe");

            var mainWindowNotepad = application.GetMainWindow(new UIA3Automation(), TimeSpan.FromSeconds(2));
            ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());


            mainWindowNotepad.FindFirstDescendant(cf.ByName("Text Editor")).AsTextBox().Enter("I'm alive!!! (c) Skynet");

            mainWindowNotepad.FindFirstDescendant("MenuBar").AsMenu().Items["File"].Click();

            Retry.Find(() => mainWindowNotepad.FindFirstDescendant(cf.ByName("Save As...")),
                new RetrySettings
                {
                    Timeout = TimeSpan.FromSeconds(2),
                    Interval = TimeSpan.FromMilliseconds(500)
                }
            ).Click();

            Retry.Find(() => mainWindowNotepad.FindFirstDescendant("1001"),
                new RetrySettings
                {
                    Timeout = TimeSpan.FromSeconds(2),
                    Interval = TimeSpan.FromMilliseconds(500)
                }
            ).AsTextBox().Enter("Message_from_your_PC");

            mainWindowNotepad.FindFirstDescendant(cf.ByName("Save")).AsButton().Click();

            application.CloseTimeout = TimeSpan.FromSeconds(2);
            application.Close();
        }
    }
}
