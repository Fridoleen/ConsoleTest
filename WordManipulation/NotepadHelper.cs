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

            //Enter text in a new window
            mainWindowNotepad.FindFirstDescendant(cf.ByName("Text Editor")).AsTextBox().Enter("I'm alive!!! (c) Skynet");

            //Find and open File menu
            var fileMenu = mainWindowNotepad.FindFirstDescendant("MenuBar");
            fileMenu.AsMenu().Items["File"].Click();
            Thread.Sleep(500);

            //mainWindowNotepad.FindFirstDescendant("4").AsButton().Click();

            mainWindowNotepad.FindFirstDescendant(cf.ByName("Save As...")).AsButton().Click();

            Thread.Sleep(500);

            mainWindowNotepad.FindFirstDescendant("1001").AsTextBox().Enter("Message_from_your_PC");

            //Press save button
            mainWindowNotepad.FindFirstDescendant(cf.ByName("Save")).AsButton().Click();

            application.CloseTimeout = TimeSpan.FromSeconds(2);
            application.Close();
        }
    }
}
