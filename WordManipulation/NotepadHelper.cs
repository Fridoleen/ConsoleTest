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
    public class NotepadHelper
    {
        public void CreateMessageFileWithHotkeys()
        {
            var application = Application.Launch("notepad.exe");
            var mainWindowNotepad = application.GetMainWindow(new UIA3Automation(), TimeSpan.FromSeconds(10));
            ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());

            var textEditor = mainWindowNotepad.FindFirstDescendant(cf.ByName("Text Editor"));
            textEditor.Focus();

            Keyboard.Type("I'm alive!!! (c) Skynet");

            Keyboard.Pressing(VirtualKeyShort.CONTROL);
            Keyboard.Press(VirtualKeyShort.KEY_S);
            Keyboard.Release(VirtualKeyShort.CONTROL);

            //TODO: replace this with checking for availability of filename textbox
            Thread.Sleep(500);

            Keyboard.Type("Message_from_your_PC");
            Keyboard.Press(VirtualKeyShort.ENTER);

            application.CloseTimeout = TimeSpan.FromSeconds(2);
            application.Close();
        }

        public void CreateMessageFileWithMenuInteraction()
        {

            //textEditor.AsTextBox().Enter("I'm alive!!! (c) Skynet");

            //var mainMenuNotepad = mainWindowNotepad.FindFirstDescendant(cf.ByAutomationId("MenuBar")).AsMenu();
            //var file = mainMenuNotepad.Items["File"];
            //file.DrawHighlight();
            ////file.Invoke();
            //file.Expand();
            //var collection = mainWindowNotepad.FindAllDescendants(cf.ByAutomationId("4"));
            //collection[0].Click();
            //var collection2 = mainWindowNotepad.FindAllDescendants(cf.ByAutomationId("1001"));
            //collection2[0].AsTextBox().Enter("Message_from_your_PC");

            //mainWindowNotepad.FindFirstDescendant(cf.ByName("Save")).AsButton().Click();
        }
    }
}
