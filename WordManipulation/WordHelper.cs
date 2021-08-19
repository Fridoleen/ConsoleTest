using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System;
using System.Threading;

namespace WordManipulation
{
    public class WordHelper
    {
        public bool OpenAndCloseWord()
        {
            var application = Application.Launch(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");

            application.CloseTimeout = TimeSpan.FromSeconds(5);

            return application.Close();
        }

        public void OpenAndCloseNotepad()
        {
            var application = Application.Launch("notepad.exe");
            var mainWindowNotepad = application.GetMainWindow(new UIA3Automation(), TimeSpan.FromSeconds(10));
            ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());

            var textEditor = mainWindowNotepad.FindFirstDescendant(cf.ByName("Text Editor"));
            textEditor.Focus();
            textEditor.AsTextBox().Enter("I'm alive!!! (c) Skynet");

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
            Keyboard.Pressing(VirtualKeyShort.CONTROL);
            Keyboard.Press(VirtualKeyShort.KEY_S);
            Keyboard.Release(VirtualKeyShort.CONTROL);
            Keyboard.Type("Message_from_your_PC");
            Keyboard.Press(VirtualKeyShort.ENTER);

            application.CloseTimeout = TimeSpan.FromSeconds(5);
            application.Close();
        }

        public void OpenSkynetMessage()
        {
            var application = Application.Launch("notepad.exe");

            var mainWindowNotepad = application.GetMainWindow(new UIA3Automation(), TimeSpan.FromSeconds(10));
            var cf = new ConditionFactory(new UIA3PropertyLibrary());

            mainWindowNotepad.FindFirstDescendant(cf.ByName("File")).AsButton().Click();
            mainWindowNotepad.FindFirstDescendant(cf.ByName("Open...")).AsButton().Click();
            mainWindowNotepad.FindFirstDescendant(cf.ByName("Save")).AsButton().Click();

            application.CloseTimeout = TimeSpan.FromSeconds(5);
            application.Close();
        }
    }
}
