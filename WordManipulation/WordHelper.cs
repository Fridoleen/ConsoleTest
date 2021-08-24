using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System;
using System.Threading;

namespace WinAppManipulator
{
    public class WordHelper
    {
        public CaptureImage OpenAndCloseWord()
        {
            using (var application = Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"))
            {
                application.WaitWhileBusy(TimeSpan.FromSeconds(2));

                using (var automation = new UIA3Automation())
                {
                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    var button = Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Help")),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(3),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        );

                    application.Close();
                    return Capture.MainScreen();                   
                }                
            }            
        }

        public void DisableBottomPannel()
        {
            using (var application = Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"))
            {
                using (var automation = new UIA3Automation())
                {
                    //TODO: Change this into acceptable form
                    Thread.Sleep(TimeSpan.FromSeconds(3));

                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    Retry.Find(() => screen.FindFirstDescendant("AIOStartDocument"),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(5),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        ).Click();

                    screen.FindFirstDescendant(cf.ByName("Collapse the Ribbon")).AsButton().Click();
                }

                application.CloseTimeout = TimeSpan.FromSeconds(2);
                application.Close();
            }
        }

        public bool CheckIfPannelIsAccesible()
        {
            using (var application = Application.Launch(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"))
            {
                using (var automation = new UIA3Automation())
                {
                    //TODO: Change this too
                    Thread.Sleep(TimeSpan.FromSeconds(3));

                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    Retry.Find(() => screen.FindFirstDescendant("AIOStartDocument"),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(5),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        ).Click();

                    var element = screen.FindFirstDescendant(cf.ByName("Lower Ribbon"));                    

                    using (Keyboard.Pressing(VirtualKeyShort.CONTROL))
                    {
                        Keyboard.Press(VirtualKeyShort.F1);
                    }

                    application.Close();

                    return element == null;
                }
            }
        }
    }
}
