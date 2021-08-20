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
            using (var application = Application.Launch(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"))
            {
                application.WaitWhileBusy(TimeSpan.FromSeconds(2));

                using (var automation = new UIA3Automation())
                {
                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    var button = Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Справка")),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(2),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        );        

                    return Capture.MainScreen();                   
                }
            }            
        }

        public void DisableBottomPannel()
        {
            using (var application = Application.Launch(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"))
            {
                application.WaitWhileBusy(TimeSpan.FromSeconds(2));

                using (var automation = new UIA3Automation())
                {
                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Новый документ")),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(2),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        ).Click();

                    screen.FindFirstDescendant(cf.ByName("Свернуть ленту")).AsButton().Click();
                }

                application.Close();
            }
        }

        public bool CheckIfPannelIsAccesible()
        {
            using (var application = Application.Launch(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"))
            {
                application.WaitWhileBusy(TimeSpan.FromSeconds(2));

                using (var automation = new UIA3Automation())
                {
                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    Retry.Find(() => screen.FindFirstDescendant("AIOStartDocument"),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(2),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        ).Click();

                    var element = Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Нижняя лента")),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(2),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        );
                    return element != null;
                }
            }
        }
    }
}
