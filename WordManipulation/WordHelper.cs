﻿using FlaUI.Core;
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

                    var button = Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Справка")),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(2),
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
                //application.WaitWhileBusy(TimeSpan.FromSeconds(5));

                using (var automation = new UIA3Automation())
                {
                    Thread.Sleep(TimeSpan.FromSeconds(3));
                    var screen = application.GetMainWindow(automation);
                    var cf = new ConditionFactory(new UIA3PropertyLibrary());

                    Thread.Sleep(TimeSpan.FromSeconds(3));
                    screen.FindFirstDescendant("AIOStartDocument").AsButton().Click();
                    //Retry.Find(() => screen.FindFirstDescendant("AIOStartDocument"),
                    //        new RetrySettings
                    //        {
                    //            Timeout = TimeSpan.FromSeconds(5),
                    //            Interval = TimeSpan.FromMilliseconds(500)
                    //        }
                    //    ).AsButton().Click();

                    Thread.Sleep(TimeSpan.FromSeconds(3));

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
                //application.WaitWhileBusy(TimeSpan.FromSeconds(2));

                using (var automation = new UIA3Automation())
                {
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

                    //Thread.Sleep(TimeSpan.FromSeconds(5));

                    var element = Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Lower Ribbon")),
                            new RetrySettings
                            {
                                Timeout = TimeSpan.FromSeconds(5),
                                Interval = TimeSpan.FromMilliseconds(500)
                            }
                        );

                    using (Keyboard.Pressing(VirtualKeyShort.CONTROL))
                    {
                        Keyboard.Press(VirtualKeyShort.F11);
                    }

                    application.Close();

                    return element == null;
                }
            }
        }
    }
}
