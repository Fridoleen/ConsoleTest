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
                application.WaitWhileBusy(TimeSpan.FromSeconds(5));
                application.WaitWhileMainHandleIsMissing(TimeSpan.FromSeconds(5));
                var screen = application.GetMainWindow(new UIA3Automation());
                var cf = new ConditionFactory(new UIA3PropertyLibrary());

                var button = Retry.Find(() => screen.FindFirstDescendant(cf.ByName("Справка")),
                        new RetrySettings
                        {
                            Timeout = TimeSpan.FromSeconds(2),
                            Interval = TimeSpan.FromMilliseconds(500)
                        }
                    );

                var img = Capture.MainScreen();

                application.Close();

                return img;
            }            
        }

        public void CreateMessageByHotkeys()
        {

        }

        public void CreateMessageByMenus()
        {

        }

        public void PinTopPannelByHotkeys()
        {

        }

        public void EnableNavigationPannel()
        {

        }

        public void DisableNavigationPannel()
        {

        }

    }
}
