using NUnit.Framework;
using System;
using System.IO;
using System.Threading;
using WinAppManipulator;
using Allure.Commons;
using NUnit.Allure.Steps;
using NUnit.Allure.Core;

namespace TestLib
{
    [AllureNUnit]
    [TestFixture]
    public class ThisPCShould
    {
        const string messageFilePath = @"G:\Downloads\Message_from_your_PC.txt";

        [AllureStep("OpenAndCloseWord")]
        [Ignore("Unnecessary")]
        [Test]
        public void OpenAndCloseWord()
        {
            new FileInfo(@"G:\Downloads\Screenshots\WordWasOpenedEvidence.png").Delete();

            var wh = new WordHelper();

            var scr = wh.OpenAndCloseWord();

            Assert.AreNotEqual(null, scr);

            scr.ToFile(@"G:\Downloads\Screenshots\WordWasOpenedEvidence.png");
            //AllureLifecycle.Instance.AddAttachment("Word has been opened", "image/png", @"G:\Downloads\WordWasOpenedEvidence.png");
        }

        [AllureStep("CheckIf_WordLowerPannel_IsMissing")]
        //[Ignore("Works just fine")]
        [Test]
        public void CheckIfPannelIsHidden()
        {
            var wh = new WordHelper();
            wh.DisableBottomPannel();

            Thread.Sleep(TimeSpan.FromSeconds(2));

            wh = new WordHelper();
            var checker = wh.CheckIfPannelIsAccesible();

            Assert.That(checker, Is.EqualTo(true));            
        }

        [AllureStep("Hotkeys-way txt message file creation")]
        [Ignore("This one works")]
        [Test]
        public void CheckMessageFromPC_ByHotkeys()
        {
            DeleteFileTxt();
            var wh = new NotepadHelper();
            wh.CreateMessageFileWithHotkeys();
            var file = new FileInfo(messageFilePath);

            string text;
            using (StreamReader reader = file.OpenText())
            {
                text = reader.ReadLine();
            }

            Assert.That(text, Is.EqualTo("I'm alive!!! (c) Skynet"));
        }

        [AllureStep("Menu-way txt message file creation")]
        [Ignore("This one works fine")]
        [Test]
        public void CheckMessageFromPc_ByMenuManipulation()
        {            
            var wh = new NotepadHelper();
            DeleteFileTxt();
            wh.CreateMessageFileWithMenuInteraction();
            var file = new FileInfo(messageFilePath);            

            string text;
            using (StreamReader reader = file.OpenText())
            {
                text = reader.ReadLine();
            }

            Assert.That(text, Is.EqualTo("I'm alive!!! (c) Skynet"));
        }


        private void DeleteFileTxt()
        {
            var file = new FileInfo(messageFilePath);
            if (file.Exists) file.Delete();
        }
    }
}
