using NUnit.Framework;
using System;
using System.IO;
using System.Threading;
using WinAppManipulator;

namespace TestLib
{
    [TestFixture]
    public class ThisPCShould
    {
        const string messageFilePath = @"G:\Downloads\Message_from_your_PC.txt";

        [Ignore("Unnecessary")]
        [Test]
        public void OpenWord()
        {
            var wh = new WordHelper();

            Assert.AreEqual(true, wh.OpenAndCloseWord());
        }

        [Test]
        public void CheckMessageFromPC_ByHotkeys()
        {
            DeleteFile();
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

        [Test]
        public void CheckMessageFromPc_ByMenuManipulation()
        {
            DeleteFile();
            var wh = new NotepadHelper();
            wh.CreateMessageFileWithMenuInteraction();
            var file = new FileInfo(messageFilePath);

            string text;
            using (StreamReader reader = file.OpenText())
            {
                text = reader.ReadLine();
            }

            Assert.That(text, Is.EqualTo("I'm alive!!! (c) Skynet"));
        }


        public void DeleteFile()
        {
            var file = new FileInfo(messageFilePath);
            if (file.Exists) file.Delete();
        }
    }
}
