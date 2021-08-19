using System;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }

    public class WordHelper
    {
        public bool OpenAndCloseWord()
        {
            var application = FlaUI.Core.Application.Launch("WINWORD.EXE");
            application.CloseTimeout = new TimeSpan(10);
            return application.Close();
        }
    }
}
