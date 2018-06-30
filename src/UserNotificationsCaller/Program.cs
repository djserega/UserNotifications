using AddIn;
using System;

namespace UserNotificationsCaller
{
    class Program
    {
        static void Main(string[] args)
        {
            Notifications com = new Notifications();
            com.SetTitle("c#7");
            do
            {
                com.Show("time now: " + DateTime.Now.ToString("s"));
            } while (Console.ReadLine() != "exit");
        }
    }
}
