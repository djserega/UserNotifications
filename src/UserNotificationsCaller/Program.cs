using UserNotifications;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserNotificationsCaller
{
    class Program
    {
        static void Main(string[] args)
        {
            Notifications com = new Notifications();
            com.SetTitle("c#7");
            com.Show("time now: " + DateTime.Now.ToString("s"));
        }
    }
}
