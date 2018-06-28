using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace UserNotifications
{
    [Guid("6314C210-62C0-4673-AC3E-24BD6A120A2B")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(INotifications))]
    public class Notifications : INotifications
    {
        private Notify _notify;

        public Notifications(string dbName)
        {
            _notify = new Notify(dbName);
        }

        public void Show(string message)
        {
            _notify.Show(message);
        }
    }
}
