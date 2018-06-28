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

        public Notifications()
        {
            
        }

        public void SetTitle(string title)
        {
            _notify = new Notify(title);
        }

        public string Show(string message)
        {
            if (_notify == null)
            {
                return "Не установлен заголовок.";
            };

            _notify.Show(message);

            return string.Empty;
        }
    }
}
