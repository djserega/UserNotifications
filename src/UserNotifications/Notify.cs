using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UserNotifications
{
    internal class Notify : IDisposable
    {
        private string _title;
        private NotifyIcon _notifyIcon;
        
        internal Notify(string title)
        {
            _title = title;

            _notifyIcon = new NotifyIcon()
            {
                BalloonTipIcon = ToolTipIcon.Info,
                Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location)
            };
        }

        public void Dispose()
        {
            _notifyIcon?.Dispose();
        }

        internal void Show(string message)
        {
            _notifyIcon.ShowBalloonTip(0, _title, message, ToolTipIcon.Info);
        }
    }
}
