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
    internal class Notify
    {
        private string _title;
        private NotifyIcon _notifyIcon;

        internal Notify()
        {
            _notifyIcon = new NotifyIcon()
            {
                BalloonTipIcon = ToolTipIcon.Info,
                Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location),
                Visible = true
            };
        }

        internal void SetTitle(string title) => _title = title;

        internal void Show(string message) => _notifyIcon.ShowBalloonTip(0, _title, message, ToolTipIcon.Info);

        internal void Hide() => _notifyIcon.Visible = false;
    }
}
