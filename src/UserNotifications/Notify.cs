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
    public class Notify : IDisposable
    {
        private readonly string _title;
        private NotifyIcon _notifyIcon;

        public Notify(string title)
        {
            _title = title;

            _notifyIcon = new NotifyIcon()
            {
                BalloonTipIcon = ToolTipIcon.Info,
                Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location),
                Visible = true
            };
        }

        public void Dispose()
        {
            _notifyIcon?.Dispose();
        }

        public void Show(string message)
        {
            _notifyIcon.ShowBalloonTip(0, _title, message, ToolTipIcon.Info);
        }
    }
}
