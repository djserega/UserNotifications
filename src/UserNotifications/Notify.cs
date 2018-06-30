using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace AddIn
{
    internal class Notify
    {
        private readonly UserBalloonTipEvent _userBalloonTipClickedEvent;
        private string _title;
        private NotifyIcon _notifyIcon;

        internal Notify(UserBalloonTipEvent userBalloonTipClickedEvents)
        {
            _userBalloonTipClickedEvent = userBalloonTipClickedEvents;

            _notifyIcon = new NotifyIcon()
            {
                BalloonTipIcon = ToolTipIcon.Info,
                Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location)
            };
            _notifyIcon.BalloonTipClicked += _notifyIcon_BalloonTipClicked;
            _notifyIcon.BalloonTipClosed += _notifyIcon_BalloonTipClosed;
            
        }

        private void _notifyIcon_BalloonTipClosed(object sender, EventArgs e)
        {
        }

        private void _notifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            _userBalloonTipClickedEvent.InvokeUserBalloonTipClicked(((NotifyIcon)sender).BalloonTipText);
        }


        internal void SetTitle(string title) => _title = title;

        internal void ShowMessage(string message)
        {
            Show();

            _notifyIcon.BalloonTipTitle = _title;
            _notifyIcon.BalloonTipText = message;
            _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

            _notifyIcon.ShowBalloonTip(0);
        }


        internal void Show() => _notifyIcon.Visible = true;
        internal void Hide() => _notifyIcon.Visible = false;
    }
}
