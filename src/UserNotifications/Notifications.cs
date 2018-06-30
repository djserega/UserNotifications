using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AddIn
{
    [ProgId("AddIn.Notifications")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Notifications : AddIn, INotifications
    {
        private readonly UserBalloonTipEvent _userBalloonTipClickedEvent = new UserBalloonTipEvent();
        private Notify _notify;

        public Notifications()
        {
            _userBalloonTipClickedEvent.UserBalloonTipClicked += _userBalloonTipClickedEvent_UserBalloonTipClicked;
            _notify = new Notify(_userBalloonTipClickedEvent);
        }

        private void _userBalloonTipClickedEvent_UserBalloonTipClicked(string param)
        {
            string message;
            if (param.StartsWith($"e1cib/"))
                message = "OpenURL";
            else
                message = "Message";

            asyncEvent.ExternalEvent("UserNotifications", message, param);
        }

        public void SetTitle(string title) => _notify.SetTitle(title);

        public void Show(string message)
        {
            _notify.Show(message);
        }

        public void Show(string message, string url)
        {
            _notify.Show(message);
        }

        public void Hide() => _notify.Hide();

    }
}
