using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AddIn
{
    [ProgId("AddIn.Notifications")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Notifications : AddIn, INotifications
    {
        private readonly string _prefixUrl = "e1cib/";

        private Dictionary<int, ObjectUserNotifications> _dictionaryNotifications = new Dictionary<int, ObjectUserNotifications>();
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

            int hashCode = param.GetHashCode();
            if (_dictionaryNotifications.ContainsKey(hashCode))
            {
                param = _dictionaryNotifications[hashCode].URL;
            }

            if (param.StartsWith(_prefixUrl))
                message = "OpenURL";
            else
                message = "Message";

            asyncEvent.ExternalEvent("UserNotifications", message, param);
        }

        public void SetTitle(string title) => _notify.SetTitle(title);

        public void ShowMessage(string message)
        {
            _notify.ShowMessage(message);
        }

        public void ShowMessageURL(string message, string url)
        {
            _dictionaryNotifications.Add(message.GetHashCode(), new ObjectUserNotifications(message, url));
            _notify.ShowMessage(message);
        }

        public void Hide() => _notify.Hide();

    }
}
