using System;

namespace AddIn
{
    internal delegate void UserBalloonTipClickedEvents(string param);
    internal class UserBalloonTipEvent : EventArgs
    {
        internal event UserBalloonTipClickedEvents UserBalloonTipClicked;

        internal void InvokeUserBalloonTipClicked(string param)
        {
            UserBalloonTipClicked?.Invoke(param);
        }
    }
}
