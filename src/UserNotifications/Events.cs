using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
