using System;
using System.Runtime.InteropServices;

namespace AddIn
{
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface INotifications
    {
        void SetTitle(string title);
        void ShowMessage(string message);
        void ShowMessageURL(string message, string url);
        void Hide();

        bool ConnectToService(string userName);
        bool CheckConnection();
        DateTime GetCurrentTime();
        string TextError { get; }
    }
}