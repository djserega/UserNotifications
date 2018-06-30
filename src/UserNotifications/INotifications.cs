using System.Runtime.InteropServices;

namespace AddIn
{
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface INotifications
    {
        void SetTitle(string title);
        void ShowMessage(string message);
        void Hide();
    }
}