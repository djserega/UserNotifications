using System.Runtime.InteropServices;

namespace UserNotifications
{
    [Guid("B9A8CAE5-FD4E-4A2A-895B-6FE44ED95B27")]
    public interface INotifications
    {
        void SetTitle(string title);
        void Show(string message);
        void Hide();
    }
}