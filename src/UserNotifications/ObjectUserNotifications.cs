namespace AddIn
{
    internal class ObjectUserNotifications
    {
        public ObjectUserNotifications(string message, string url)
        {
            Message = message;
            URL = url;
        }

        internal string Message { get; set; }
        internal string URL { get; set; }
    }
}
