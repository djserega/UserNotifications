using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace UserNotifications
{
    internal class TcpClientNotification
    {
        private TcpClient _tcpClient;

        private readonly IPAddress _currentIP;
        private readonly int _port = 44566;

        internal bool Connected { get; set; }

        #region Exchange

        internal bool SendMessage(string message)
        {
            NetworkStream serverStream = _tcpClient.GetStream();

            byte[] data = Encoding.UTF8.GetBytes(message);

            try
            {
                serverStream.Write(data, 0, data.Length);
                return true;
            }
            catch (IOException)
            {
                return false;
            }
        }

        internal bool SendServiceMessage(string message, params string[] param)
        {
            switch (message)
            {
                case "#ConnectUser":
                    return SendMessage(GetServiceMessageConnectUser(param));
                default:
                    return false;
            }
        }

        internal async Task<string> ReadMessageAsync()
        {
            return await Task.Run(() => ReadMessage());
        }

        private string ReadMessage()
        {
            if (!Connected)
                return null;

            NetworkStream clientStream = _tcpClient.GetStream();

            byte[] data = new byte[256];

            int lenghtData;
            try
            {
                lenghtData = clientStream.Read(data, 0, data.Length);
            }
            catch (IOException ex)
            {
                return ex.Message;
            }

            if (lenghtData > 0)
            {
                return Encoding.UTF8.GetString(data, 0, lenghtData);
            }

            return null;
        }

        #endregion

        #region TCP

        internal TcpClientNotification()
        {
            _currentIP = Dns.GetHostEntry(Dns.GetHostName()).AddressList.FirstOrDefault(f => f.AddressFamily == AddressFamily.InterNetwork);
        }

        internal void ConnectedTcpServer()
        {
            _tcpClient = new TcpClient(_currentIP.ToString(), _port);
            Connected = _tcpClient.Connected;
        }



        #endregion

        #region Service

        private string GetServiceMessageConnectUser(params string[] param)
        {
            StringBuilder stringBuilder = new StringBuilder("#ConnectUser;");

            stringBuilder.AppendLine("#UserName=");
            stringBuilder.Append(param[0]);
            stringBuilder.Append(";");

            stringBuilder.AppendLine("#IP=");
            stringBuilder.Append(_currentIP);
            stringBuilder.Append(";");

            return stringBuilder.ToString();
        }

        #endregion

    }
}
