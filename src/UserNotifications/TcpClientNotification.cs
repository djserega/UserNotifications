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

        internal string IDConnection { get; set; }
        internal bool Connected
        {
            get
            {
                if (_tcpClient == null)
                    return false;

                return _tcpClient.Connected;
            }
        }

        internal TcpClientNotification()
        {
            _currentIP = Dns.GetHostEntry(Dns.GetHostName()).AddressList.FirstOrDefault(f => f.AddressFamily == AddressFamily.InterNetwork);
        }


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
                case "#DisconnectUser":
                    return SendMessage(GetServiceMessageDisconnectUser(param));
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

        internal void ConnectedTcpServer()
        {
            _tcpClient = new TcpClient(_currentIP.ToString(), _port);
        }



        #endregion

        #region Service

        private string GetServiceMessageConnectUser(params string[] param)
        {
            StringBuilder stringBuilder = new StringBuilder("#ConnectUser");

            stringBuilder.AppendLine("");

            stringBuilder.Append("#UserName=");
            stringBuilder.Append(param[0]);
            stringBuilder.AppendLine("");

            stringBuilder.Append("#ID=");
            stringBuilder.Append(IDConnection);
            stringBuilder.AppendLine("");

            stringBuilder.Append("#IP=");
            stringBuilder.Append(_currentIP);
            stringBuilder.AppendLine("");

            return stringBuilder.ToString();
        }

        private string GetServiceMessageDisconnectUser(params string[] param)
        {
            StringBuilder stringBuilder = new StringBuilder("#DisconnectUser");

            stringBuilder.AppendLine("");

            stringBuilder.Append("#ID=");
            stringBuilder.Append(IDConnection);
            stringBuilder.AppendLine("");

            return stringBuilder.ToString();
        }

        #endregion

    }
}
