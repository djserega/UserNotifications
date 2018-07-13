using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using UserNotifications;

namespace AddIn
{
    [ProgId("AddIn.Notifications")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Notifications : INotifications, IInitDone, ILanguageExtender
    {

        #region Notifications

        private readonly string _prefixUrl = "e1cib/";

        private Dictionary<int, ObjectUserNotifications> _dictionaryNotifications = new Dictionary<int, ObjectUserNotifications>();
        private readonly UserBalloonTipEvent _userBalloonTipClickedEvent = new UserBalloonTipEvent();
        private Notify _notify;

        public string TextError { get; set; }

        public Notifications()
        {
            _userBalloonTipClickedEvent.UserBalloonTipClicked += _userBalloonTipClickedEvent_UserBalloonTipClicked;
            _notify = new Notify(_userBalloonTipClickedEvent);

            _tcpClientNotification = new TcpClientNotification();
        }

        ~Notifications()
        {
            if (_tcpClientNotification != null)
            {
                _tcpClientNotification.SendServiceMessage("#DisconnectUser");
                _tcpClientNotification = null;
            }
        }

        private void _userBalloonTipClickedEvent_UserBalloonTipClicked(string param)
        {
            string message;

            int hashCode = param.GetHashCode();
            if (_dictionaryNotifications.ContainsKey(hashCode))
            {
                param = _dictionaryNotifications[hashCode].URL;
                _dictionaryNotifications.Remove(hashCode);
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
            ShowMessage(message);
        }

        public void Hide() => _notify.Hide();

        #endregion

        #region TCP

        private TcpClientNotification _tcpClientNotification;
        private string _idConnect;

        private string ReadMessage()
        {
            if (!CheckConnection())
                return "Не подключено к сервису.";

            return _tcpClientNotification.ReadMessageAsync().Result;
        }

        public string ConnectToService(string userName, string id)
        {
            if (_tcpClientNotification.Connected)
                return "Подключение было выполнено ранее.";

            _tcpClientNotification.ConnectedTcpServer();

            if (!CheckConnection())
                return "Не подключено к сервису.";

            _tcpClientNotification.IDConnection = id;
            _tcpClientNotification.SendServiceMessage("#ConnectUser", userName);

            string data = ReadMessage();

            try
            {
                _idConnect = id;
                return data;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string DisconnectService()
        {
            if (!CheckConnection())
                return "Не подключено к сервису.";

            _tcpClientNotification.SendServiceMessage("#DisconnectUser");

            return ReadMessage();
        }

        public bool CheckConnection()
        {
            TextError = string.Empty;
            return _tcpClientNotification.Connected;
        }


        public DateTime GetCurrentTime()
        {
            if (!CheckConnection())
            {
                TextError = "Не подключено к сервису.";
                return DateTime.MinValue;
            }

            _tcpClientNotification.SendMessage("#GetCurrentTime");

            string data = ReadMessage();

            try
            {
                return Convert.ToDateTime(data);
            }
            catch (Exception ex)
            {
                TextError = $"{ex.Message}\nПередана строка: {data}";
                return DateTime.MinValue;
            }
        }

        #endregion

        #region Registration dll

        public const string AddInName = "Notifications";

        public const uint S_OK = 0;
        public const uint S_FALSE = 1;
        public const uint E_POINTER = 0x80004003;
        public const uint E_FAIL = 0x80004005;
        public const uint E_UNEXPECTED = 0x8000FFFF;

        public const short ADDIN_E_NONE = 1000;
        public const short ADDIN_E_ORDINARY = 1001;
        public const short ADDIN_E_ATTENTION = 1002;
        public const short ADDIN_E_IMPORTANT = 1003;
        public const short ADDIN_E_VERY_IMPORTANT = 1004;
        public const short ADDIN_E_INFO = 1005;
        public const short ADDIN_E_FAIL = 1006;
        public const short ADDIN_E_MSGBOX_ATTENTION = 1007;
        public const short ADDIN_E_MSGBOX_INFO = 1008;
        public const short ADDIN_E_MSGBOX_FAIL = 1009;

        /// <summary>Указатель на IDispatch</summary>
        protected object connect1c;

        /// <summary>Вызов событий 1С</summary>
        protected IAsyncEvent asyncEvent;

        /// <summary>Статусная строка 1С</summary>
        protected IStatusLine statusLine;

        private Type[] allInterfaceTypes;  // Коллекция интерфейсов
        private MethodInfo[] allMethodInfo;  // Коллекция методов
        private PropertyInfo[] allPropertyInfo; // Коллекция свойств

        private Hashtable nameToNumber; // метод - идентификатор
        private Hashtable numberToName; // идентификатор - метод
        private Hashtable numberToParams; // количество параметров метода
        private Hashtable numberToRetVal; // имеет возвращаемое значение (является функцией)
        private Hashtable propertyNameToNumber; // свойство - идентификатор
        private Hashtable propertyNumberToName; // идентификатор - свойство
        private Hashtable numberToMethodInfoIdx; // номер метода - индекс в массиве методов
        private Hashtable propertyNumberToPropertyInfoIdx; // номер свойства - индекс в массиве свойств

        /// <summary>
        /// При загрузке 1С:Предприятие V8 инициализирует объект компоненты,
        /// вызывая метод Init и передавая указатель на IDispatch.
        /// Объект может сохранить этот указатель для дальнейшего использования.
        /// Все остальные интерфейсы 1С:Предприятия объект может получить, вызвав метод QueryInterface
        /// переданного ему интерфейса IDispatch. Объект должен возвратить S_OK,
        /// если инициализация прошла успешно, и E_FAIL при возникновении ошибки.
        /// Данный метод может использовать интерфейс IErrorLog для вывода информации об ошибках.
        /// При этом инициализация считается неудачной, если одна из переданных структур EXCEPINFO
        /// имеет поле scode, не равное S_OK. Все переданные в IErrorLog данные обрабатываются
        /// при возврате из данного метода. В момент вызова этого метода свойство AppDispatch не определено.
        /// </summary>
        /// <param name="connection">reference to IDispatch</param>
        public void Init([MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            connect1c = connection;
            statusLine = (IStatusLine)connection;
            asyncEvent = (IAsyncEvent)connection;
        }

        /// <summary>
        /// 1С:Предприятие V8 вызывает этот метод при завершении работы с объектом компоненты.
        /// Объект должен возвратить S_OK. Этот метод вызывается независимо от результата
        /// инициализации объекта (метод Init).
        /// </summary>
        public void Done() { }

        /// <summary>
        /// 1С:Предприятие V8 вызывает этот метод для получения информации о компоненте.
        /// В текущей версии 2.0 компонентной технологии в элемент с индексом 0 необходимо
        /// записать версию поддерживаемой компонентной технологии в формате V_I4 — целого числа,
        /// при этом старший номер версии записывается в тысячные разряды,
        /// младший номер версии — в единицы. Например: версия 3.56 — число 3560.
        /// В настоящее время все объекты внешних компонент могут поддерживать версию 1.0
        /// (соответствует числу 1000) или 2.0 (соответствует 2000).
        /// </summary>
        /// <param name="info">Component information</param>
        public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
            => info[0] = 2000;

        /// <summary>Регистрация свойств и методов компоненты в 1C</summary>
        /// <param name="extensionName"></param>
        public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName)
        {
            try
            {
                Type type = this.GetType();

                allInterfaceTypes = type.GetInterfaces();
                allMethodInfo = type.GetMethods();
                allPropertyInfo = type.GetProperties();

                // Хэш-таблицы с именами методов и свойств компоненты
                nameToNumber = new Hashtable();
                numberToName = new Hashtable();
                numberToParams = new Hashtable();
                numberToRetVal = new Hashtable();
                propertyNameToNumber = new Hashtable();
                propertyNumberToName = new Hashtable();
                numberToMethodInfoIdx = new Hashtable();
                propertyNumberToPropertyInfoIdx = new Hashtable();

                int Identifier = 0;

                foreach (Type interfaceType in allInterfaceTypes)
                {
                    // Интересуют только методы в пользовательских интерфейсах, стандартные пропускаем
                    if (interfaceType.Name.Equals("IDisposable")
                      || interfaceType.Name.Equals("IManagedObject")
                      || interfaceType.Name.Equals("IRemoteDispatch")
                      || interfaceType.Name.Equals("IServicedComponentInfo")
                      || interfaceType.Name.Equals("IInitDone")
                      || interfaceType.Name.Equals("ILanguageExtender"))
                    {
                        continue;
                    };

                    // Обработка методов интерфейса
                    MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                    foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                    {
                        nameToNumber.Add(interfaceMethodInfo.Name, Identifier);
                        numberToName.Add(Identifier, interfaceMethodInfo.Name);
                        numberToParams.Add(Identifier, interfaceMethodInfo.GetParameters().Length);
                        numberToRetVal.Add(Identifier, (interfaceMethodInfo.ReturnType != typeof(void)));
                        Identifier++;
                    }

                    // Обработка свойств интерфейса
                    PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                    foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                    {
                        propertyNameToNumber.Add(interfacePropertyInfo.Name, Identifier);
                        propertyNumberToName.Add(Identifier, interfacePropertyInfo.Name);
                        Identifier++;
                    }
                }

                // Отображение номера метода на индекс в массиве
                foreach (DictionaryEntry entry in numberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allMethodInfo.Length; ii++)
                    {
                        if (allMethodInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            numberToMethodInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Метод " + entry.Value.ToString() + " не реализован");
                }

                // Отображение номера свойства на индекс в массиве
                foreach (DictionaryEntry entry in propertyNumberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allPropertyInfo.Length; ii++)
                    {
                        if (allPropertyInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            propertyNumberToPropertyInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Свойство " + entry.Value.ToString() + " не реализовано");
                }

                // Компонент инициализирован успешно. Возвращаем имя компонента.
                extensionName = AddInName;
            }
            catch (Exception)
            {
                return;
            }
        }

        /// <summary>Возвращает количество свойств</summary>
        /// <param name="props"></param>
        public void GetNProps(ref int props)
            => props = propertyNameToNumber.Count;

        /// <summary>Возвращает целочисленный идентификатор свойства, соответствующий переданному имени</summary>
        /// <param name="propName"></param>
        /// <param name="propNum"></param>
        public void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref int propNum)
            => propNum = (int)propertyNameToNumber[propName];

        /// <summary>Возвращает имя свойства, соответствующее переданному целочисленному идентификатору</summary>
        /// <param name="propNum"></param>
        /// <param name="propAlias"></param>
        /// <param name="propName"></param>
        public void GetPropName(int propNum, int propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName)
            => propName = (String)propertyNumberToName[propNum];

        /// <summary>Возвращает значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propVal"></param>
        public void GetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
            => propVal = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].GetValue(this, null);

        /// <summary>Устанавливает значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propVal"></param>
        public void SetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
            => allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal, null);

        /// <summary>Определяет, можно ли читать значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propRead"></param>
        public void IsPropReadable(int propNum, ref bool propRead)
            => propRead = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanRead;

        /// <summary>Определяет, можно ли изменять значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propWrite"></param>
        public void IsPropWritable(int propNum, ref Boolean propWrite)
            => propWrite = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanWrite;

        /// <summary>Возвращает количество методов</summary>
        /// <param name="pMethods"></param>
        public void GetNMethods(ref int pMethods)
            => pMethods = nameToNumber.Count;

        /// <summary>Возвращает идентификатор метода по его имени</summary>
        /// <param name="methodName">Имя метода</param>
        /// <param name="methodNum">Идентификатор метода</param>
        public void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref int methodNum)
            => methodNum = (int)nameToNumber[methodName];

        /// <summary>Возвращает имя метода по идентификатору</summary>
        /// <param name="methodNum"></param>
        /// <param name="methodAlias"></param>
        /// <param name="methodName"></param>
        public void GetMethodName(int methodNum, int methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName)
            => methodName = (String)numberToName[methodNum];

        /// <summary>Возвращает число параметров метода по его идентификатору</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Число параметров</param>
        public void GetNParams(int methodNum, ref int pParams)
            => pParams = (int)numberToParams[methodNum];

        /// <summary>Получить значение параметра метода по умолчанию</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="paramNum">Номер параметра</param>
        /// <param name="paramDefValue">Возвращаемое значение</param>
        public void GetParamDefValue(int methodNum, int paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue) { }

        /// <summary>Указывает, что у метода есть возвращаемое значение</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Наличие возвращаемого значения</param>
        public void HasRetVal(int methodNum, ref Boolean retValue)
            => retValue = (Boolean)numberToRetVal[methodNum];

        /// <summary>Вызов метода как процедуры с использованием идентификатора</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Параметры</param>
        public void CallAsProc(int methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                asyncEvent.ExternalEvent(AddInName, e.Message, e.ToString());
            }
        }

        /// <summary>Вызов метода как функции с использованием идентификатора</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Возвращаемое значение</param>
        /// <param name="pParams">Параметры</param>
        public void CallAsFunc(int methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                asyncEvent.ExternalEvent(AddInName, e.Message, e.ToString());
            }
        }

        #endregion

    }

    #region Registration dll

    /// <summary>Функции данного интерфейса вызываются при подключении компоненты</summary>
    /// 
    [Guid("AB634001-F13D-11d0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInitDone
    {
        void Init([MarshalAs(UnmanagedType.IDispatch)] object connection);
        void Done();
        void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info);
    }

    /// <summary>Интерфейс определяет логику вызова функций, процедур и свойств компоненты из 1С</summary>
    /// 
    [Guid("AB634003-F13D-11d0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILanguageExtender
    {
        void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref string extensionName);
        void GetNProps(ref int props);
        void FindProp([MarshalAs(UnmanagedType.BStr)] string propName, ref int propNum);
        void GetPropName(int propNum, int propAlias, [MarshalAs(UnmanagedType.BStr)] ref string propName);
        void GetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
        void SetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
        void IsPropReadable(int propNum, ref bool propRead);
        void IsPropWritable(int propNum, ref Boolean propWrite);
        void GetNMethods(ref int pMethods);
        void FindMethod([MarshalAs(UnmanagedType.BStr)] string methodName, ref int methodNum);
        void GetMethodName(int methodNum, int methodAlias, [MarshalAs(UnmanagedType.BStr)] ref string methodName);
        void GetNParams(int methodNum, ref int pParams);
        void GetParamDefValue(int methodNum, int paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue);
        void HasRetVal(int methodNum, ref Boolean retValue);
        void CallAsProc(int methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
        void CallAsFunc(int methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
    }

    /// <summary>Интерфейс реализован 1С для получения событий от компоненты</summary>
    /// 
    [Guid("ab634004-f13d-11d0-a459-004095e1daea")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAsyncEvent
    {
        void SetEventBufferDepth(int depth);
        void GetEventBufferDepth(ref long depth);
        void ExternalEvent([MarshalAs(UnmanagedType.BStr)] string source, [MarshalAs(UnmanagedType.BStr)] string message, [MarshalAs(UnmanagedType.BStr)] string data);
        void CleanBuffer();
    }

    /// <summary>С помощью этого интерфейса компонента получает доступ к строке состояния 1С</summary>
    /// 
    [Guid("AB634005-F13D-11D0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStatusLine
    {
        void SetStatusLine([MarshalAs(UnmanagedType.BStr)] string bstrStatusLine);
        void ResetStatusLine();
    }

    #endregion

}
