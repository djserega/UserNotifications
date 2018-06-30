using System;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;

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
    void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName);
    void GetNProps(ref Int32 props);
    void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum);
    void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName);
    void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
    void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal);
    void IsPropReadable(Int32 propNum, ref bool propRead);
    void IsPropWritable(Int32 propNum, ref Boolean propWrite);
    void GetNMethods(ref Int32 pMethods);
    void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum);
    void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName);
    void GetNParams(Int32 methodNum, ref Int32 pParams);
    void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue);
    void HasRetVal(Int32 methodNum, ref Boolean retValue);
    void CallAsProc(Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
    void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
}

/// <summary>Интерфейс реализован 1С для получения событий от компоненты</summary>
/// 
[Guid("ab634004-f13d-11d0-a459-004095e1daea")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAsyncEvent
{
    void SetEventBufferDepth(Int32 depth);
    void GetEventBufferDepth(ref long depth);
    void ExternalEvent([MarshalAs(UnmanagedType.BStr)] String source, [MarshalAs(UnmanagedType.BStr)] String message, [MarshalAs(UnmanagedType.BStr)] String data);
    void CleanBuffer();
}

/// <summary>С помощью этого интерфейса компонента получает доступ к строке состояния 1С</summary>
/// 
[Guid("AB634005-F13D-11D0-A459-004095E1DAEA")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IStatusLine
{
    void SetStarusLine([MarshalAs(UnmanagedType.BStr)] String bstrStatusLine);
    void ResetStatusLine();
}

namespace AddIn
{
    /// <summary>Класс, реализующий свойства и методы для подключения внешней компоненты к 1С</summary>
    public class AddIn : IInitDone, ILanguageExtender
    {
        /// <summary>ProgID COM-объекта компоненты</summary>
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
        public void Done()
        {

        }

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
        {
            info[0] = 2000;
        }

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
        public void GetNProps(ref Int32 props)
        {
            props = (Int32)propertyNameToNumber.Count;
        }

        /// <summary>Возвращает целочисленный идентификатор свойства, соответствующий переданному имени</summary>
        /// <param name="propName"></param>
        /// <param name="propNum"></param>
        public void FindProp([MarshalAs(UnmanagedType.BStr)] String propName, ref Int32 propNum)
        {
            propNum = (Int32)propertyNameToNumber[propName];
        }

        /// <summary>Возвращает имя свойства, соответствующее переданному целочисленному идентификатору</summary>
        /// <param name="propNum"></param>
        /// <param name="propAlias"></param>
        /// <param name="propName"></param>
        public void GetPropName(Int32 propNum, Int32 propAlias, [MarshalAs(UnmanagedType.BStr)] ref String propName)
        {
            propName = (String)propertyNumberToName[propNum];
        }

        /// <summary>Возвращает значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propVal"></param>
        public void GetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            propVal = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].GetValue(this, null);
        }

        /// <summary>Устанавливает значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propVal"></param>
        public void SetPropVal(Int32 propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal, null);
        }

        /// <summary>Определяет, можно ли читать значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propRead"></param>
        public void IsPropReadable(Int32 propNum, ref bool propRead)
        {
            propRead = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanRead;
        }

        /// <summary>Определяет, можно ли изменять значение свойства</summary>
        /// <param name="propNum"></param>
        /// <param name="propWrite"></param>
        public void IsPropWritable(Int32 propNum, ref Boolean propWrite)
        {
            propWrite = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanWrite;
        }

        /// <summary>Возвращает количество методов</summary>
        /// <param name="pMethods"></param>
        public void GetNMethods(ref Int32 pMethods)
        {
            pMethods = (Int32)nameToNumber.Count;
        }

        /// <summary>Возвращает идентификатор метода по его имени</summary>
        /// <param name="methodName">Имя метода</param>
        /// <param name="methodNum">Идентификатор метода</param>
        public void FindMethod([MarshalAs(UnmanagedType.BStr)] String methodName, ref Int32 methodNum)
        {
            methodNum = (Int32)nameToNumber[methodName];
        }

        /// <summary>Возвращает имя метода по идентификатору</summary>
        /// <param name="methodNum"></param>
        /// <param name="methodAlias"></param>
        /// <param name="methodName"></param>
        public void GetMethodName(Int32 methodNum, Int32 methodAlias, [MarshalAs(UnmanagedType.BStr)] ref String methodName)
        {
            methodName = (String)numberToName[methodNum];
        }

        /// <summary>Возвращает число параметров метода по его идентификатору</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Число параметров</param>
        public void GetNParams(Int32 methodNum, ref Int32 pParams)
        {
            pParams = (Int32)numberToParams[methodNum];
        }

        /// <summary>Получить значение параметра метода по умолчанию</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="paramNum">Номер параметра</param>
        /// <param name="paramDefValue">Возвращаемое значение</param>
        public void GetParamDefValue(Int32 methodNum, Int32 paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue) { }

        /// <summary>Указывает, что у метода есть возвращаемое значение</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Наличие возвращаемого значения</param>
        public void HasRetVal(Int32 methodNum, ref Boolean retValue)
        {
            retValue = (Boolean)numberToRetVal[methodNum];
        }

        /// <summary>Вызов метода как процедуры с использованием идентификатора</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Параметры</param>
        public void CallAsProc(Int32 methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                asyncEvent.ExternalEvent(AddIn.AddInName, e.Message, e.ToString());
            }
        }

        /// <summary>Вызов метода как функции с использованием идентификатора</summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Возвращаемое значение</param>
        /// <param name="pParams">Параметры</param>
        public void CallAsFunc(Int32 methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                asyncEvent.ExternalEvent(AddIn.AddInName, e.Message, e.ToString());
            }
        }
    }
}
