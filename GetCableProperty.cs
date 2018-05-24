using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.DataModel;
using System;
using System.Text.RegularExpressions;

namespace Eplan.EplAddin.PlaceholderAction
{
    /// <summary>
    /// Фильтрует поле "Трасса маршрутизации" в отчете по кабелю для формирования таблицы соединений
    /// </summary>
    public class AddInModule : IEplAddIn
    {
        /// <summary>
        /// The function is called once during registration add-in.
        /// </summary>
        /// <param name="bLoadOnStart"> true: In the next P8 session, add-in will be loaded during initialization</param>
        /// <returns></returns>
        public bool OnRegister(ref System.Boolean bLoadOnStart)
        {
            bLoadOnStart = true;

            return true;
        }
        /// <summary>
        /// The function is called during unregistration the add-in.
        /// </summary>
        /// <returns></returns>
        public bool OnUnregister()
        {
            return true;
        }

        /// <summary>
        /// The function is called during P8 initialization or registration the add-in. 
        /// </summary>
        /// <returns></returns>
        public bool OnInit()
        {

            return true;

        }
        /// <summary>
        /// The function is called during P8 initialization or registration the add-in, when GUI was already initialized and add-in can modify it.
        /// </summary>
        /// <returns></returns>
        public bool OnInitGui()
        {
            return true;

        }
        /// <summary>
        /// This function is called during closing P8 or unregistration the add-in.
        /// </summary>
        /// <returns></returns>
        public bool OnExit()
        {
            return true;
        }

    }

        class GetCableProperty : IEplAction
    {
        // Выводим в отчет только участки трассы кабеля вида К1, К13, и т.д. Всё остальное не надо.
        private static Regex regex = new Regex("K[0-9]*");
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            // Записать Name - Eplan не подскажет...
            Name = "GetCableProperty";
            Ordinal = 20;
            return true;
        }
        // Поехали!!!
        public bool Execute(ActionCallingContext oActionCallingContext)
        {
            string objectNames = "";
            oActionCallingContext.GetParameter("objects", ref objectNames);
            // Получим объект текущей строки. Конечно, объектов может быт несколько, но не в этом отчете...
            StorableObject cable = StorableObject.FromStringIdentifier(objectNames);
            // Получим свойство 20237 "Топология: Трасса маршрутизации"
            string sCABLING_PATH = cable.Properties[20237];

            MatchCollection matches = regex.Matches(sCABLING_PATH);
            String filteredPath = "";

            var enumerator = matches.GetEnumerator();
            if (enumerator.MoveNext())
            {
                filteredPath = ((Match)enumerator.Current).Value;
                while (enumerator.MoveNext())
                {
                    filteredPath += ";" + (Match)enumerator.Current;
                }
            }

            string[] strings = new string[1];
            strings[0] = filteredPath;
            oActionCallingContext.SetStrings(strings);
            return true;
        }
        // Obsolete - игнорируем.
        public void GetActionProperties(ref ActionProperties actionProperties)
        {
            return;
        }

    }
}
