using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Sgrips.Framework.FileSystemDatabase
{
    public class JSONDataReader : ICursor
    {
        JSONRepository _jrepository;

        private JArray _recordset;
        private Dictionary<string, object> _arrFields = new Dictionary<string, object>();
        private Int32 _currentPosition;
        private Int32 _fieldCount;

        /// <summary>
        /// Ritorna il numero di colonne legate al recordset corrente
        /// </summary>
        public int FieldCount
        {
            get
            {
                return _fieldCount;
            }
        }

        /// <summary>
        /// Rendo privato il costruttore senza argomenti così non renderlo disponibile
        /// </summary>
        private JSONDataReader() { }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="jrepository">passare l'istanza dell' oggetto JSONRepository che rappresenta una IList<T>
        /// dove T è un oggetto, che può essere un tipo standard o utente es. Class People</param>
        public JSONDataReader(JSONRepository jrepository)
        {
            this._jrepository = jrepository;
            this._recordset = this._jrepository.Recordset;

            _currentPosition = this._recordset.Count() > 0 ? 0 : -1;
            if (_currentPosition != -1)
                this._fieldCount = (_recordset.ElementAt(_currentPosition) as JObject).Properties().Count();

        }


        /// <summary>
        /// ritorna il valore della colonna individuata dal indice passato
        /// </summary>
        /// <param name="index">indice numerico</param>
        /// <returns></returns>
        public object this[int index]
        {
            get
            {
                return _arrFields.ElementAt(index).Value;
            }

            set
            {
                KeyValuePair<string, object> element = _arrFields.ElementAt(index);
                _arrFields[element.Key] = value;
            }
        }

        /// <summary>
        /// ritorna il valore della colonna individuata dal indice passato
        /// </summary>
        /// <param name="index">indice è di tipo stringa</param>
        /// <returns></returns>
        public object this[string index]
        {
            get
            {
                return _arrFields[index];
            }

            set
            {
                _arrFields[index] = value;
            }
        }

        /// <summary>
        /// ritorna il valore della colonna individuata dal indice passato
        /// </summary>
        /// <param name="index">indice numerico</param>
        /// <returns></returns>
        public object GetValue(int index)
        {
            return this[index];
        }

        /// <summary>
        /// ritorna il nome della colonna individuata dal indice passato
        /// </summary>
        /// <param name="index">indice numerico</param>
        /// <returns></returns>
        public string GetName(int index)
        {
            KeyValuePair<string, object> element = _arrFields.ElementAt(index);
            return element.Key;
        }


        /// <summary>
        /// si muove sul cursore di sola lettura, una volta raggiunta la fine viene ritornato false.
        /// </summary>
        /// <returns></returns>
        public bool Read()
        {

            _arrFields.Clear();

            if (_currentPosition < this._recordset.Count() && _currentPosition != -1)
            {
                try
                {
                    JObject jobject = _recordset.ElementAt(_currentPosition) as JObject;
                    foreach (JProperty prop in jobject.Properties())
                        _arrFields.Add(prop.Name, prop.Value);

                    _currentPosition++;

                    return true;
                }
                catch (Exception ex)
                {
                    throw;

                }

            }
            else { _currentPosition = -1; return false; }
        }
    }
}
