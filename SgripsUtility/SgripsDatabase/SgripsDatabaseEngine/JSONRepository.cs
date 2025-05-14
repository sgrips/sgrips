using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;


namespace Sgrips.Framework.FileSystemDatabase
{
    public class JSONRepository : FileSystemRepository

    {

        private JArray _recordset;
        private bool mustRefresh = false;
        private FileSystemWatcher fsw = null;

        private JSONRepository() { }

        public JSONRepository(string database, string id = "Id", bool encrypt = false) : this(null, database, id, encrypt)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <param name="database"></param>
        /// <param name="id"></param>
        /// <param name="encrypt"></param>
        /// <param name="encryptFields"></param>
        public JSONRepository(string path, string database, string id = "Id", bool encrypt = false) : base(path, database, id, encrypt)
        {
            //  if (!System.IO.File.Exists(System.IO.Path.Combine(RootPath, DataBaseName)))
            //    File.CreateText ( System.IO.Path.Combine(RootPath, DataBaseName));

            fsw = new FileSystemWatcher(RootPath, DataBaseName);
            fsw.IncludeSubdirectories = false;
            fsw.EnableRaisingEvents = true;
            fsw.Changed += new FileSystemEventHandler((a, b) =>
            {
                mustRefresh = true;
            });
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        protected override void Write<T>(IList<T> rows)
        {
            _recordset = JArray.FromObject(EncryptDescryptProperties(rows));
            WriteRecordset();

        }


        /// <summary>
        /// ritorna la lista di tipo T serializzata su di un file JSON
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        protected override IList<T> Read<T>()
        {
            if (Recordset == null) return default(IList<T>);

            Type typeParameterType = typeof(IList<T>);

            return EncryptDescryptProperties(Recordset.ToObject(typeParameterType) as IList<T>, true);
        }



        /// <summary>
        /// Legge da un file JSON l'elemento con IDValue passato , tornando l'oggetto T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="idValue"></param>
        /// <returns></returns>
        protected override T Read<T>(string idValue)
        {

            JToken jtoken = FindEntity(idValue);

            return (jtoken == null) ? default(T) : EncryptDescryptProperties(jtoken.ToObject<T>(), true);
        }



        /// <summary>
        /// Aggiunge un elemento al Recordset e lo serializza sul file testo in formato JSON
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <exception cref="DuplicateIDException"></exception>
        protected override void Append<T>(T row)
        {
            if (Recordset == null) return;

            if (FindEntity(row) != null) throw new DuplicateIDException();


            JToken tokenToAdd = JToken.Parse(JsonConvert.SerializeObject(EncryptDescryptProperties(row), Formatting.Indented));

            Recordset.Add(tokenToAdd);

            WriteRecordset();
        }

        /// <summary>
        /// Cancella il record memorizzato su file in formato JSON con chiave KeyID, che contiene il nome della proprietà
        /// nel oggetto passato che rappresenta la chiave primaria es. Id
        /// classe JSONRepository.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        protected override bool Delete<T>(T row)
        {

            JObject jo = (JObject)FindEntity(row);
            if (jo != null)
            {
                Recordset.Remove(jo);
                WriteRecordset();
                return true;
            }


            return false;
        }


        /// <summary>
        /// Aggiorna il record memorizzato su file in formato JSON con chiave KeyID, che contiene il nome della proprietà
        /// nel oggetto passato che rappresenta la chiave primaria es. Id
        /// classe JSONRepository.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        protected override bool Update<T>(T row)
        {
            Type type = typeof(T);
            JObject jo = (JObject)FindEntity(row);
            if (jo != null)
            {
                Recordset.Remove(jo);
                Recordset.Add(JObject.FromObject(EncryptDescryptProperties(row)));
                WriteRecordset();
                return true;
            }
            return false;
        }


        /// <summary>
        /// Serializza su file testo in formato JSON il JArray rappresentante il recordset dati
        /// </summary>
        private void WriteRecordset()
        {

            fsw.EnableRaisingEvents = false;

            using (StringWriter sw = new StringWriter())
            {
                using (StreamWriter sr = new StreamWriter(File.Open(FullDataBaseName,
                                        FileMode.Create,
                                        FileAccess.Write,
                                        FileShare.ReadWrite)))
                {
                    using (var jw = new JsonTextWriter(sw))
                    {

                        JsonSerializer js = new JsonSerializer();

                        js.Formatting = Formatting.Indented;
                        fsw.EnableRaisingEvents = true;
                        js.Serialize(sw, Recordset);

                        string jsonString = sw.ToString();

                        if (this.EncryptDB)
                            jsonString = EncryptString.Encrypt(jsonString);

                        sr.Write(jsonString);
                        sr.Flush();
                    }
                }
            }



        }

        /// <summary>
        /// Cerca il record memorizzato su file in formato JSON con chiave KeyID, che contiene il nome della proprietà
        /// nel oggetto passato che rappresenta la chiave primaria es. Id
        /// classe JSONRepository.
        /// </summary>
        /// <param name="idValue"></param>
        /// <exception cref="UniqueIDNotFoundException"></exception>
        /// <returns ></returns>
        private JToken FindEntity(object id)
        {
            object idValue = null;
            if (!IsPrimitiveType.Test(id.GetType()))
            {
                Type type = id.GetType();

                PropertyInfo p = type.GetProperties().FirstOrDefault(f => f.Name == base.KeyID);
                if (p != null)
                    idValue = p.GetValue(id);
                else throw new UniqueIDNotFoundException();
            }
            else idValue = id;


            return Recordset.SelectToken(string.Format("$..*.[?(@.{0}=={1})]",
                    base.KeyID, idValue.GetType() == typeof(String) ? "'" + idValue + "'" : idValue));
        }


        private Object recordsetLock = new Object();
        /// <summary>
        /// De-serializza l'oggetto JSON in JArry
        /// </summary>
        protected internal JArray Recordset
        {
            get
            {
                lock (recordsetLock)
                {
                    if (_recordset == null || mustRefresh)
                    {

                        mustRefresh = false;
                        if (!System.IO.File.Exists(FullDataBaseName)) _recordset = null;
                        else
                        {
                            using (StreamReader sr = new StreamReader(File.Open(FullDataBaseName,
                                                           FileMode.Open,
                                                           FileAccess.Read,
                                                           FileShare.Read)))

                            {

                                string jsonString = sr.ReadToEnd();

                                if (this.EncryptDB)
                                    jsonString = EncryptString.Decrypt(jsonString);

                                using (JsonTextReader reader = new JsonTextReader(new StringReader(jsonString)))
                                    _recordset = (JArray)JToken.ReadFrom(reader);

                            }
                        }
                    }

                    return _recordset;
                }

            }
        }

    }
}
