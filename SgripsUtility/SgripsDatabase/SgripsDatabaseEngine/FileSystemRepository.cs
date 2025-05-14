using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace Sgrips.Framework.FileSystemDatabase
{
    public abstract class FileSystemRepository
    {

        private string _rootPath;


        #region gestione collezioni
        /* i dato stanno all'interno di una collezione */
        protected abstract IList<T> Read<T>();
        protected abstract void Write<T>(IList<T> rows);
        protected abstract void Append<T>(T row);
        protected abstract bool Update<T>(T row);
        protected abstract bool Delete<T>(T row);
        protected abstract T Read<T>(string idValue);
        #endregion

        #region constructors
        protected FileSystemRepository() { }

        public FileSystemRepository(string DataBaseName, string KeyID = "Id", bool EncryptDB = false, string[] EncryptFields = null) :
            this(null, DataBaseName, KeyID)
        { }

        public FileSystemRepository(string PathDataBase, string DataBaseName, string KeyID = "Id", bool EncryptDB = false)
        {
            this.EncryptDB = EncryptDB;
            this.DataBaseName = DataBaseName.ToUpper().IndexOf(".JSON") == -1 ? DataBaseName + ".json" : DataBaseName;
            this.RootPath = PathDataBase == null ? "DEFAULT" : PathDataBase;
            this.KeyID = KeyID;
        }
        #endregion

        #region properties
        public string DataBaseName { get; private set; }
        public string KeyID { get; private set; }
        public bool EncryptDB { get; private set; }

        public String FullDataBaseName
        {
            get
            {
                if (this.RootPath != null && DataBaseName != null)
                    return System.IO.Path.Combine(this.RootPath, DataBaseName);
                else
                    return null;
            }
        }

        public string RootPath
        {
            get
            {
                if (_rootPath == "DEFAULT")
                {
                    _rootPath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        Company(), "SgripsDataBase");
                }
                return _rootPath;
            }
            set { _rootPath = value; }
        }

        #endregion

        #region encrypproperties
        /// <summary>
        /// cripta le propietà indicate in EncryptProperties es. usare quando ci sono dati sensibili che non si vuole 
        /// mettere in chiaro
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="records"></param>
        /// <returns></returns>
        protected T EncryptDescryptProperties<T>(T records, bool decrypt = false)
        {
            if (records is IList)
                foreach (var item in records as IList)
                    EncryptDescryptProperties(item, decrypt);
            else
            {
                (from p in records.GetType().GetProperties()
                 where p.GetCustomAttributes(typeof(FileSystemDatabase.FieldAttributes), true).Length > 0 &&
                  (p.GetCustomAttributes(typeof(FileSystemDatabase.FieldAttributes), true)[0] as FileSystemDatabase.FieldAttributes).Encrypt == true
                 select p).ToList().ForEach(x =>
                 {
                     if (x.GetValue(records) != null)
                     {
                         if (IsPrimitiveType.Test(x.PropertyType))
                             x.SetValue(records, decrypt ? EncryptString.Decrypt((string)x.GetValue(records))
                                : EncryptString.Encrypt((string)x.GetValue(records)));
                         else
                         {
                             IList valueList = x.GetValue(records) as IList;
                             for (var i = 0; i < valueList.Count; i++)
                                 valueList[i] = decrypt ? EncryptString.Decrypt((string)valueList[i])
                                        : EncryptString.Encrypt((string)valueList[i]);

                         }
                     }
                 });


                Func<PropertyInfo, object, bool> isList =
                    (x, y) =>
                    {
                        try
                        {
                            return (x.GetValue(y) is IList);
                        }
                        catch (Exception)
                        {
                            return false;
                        }
                    };

                (from p in records.GetType().GetProperties()
                 where isList(p,records )
                 select p).ToList().ForEach(f => EncryptDescryptProperties(f.GetValue(records), decrypt));

            }

            return records;
        }
        #endregion

        private string Company()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            return FileVersionInfo.GetVersionInfo(assembly.Location).CompanyName;
        }

        public void Save<T>(IList<T> rows)
        {
            Write<T>(rows);
        }

        public void Modify<T>(T row)
        {
            Update<T>(row);
        }

        public void Insert<T>(T row)
        {
            Append<T>(row);
        }

        public void Remove<T>(T row)
        {
            Delete<T>(row);
        }

        public T Get<T>(string idValue) where T : class
        {

            return Read<T>(idValue);
        }

        public IList<T> GetAll<T>()
        {

            return Read<T>();
        }

    }
}
