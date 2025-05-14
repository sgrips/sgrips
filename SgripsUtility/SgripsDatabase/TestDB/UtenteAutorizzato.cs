using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BRT.BartMailSync.Model
{
    public class UtenteAutorizzato : System.ComponentModel.INotifyPropertyChanged, IDataErrorInfo
    {

        private int id;

        public int Id
        {
            get { return id; }
            set { id = value; this.RaisePropertyChanged("Id"); }
        }

        private string utente;

        public string Utente
        {
            get { return utente; }
            set { utente = value; this.RaisePropertyChanged("Utente"); }
        }

        private bool admin;

        public bool Admin
        {
            get { return admin; }
            set { admin = value; this.RaisePropertyChanged("Admin"); }
        }


        #region validation rules    
        private  Dictionary<string, string> _errors = new Dictionary<string, string>();
        public string Error => (_errors.Count==0  ? "" : "In Errore" ) ;

        private void ValidaCampi(string columnName)
        {
            if (_errors.ContainsKey (columnName )) _errors.Remove(columnName);

            switch (columnName)
            {
                case "Utente":
                    if (string.IsNullOrEmpty(this.Utente))
                        _errors.Add(columnName, "Valore Obbligatorio");
                    break;
            }

        }

        public string this[string columnName]
        {
            get
            {
                ValidaCampi(columnName);

                if (_errors.ContainsKey(columnName)) return _errors[columnName];

                return string.Empty;
            }
        }


        #endregion


        #region propertyChange
        public event PropertyChangedEventHandler PropertyChanged;

        protected void RaisePropertyChanged(string propertyName)
        {
            System.ComponentModel.PropertyChangedEventHandler propertyChanged = this.PropertyChanged;
            if ((propertyChanged != null))
            {
                propertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

    }
}
