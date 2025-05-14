using Sgrips.Framework.FileSystemDatabase;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BRT.BartMailSync.Model
{
    class UtentiAutorizzati
    {
        ObservableCollection<UtenteAutorizzato> utentiAutorizzati = null;
        private FileSystemRepository _utentiRepository;

        static UtentiAutorizzati utenti = null;

        private UtentiAutorizzati()
        {
            _utentiRepository = new JSONRepository(Environment.CurrentDirectory, "BartMailUser", "Utente", true);
            //delete all data
            if (!System.IO.File.Exists(_utentiRepository.FullDataBaseName))
            {
                File.Delete(_utentiRepository.FullDataBaseName);
                utentiAutorizzati = new ObservableCollection<UtenteAutorizzato>();
                Add(new UtenteAutorizzato() { Utente = "antonio.sgrignoli", Admin = true, Id = 1 });
                _utentiRepository.Save(utentiAutorizzati);
            }


            utentiAutorizzati = new ObservableCollection<UtenteAutorizzato>((IList<UtenteAutorizzato>)_utentiRepository.GetAll<UtenteAutorizzato>());

        }

        public static UtentiAutorizzati Instance()
        {
            if (utenti == null)
            {
                utenti = new UtentiAutorizzati();
            }

            return utenti;
        }

        internal void Delete(UtenteAutorizzato  utenteAutorizzato )
        {

            _utentiRepository.Remove<UtenteAutorizzato>(utenteAutorizzato);
        }

        #region Methods
        public void Save()
        {
            _utentiRepository.Save (utentiAutorizzati);
        }


        public void Add(UtenteAutorizzato utente)
        {
            utentiAutorizzati.Add(utente);
            
        }

        public bool IsValid()
        {
            return utentiAutorizzati.All (f => f.Error == "");
        }

        public bool IsAdmin(string utente)
        {
            return utentiAutorizzati.Any(f => f.Utente.ToLower() == utente.ToLower() && f.Admin == true);
        }

        public bool IsAuthenticate(string utente)
        {
            return utentiAutorizzati.Any (f => f.Utente.ToLower() == utente.ToLower());
        }

        public ObservableCollection<UtenteAutorizzato> GetAll()
        {
            return utentiAutorizzati;
        }

        #endregion




    }
}
