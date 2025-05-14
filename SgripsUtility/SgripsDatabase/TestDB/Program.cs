using BRT.BartMailSync.Model;
using Sgrips.Framework.FileSystemDatabase;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;

namespace TestDB
{


    class Cliente
    {
        public int ID { get; set; }
        public string Nome { get; set; }
        public string Cognome { get; set; }
        public string UserID { get; set; }
        [FieldAttributes(Encrypt = true)]
        public string Password { get; set; }

        public List<OrdineCliente> OrdineCliente { get; set; }

        public Cliente()
        {
            OrdineCliente = new List<OrdineCliente>();
        }


        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(
            "id:{0}, nome:{1}, cognome:{2}, userID:{3}, userPWD:{4}",
                ID, Nome, Cognome, UserID, Password).AppendLine();
            foreach (var item in OrdineCliente)
            {
                sb.Append("\t").AppendLine(item.ToString());
            }
            return sb.ToString();
        }
    }


    class OrdineCliente
    {
        public int IDOrdine { get; set; }
        public int IDCLiente { get; set; }
        public string CodProdotto { get; set; }
        public string Prezzo { get; set; }
        [FieldAttributes(Encrypt = true)]
        public string CodScontoSegreto { get; set; }

        public override string ToString()
        {
            return string.Format("id:{0}, idCliente:{1}, codProdotto:{2}, Prezzo:{3}, codScontoSegreto:{4}",
                IDOrdine, IDCLiente, CodProdotto, Prezzo, CodScontoSegreto);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ObservableCollection<UtenteAutorizzato> utentiAutorizzati;

            var _utentiRepository = new JSONRepository(Environment.CurrentDirectory, "BartMailUser", "Utente", true);
            //delete all data
            if (!System.IO.File.Exists(_utentiRepository.FullDataBaseName))
            {
                File.Delete(_utentiRepository.FullDataBaseName);
                utentiAutorizzati = new ObservableCollection<UtenteAutorizzato>();
                utentiAutorizzati.Add(new UtenteAutorizzato() { Utente = "antonio.sgrignoli", Admin = true, Id = 1 });
                _utentiRepository.Save(utentiAutorizzati);
            }


            utentiAutorizzati = new ObservableCollection<UtenteAutorizzato>((IList<UtenteAutorizzato>)_utentiRepository.GetAll<UtenteAutorizzato>());











            JSONRepository jr = new JSONRepository(@"c:\temp\", "JSONDB", "ID",true);

            List<Cliente> clienti = new List<Cliente>() { new Cliente() { ID = 5, Nome = "Antonio", Cognome = "Piccoli", UserID = "pippo", Password = null } };


            //delete all data
            if (System.IO.File.Exists("c:\\temp\\JSONDB.json"))
                System.IO.File.Delete("c:\\temp\\JSONDB.json");

            //crea database
            jr.Save(clienti);

            jr.Insert(new Cliente()
            {
                ID = 1,
                Nome = "Antonio",
                Cognome = "Piccoli",
                UserID = "pippo",
                Password = "password1"
                       ,
                OrdineCliente = new List<OrdineCliente>() { new OrdineCliente { IDOrdine = 1, IDCLiente = 1, CodProdotto = "AAA222", Prezzo = "10.55", CodScontoSegreto = "SGRIPS" } }
            });
            jr.Insert(new Cliente() { ID = 2, Nome = "Andrea", Cognome = "Caro" });

            foreach (var item in jr.GetAll<Cliente>())
                Console.WriteLine(item);

            Console.WriteLine( jr.Get<Cliente>("4") == null ? "Non Trovato" : "Trovato");

            Console.WriteLine();
            Console.WriteLine("Lettura con DataReader, premere un tasto");
            Console.WriteLine();
            Console.ReadLine();

            JSONDataReader jDataReader = new JSONDataReader(jr);



            while (jDataReader.Read())
            {
                Console.WriteLine("id:{0}, nome:{1}, cognome:{2}, userID:{3}, password:{4}", jDataReader[0], jDataReader[1],
                    jDataReader[2], jDataReader[3], jDataReader[4]);
            }




            Console.ReadLine();
            Console.WriteLine("Cancello il record con ID=2");
            jr.Remove(new Cliente() { ID = 2 });

            jDataReader = new JSONDataReader(jr);
            while (jDataReader.Read())
            {
                Console.WriteLine("id:{0}, nome:{1}, cognome:{2}, userID:{3}, password:{4}", jDataReader[0], jDataReader[1], jDataReader[2], jDataReader[3], jDataReader[4]);
            }

            Console.ReadLine();

        }
    }
}
