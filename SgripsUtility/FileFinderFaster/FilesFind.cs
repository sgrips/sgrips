using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BRT.Framework.FileFinderFaster.UI
{
    public enum FileCompareEnum
    {
        Like,
        Equal
    }
    class FilesFind
    {
        //Percorso in cui è stato trovato il file
        public string PathNameFound { get; set; }
        //Nome del file reale trovato nel caso di LIKE viene preso il file name trovato in PathFileSearch
        public string RealFileNameFound { get; set; }
        //Percorso in cui cercare il file
        public string PathFileSearch { get; set; }
        //Nome del file da cercare
        public string FileNameSearch { get; set; }
        //Percorso in cui si strovano i file da cercare
        public string PathFileSource { get; set; }
        //metodo da usare per verificare se presente il FileNameSearch in PathFileSearch
        public FileCompareEnum FileCompare  { get; set; }
        //parte del file rimossa quando trovato | nel nome del file
        public string PartFileNameRemove { get; set; }

}
}
