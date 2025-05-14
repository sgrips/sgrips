using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Sgrips.Lumia.Import
{
    class Program
    {
        const String path_pictures = @"D:\antonio\Pictures\2016-03";
        const string path_pictures_delete = @"D:\temp\Delete";

        static void Main(string[] args)
        {
            Console.WriteLine("Ricerca Foto Duplicate");
            Console.WriteLine("1 - Ricerca e Conta");
            Console.WriteLine("2 - Ricerca e Cancella");


            FindDuplicateFiles(Console.ReadKey().KeyChar == '1');

            Console.WriteLine();
            Console.WriteLine("Premere un Tasto per uscire");
            Console.ReadLine();

        }

        private static void FindDuplicateFiles(bool conta)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(path_pictures);
            var files = (from f in dirInfo.GetFiles("*.*", SearchOption.AllDirectories)
                         group f by new { name = removeDuplicatorChars(f.Name), len = f.Length } into grpNames
                         let values = new { grp = grpNames, count = grpNames.Count() }
                         where values.count > 1
                         select values.grp);


            Console.WriteLine();
            Console.WriteLine("File Doppi {0}", files.Count());
            List<string> originals = new List<string>();
            ConsoleColor lastConsoleColor;
            files.ToList().ForEach(f => f.ToList().ForEach(
                f1 =>
                {
                    //se sono presenti solo file con i caratteri specifici che identificano un file duplicato , ma non è presente il file originale, 
                    //devo rinominare uno dei file duplicati con il nome del file originale.

                    if (isMatchWithDuplicatorChars(f1.Name))
                    {

                        string nomeFile = removeDuplicatorChars(f1.Name);
                        bool createOrig = false;
                        if (!originals.Contains (nomeFile ) && !System.IO.File.Exists(System.IO.Path.Combine(f1.DirectoryName, nomeFile)))
                        {
                            if (!conta)
                            {
                                System.IO.File.Move(f1.FullName, System.IO.Path.Combine(f1.DirectoryName, nomeFile));
                                return ;
                            }
                            else { 
                                createOrig = true;
                                originals.Add(nomeFile);
                            }
                        }


                        if (conta) {
                            lastConsoleColor = Console.ForegroundColor;
                            Console.ForegroundColor = createOrig ? ConsoleColor.Red : Console.ForegroundColor  ;
                            Console.WriteLine(createOrig ? f1.Name + " Creato Originale " : f1.Name + " Da Cancellare");
                            Console.ForegroundColor = lastConsoleColor;
                        }
                        else
                        {

                            string pathTo = f1.DirectoryName.Replace(path_pictures, path_pictures_delete);
                            if (!System.IO.Directory.Exists(pathTo)) System.IO.Directory.CreateDirectory(pathTo);
                            Console.WriteLine("Spostato {0} in {1}", f1.Name, f1.FullName.Replace(path_pictures, path_pictures_delete));

                            File.Move(f1.FullName, f1.FullName.Replace(path_pictures, path_pictures_delete));
                        }
                    }
                }



                ));
        }

        static Regex rex = new Regex(@"^(?<nomefile>[a-zA-Z0-9_-]+)(\s{0,3}\([0-9]\))(?<estensione>\.\w{3})$");
        private static bool isMatchWithDuplicatorChars(string check)
        {
            return rex.IsMatch(check);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string removeDuplicatorChars(string name)
        {
            return rex.Replace(name, "${nomefile}${estensione}");
        }
    }
}
