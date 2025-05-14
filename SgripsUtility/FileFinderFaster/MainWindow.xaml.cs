using BRT.Framework.FileFinderFaster.UI;
using Microsoft.Windows.Controls.Ribbon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace FileFinderFaster
{



    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// ricerca un file passato in un percorso stabilito, altrimenti è anche possibile utilizzare un file
    /// con l'elenco dei file da cercare in un dettermintato percorso, il file deve avere la seguente struttura
    /// @\\bartvas3\BARTVAS3\BartFTP\1662139\OLDIN
    /// [d:\temp\]2018-11-22-051712.BARTOLINI.1027054175603002.MXP5.xml[|]20181122142626.edifactBRT[|][#Like]
    /// [percorso sorgente]file2.pippo[|]_20180911[|].xml[#Like] 
    /// [percorso sorgente]fileN.XXXXX[|]YYYY[|]ZZZZ[#Like] 
    ///
    /// se indicato [percorso sorgente] es. [d:\temp\] i file presenti nel percorso se trovati nel destinatario sono spostati nella
    /// cartella [percorso sorgente/found]
    /// se indicato [|] nel file da cerare, deve essere presente 2 volte ed indica la porzione di file che deve essere
    /// esclusa dalla ricerca
    /// se indicato [#Like] permette di definire il metodo con cui fare la ricerca del file nel percorso indicato con [@percorso]
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        BackgroundWorker bw;
        StateFinder fileStateFinder;
        List<FilesFind> _fileFinds = new List<FilesFind>();
        private string _fileNameImported;

        /// <summary>
        /// 
        /// </summary>
        public MainWindow()
        {

            InitializeComponent();

            txtNomeFile.Text = "FNVAC20150109110000_20150109110002.ZIP";
            txtPercorsoRicerca.Text = @"\\cedstorebo\BartBackup\Bartvas\20150109_01\";
            fileStateFinder = (StateFinder)this.Resources["StateFinder"];
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RicercaButton_Click(object sender, RoutedEventArgs e)
        {
            bw = new BackgroundWorker();
            bw.WorkerSupportsCancellation = true;
            bw.WorkerReportsProgress = false;

            bw.DoWork += bw_DoWork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;


            RicercaButton.IsEnabled = false;

            if (_fileFinds.Count == 0)
            {
                _fileNameImported = "";
                _fileFinds.Add(new FilesFind() { FileNameSearch = txtNomeFile.Text.ToUpper(), PathNameFound = "", PathFileSearch = txtPercorsoRicerca.Text });
            }


            bw.RunWorkerAsync(_fileFinds);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            RicercaButton.IsEnabled = true;

            ///
            if (e.Result != null && !string.IsNullOrEmpty(_fileNameImported))
            {
                List<FilesFind> fileTofind = (List<FilesFind>)e.Result;


                if (fileTofind.Count > 0)
                {
                    using (FileStream fs = new FileStream(_fileNameImported + ".csv", FileMode.Create))
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("FILE_SEARCH").Append(";");
                        sb.Append("METHOD_SEARCH").Append(";");
                        sb.Append("FILE_FOUND").Append(";");
                        sb.AppendLine("PATH_NAME_FOUND");
                        foreach (var item in fileTofind)
                        {
                            sb.AppendLine(string.Format("{0};{1};{2};{3}", item.FileNameSearch, item.FileCompare , item.RealFileNameFound , item.PathNameFound));
                            //se presente il percorso PathFileSource sposto i file trovati in PathFileSource/found
                            if (!string.IsNullOrWhiteSpace(item.PathFileSource) &&
                                !string.IsNullOrWhiteSpace(item.PathNameFound))
                            {
                                //se percorso non presente lo creo
                                if (!Directory.Exists(Path.Combine(item.PathFileSource, "Found")))
                                    Directory.CreateDirectory(Path.Combine(item.PathFileSource, "Found"));

                                //se file name non trovato provo a cercare per pattern nomefile*
                                string filename = Path.Combine(item.PathFileSource, item.FileNameSearch+item.PartFileNameRemove);

                                //if (!File.Exists(filename))
                                //filename = Directory.GetFiles(item.PathFileSource, item.FileNameSearch + "*").First();

                                File.Move(filename,
                                    Path.Combine(item.PathFileSource, @"Found\" + Path.GetFileName(filename)));

                            }
                        }

                        byte[] arr = UTF8Encoding.UTF8.GetBytes(sb.ToString());
                        fs.Write(arr, 0, arr.Length);
                        fs.Flush();
                        fs.Close();
                    }
                }
            }

            _fileFinds.Clear();
            MessageBox.Show("Ricerca terminata", "Cerca File", MessageBoxButton.OK);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="typeHolder"></param>
        /// <param name="x"></param>
        /// <returns></returns>
        private static T Cast<T>(T typeHolder, Object x)
        {
            // typeHolder above is just for compiler magic
            // to infer the type to cast x to
            return (T)x;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            Func<DirectoryInfo, List<FilesFind>, bool> _RicercaFile = (d, filesFind) =>
          {
              if (filesFind.Count > 0)
                  d.EnumerateFiles("*.*", SearchOption.TopDirectoryOnly).AsParallel().ForAll(f =>
                  {
                      try
                      {
                          FilesFind fileFind = filesFind.Find(ff =>
                          {
                              bool trovato = false;
                              switch (ff.FileCompare)
                              {
                                  case FileCompareEnum.Like:
                                      trovato = f.Name.ToUpper().Contains(ff.FileNameSearch);
                                      break;
                                  case FileCompareEnum.Equal:
                                  default:
                                      trovato = ff.FileNameSearch == f.Name.ToUpper();
                                      break;
                              }
                              return trovato;
                          });

                          if (fileFind != null)
                          {
                              fileFind.PathNameFound = f.DirectoryName;
                              fileFind.RealFileNameFound = Path.GetFileName(f.FullName);
                              fileStateFinder.FileCount++;
                          }
                      }
                      catch (Exception)
                      {

                      }
                  });

              return true;
          };

            BackgroundWorker bw = sender as BackgroundWorker;
            List<FilesFind> fileTofind = (List<FilesFind>)e.Argument;

            fileStateFinder.DirCount = 0;
            fileStateFinder.DirProgess = 0;
            fileStateFinder.FileCount = 0;

            var pathToSearch = fileTofind.GroupBy((f) => f.PathFileSearch);

            DirectoryInfo dirInfo;
            foreach (var item in pathToSearch)
            {
                dirInfo = new DirectoryInfo(item.First().PathFileSearch);

                fileStateFinder.DirProgess++;

                List<FilesFind> fileListToCheck = item.ToList();

                //Ricerca file
                _RicercaFile(dirInfo, fileListToCheck);


                dirInfo.EnumerateDirectories("*.*", SearchOption.TopDirectoryOnly).AsParallel().ForAll(
                     (dirCheck) =>
                     {
                         try
                         {
                             fileStateFinder.DirProgess++;
                             _RicercaFile(dirCheck, fileListToCheck);

                             dirCheck.EnumerateDirectories("*.*", SearchOption.AllDirectories).AsParallel().ForAll(d =>
                             {
                                 fileStateFinder.DirProgess++;
                                 _RicercaFile(d, fileListToCheck);
                             });
                         }
                         catch (Exception ex)
                         {
                         }
                     }
                    );
                //Fine ricerca file
            }
            e.Result = e.Argument;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelPercorsoButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog browser = new System.Windows.Forms.FolderBrowserDialog();
            browser.Description = "selezionare la directory dove cercare il file";
            browser.ShowNewFolderButton = false;
            browser.ShowDialog();
            txtPercorsoRicerca.Text = browser.SelectedPath;
            _fileFinds.Clear();

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportaButton_Click(object sender, RoutedEventArgs e)
        {
            string pathToSearch="";
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".txt";
            dlg.Filter = "Text documents (.txt)|*.txt";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                _fileFinds.Clear();
                // Open document 
                string filename = dlg.FileName;
                txtNomeFile.Text = System.IO.Path.GetFileName(dlg.FileName);
                txtPercorsoRicerca.Text = "";

                Func<string, object> SplitFileName = (filenameToSplit) =>
                {

                    //rimuovo la parte del file separata da |
                    string _filename = filenameToSplit.IndexOf("|") != -1 ?
                            filenameToSplit.Substring(0, filenameToSplit.IndexOf("|")) +
                            filenameToSplit.Substring(filenameToSplit.LastIndexOf("|") + 1)
                                : filenameToSplit;
                    //prendo il percorso se presente dove sono i file ricercati, spostarli nel
                    //percorso found se trovata corrispondenza con quelli presenti nel percorso di ricerca
                    string _pathname = Path.GetDirectoryName(_filename);
                    
                    _filename =filenameToSplit.IndexOf("#") != -1 ?
                                   Path.GetFileName(_filename.Substring(0, _filename.IndexOf("#")).ToUpper()) :
                                   Path.GetFileName(_filename.ToUpper());


                    return new
                    {
                        filename = _filename,
                        pathname = _pathname,
                        partfilename=filenameToSplit.IndexOf("|") != -1 ? 
                                     filenameToSplit.Substring( filenameToSplit.IndexOf("|")+1,
                                      filenameToSplit.LastIndexOf("|") - 1 - filenameToSplit.IndexOf("|")) : "",
                        compare = filenameToSplit.IndexOf("#") != -1 ?
                                    (FileCompareEnum)Enum.Parse(typeof(FileCompareEnum),
                                    filenameToSplit.Substring(filenameToSplit.IndexOf("#") + 1)) : FileCompareEnum.Equal
                    };
                };
                

                _fileNameImported = filename;
                foreach (var item in System.IO.File.ReadAllLines(filename))
                {
                    if (pathToSearch != item && item.IndexOf("@") != -1)
                        pathToSearch = item;
                    else
                    {
                        dynamic  objFileNameSplit =  SplitFileName(item);

                        _fileFinds.Add(new FilesFind()
                        {
                            FileNameSearch = objFileNameSplit.filename,
                            PathNameFound = "",
                            PathFileSource = objFileNameSplit.pathname,
                            PathFileSearch = pathToSearch.Remove(0, 1),
                            PartFileNameRemove = objFileNameSplit.partfilename,
                            FileCompare = objFileNameSplit.compare
                        });
                    }
                }

                fileStateFinder.FileSearchCount = _fileFinds.Count;
            }
        }


    }
}
