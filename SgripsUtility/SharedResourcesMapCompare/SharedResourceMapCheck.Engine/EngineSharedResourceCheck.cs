using Sgrips.Framework.FileSystemDatabase;
using Sgrips.Utility.IO.SharedResourceMapCheck.Engine.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Engine
{
    public class EngineSharedResourceCheck
    {
        const string HTML_START = @"<!DOCTYPE html>
                                    <html>
                                    <head>
                                    <style >
                                    th,td { border: 1px solid #dadada;
                                        border-left: none;
                                        border-top: none;
                                        padding: 3px 3px 3px 3px;}
                                    th{ background-color: #d7172a;color:#ffffff}
                                    table {     border-collapse: collapse;
                                        border-spacing: 0px;  } 
                                    td { height:20px;  
                                        border-right: 1px solid #dadada;
                                        border-bottom: 1px solid #dadada;
                                        border-left: none;
                                        border-top: 1px solid #dadada;
                                        padding: 2px; } 
                                    tr:nth-child(odd) { background-color: #eeeeee;  }
                                    </style>
                                    </head>
                                    <body>";
        private const string HTML_TABLE_EXISTS = @"<table border=0>
                                    <thead>
                                    <th>Server Name</th>
                                    <th>Share Name</th>
                                    <th>Share State</th>
                                    <th>Error Description</th>
                                    </thead>
                                    <tbody>";
        private const string HTML_TABLE_COMPARE = @"<table border=0>
                                    <thead>
                                    <th>Source</th>
                                    <th>Destination</th >
                                    <th>Share State</th>
                                    <th>Error Description</th>
                                    </thead>
                                    <tbody>";
        private const string HTML_END = @"</tbody>
                            </table>
                            </body>
                            </html>";


        #region properties
        public IList<SharedResourcesModel> SharedResourceCheck { get; private set; }
        #endregion

        private StringBuilderExtension _sbOut = null;
        private bool _addTo;

        private JSONRepository _database;
        public string DataBasePath { get; private set; }


        /// <summary>
        /// 
        /// </summary>
        public EngineSharedResourceCheck()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            string company = FileVersionInfo.GetVersionInfo(assembly.Location).CompanyName;
            DataBasePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

            LoadConfiguration(System.IO.Path.Combine(DataBasePath, company, "SharedResourceCheck"));

        }

        /// <summary>
        /// 
        /// </summary>
        private void LoadConfiguration(string pathDataBase = null)
        {
            _database = new JSONRepository( pathDataBase, "SharedResourcesCheck","Id");

            SharedResourceCheck = _database.GetAll<SharedResourcesModel>() as IList<SharedResourcesModel>;

            if (SharedResourceCheck == null || SharedResourceCheck.All(f => f.ShareResources.Count() == 0))
                _database.Save(new[] { new SharedResourcesModel()
                {
                    ShareResources = new List<SharedResourceModel>() { new SharedResourceModel() { ServerSource = "Source", ServerTarget = "Target" } },
                    Id = 0,
                    EmailAlert = "cedalert@brt.it",
                    EmailSubject =  "Stato Share Compare: {0}",
                    ShellCommand = "NET.EXE",
                    ShellCommandParameters = "VIEW",
                    DateTimeStart = DateTime.Now,
                    DayOfWeek = Scheduler.DayOfWeekOperation.Monday,
                    RecursiveTime = "D",
                    RecursivePeriod = null
                }});


        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void RunTask(SharedResourcesModel sharedResourceModel)
        {
            StringBuilder body = null;
            string subject = "";


            if (sharedResourceModel == null || SharedResourceCheck == null)
            {
                subject = string.Format(sharedResourceModel.EmailSubject , sharedResourceModel.EmailSubject, "ERROR");
                body = new StringBuilder("ATTENZIONE CONFIGURARE SERVIZIO </br>Utilizzando il file di configurazione presente in </br>").
                    AppendLine(_database.FullDataBaseName);
                SendEmail(body.ToString(), subject, "cedalert@brt.it",true);
                return;
            }


            switch (sharedResourceModel.CheckType)
            {
                case CheckTypeEnum.Exists:
                    CompareShareExists(sharedResourceModel, sharedResourceModel.EmailAlert);
                    break;
                case CheckTypeEnum.Compare:
                    CompareShareDir(sharedResourceModel, sharedResourceModel.EmailAlert);
                    break;
                default:
                    break;
            }

        }

        private void CompareShareExists(SharedResourcesModel sharedResourceModel, string addressEmail)
        {
            string sError = "";
            bool bError = false;
            StringBuilder body = null;
            string subject = "";

            body = new StringBuilder(HTML_START).AppendLine(HTML_TABLE_EXISTS);

            foreach (var item in sharedResourceModel.ShareResources)
            {
                subject = null;
                sError = null;


                List<NetViewModel> list1 = null;

                try
                {
                    list1 = ExecNetView(item.ServerSource);
                }
                catch (EngineSharedResourceCheckException ex)
                {
                    sError = string.Format("{0}:{1}", item.ServerSource, ex.Message);
                }


                foreach (var shareCheck in item.ServerSourceShareList.Split (',') )
                {
                    body.Append("<tr>");
                    body.Append(string.Format("<td>{0}</td>", item.ServerSource));
                    body.Append(string.Format("<td>{0}</td>", shareCheck ));

                    if (sError != null)
                    {
                        bError = true;
                        body.Append("<td>ERROR</td>").
                             Append(string.Format("<td>{0}</td>", sError));

                        break;
                    }

                    if (list1 != null )
                    {
                        if (!list1.Exists(f => f.SharedName.ToUpper() == shareCheck.ToUpper ()))
                        {
                            bError = true;
                            body.Append("<td>Error</td>").
                            Append(string.Format("<td>Not Found</td>"));
                        }
                        else
                            body.Append("<td>OK</td>").Append("<td></td>");
                    }
                }

                body.AppendLine("</tr>");
            }

            subject = string.Format(sharedResourceModel.EmailSubject , bError ? "ERROR" : "OK");
            body.AppendLine(HTML_END);
            SendEmail(body.ToString(), subject, addressEmail, bError);

        }


        private void CompareShareDir(SharedResourcesModel sharedResourceModel, string addressEmail)
        {
            string sError = "";
            bool bError = false;
            StringBuilder body = null;
            string subject = "";

            body = new StringBuilder(HTML_START).AppendLine(HTML_TABLE_COMPARE );

            foreach (var item in sharedResourceModel.ShareResources)
            {
                subject = null;
                sError = null;


                List<NetViewModel> list1 = null;
                List<NetViewModel> list2 = null;
                body.Append("<tr>");
                body.Append(string.Format("<td>{0}</td>", item.ServerSource));
                body.Append(string.Format("<td>{0}</td>", item.ServerTarget));
                try
                {
                    list1 = ExecNetView(item.ServerSource);
                }
                catch (EngineSharedResourceCheckException ex)
                {
                    sError = string.Format("{0}:{1}", item.ServerSource, ex.Message);
                }

                try
                {
                    list2 = ExecNetView(item.ServerTarget);
                }
                catch (EngineSharedResourceCheckException ex)
                {
                    sError += string.Format(" {0}:{1}", item.ServerTarget, ex.Message);
                }

                if (sError != null)
                {
                    bError = true;
                    body.Append("<td>ERROR</td>").
                         Append(string.Format("<td>{0}</td>", sError));
                }

                if (list1 != null && list2 != null)
                {

                    bool same = list1.SequenceEqual(list2, new NetViewModelComparer());
                    if (!same)
                    {
                        bError = true;
                        body.Append("<td>Error</td>").
                             Append(string.Format("<td>{0}, Share Count={1} - {2}, Share Count={3}</td>",
                                       item.ServerSource, list1.Count(),
                                       item.ServerTarget, list2.Count()));
                    }
                    else
                        body.Append("<td>OK</td>").Append("<td></td>");
                }

                body.AppendLine("</tr>");
            }

            subject = string.Format(sharedResourceModel.EmailSubject , bError ? "ERROR" : "OK");
            body.AppendLine(HTML_END);
            SendEmail(body.ToString(), subject, addressEmail, bError);

        }


        /// <summary>
        /// 
        /// </summary>
        public void UpdateDateTimeStart(SharedResourcesModel sharedResourceModel)
        {

            DateTime DateTimeStart = Scheduler.CalculateNextTimeTask(DateTime.Now, sharedResourceModel.DateTimeStart,
                sharedResourceModel.RecursivePeriod, sharedResourceModel.RecursiveTime, sharedResourceModel.DayOfWeek);

            sharedResourceModel.DateTimeStart = DateTimeStart;




            _database.Modify(sharedResourceModel);
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="remoteAddress"></param>
        /// <returns></returns>
        private List<NetViewModel> ExecNetView(string remoteAddress)
        {
            Process p = new Process();
            ProcessStartInfo pInfo = new ProcessStartInfo();
            pInfo.Arguments = string.Format(@"VIEW {0}", remoteAddress);
            pInfo.FileName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "Net.exe");
            pInfo.RedirectStandardOutput = true;
            pInfo.RedirectStandardError = true;
            pInfo.UseShellExecute = false;

            _addTo = false;
            _sbOut = new StringBuilderExtension();
            p.StartInfo = pInfo;
            p.StartInfo.CreateNoWindow = false;
            p.OutputDataReceived += new DataReceivedEventHandler(ConsoleOutputHandler);
            p.ErrorDataReceived += new DataReceivedEventHandler(ConsoleOutputHandler);
            p.Start();

            p.BeginOutputReadLine();
            p.BeginErrorReadLine();


            p.WaitForExit();

            if (_sbOut.Error != null) throw _sbOut.Error;

            var qry = from item in _sbOut.ToString().Split('\n')
                      let obj = SplitRegx(item.Trim())
                      where obj.Length == 4 || obj.Length == 3 || obj.Length == 2
                      select new NetViewModel() { SharedName = obj[0], ShareType = obj[1], ShareComment = obj.Length == 3 ? obj[2] : "" };

            return qry.OrderBy(f => f.SharedName).ToList();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private string[] SplitRegx(string value)
        {
            string pattern = @"\s+";
            Regex rex = new Regex(pattern);
            return rex.Split(value);
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConsoleOutputHandler(object sender, DataReceivedEventArgs e)
        {
            if (e.Data != null)
            {
                if (e.Data.IndexOf("-") != -1) _addTo = true;
                else if (e.Data.ToLower().IndexOf("esecuzione comando") != -1 ||
                        e.Data.ToLower().IndexOf("the command completed successfully") != -1) _addTo = false;
                else if (e.Data.ToLower().IndexOf("errore di sistema") != -1 ||
                        e.Data.ToLower().IndexOf("system error") != -1)
                {
                    _addTo = true;
                    _sbOut.Error = new EngineSharedResourceCheckException(e.Data);
                };

                if (_addTo) _sbOut.AppendLine(e.Data);
            }
        }




        /// <summary>
        /// 
        /// </summary>
        /// <param name="body"></param>
        /// <param name="subject"></param>
        private void SendEmail(string body, string subject, string mailTo,bool important)
        {

            System.Net.Mail.MailAddress addressFrom = null;
            System.Net.Mail.MailAddress addressTo = null;
            System.Net.Mail.MailMessage message = null;
            System.Net.Mail.SmtpClient smtp = null;

            try
            {
                addressFrom = new System.Net.Mail.MailAddress("Check.Share@brt.it", String.Format("{0}", "CheckShareBRTService"));
                addressTo = new System.Net.Mail.MailAddress(mailTo);
                message = new System.Net.Mail.MailMessage(addressFrom, addressTo);
                message.IsBodyHtml = true;
                smtp = new System.Net.Mail.SmtpClient();
                message.Priority = important ?  System.Net.Mail.MailPriority.High : System.Net.Mail.MailPriority.Normal ;
                message.Subject = subject;
                message.Body = body;
                smtp.Send(message);
            }
            catch (Exception ex)
            {
            }
        }

    }
}
