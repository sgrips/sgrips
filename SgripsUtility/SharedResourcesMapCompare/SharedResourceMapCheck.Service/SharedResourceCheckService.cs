using Sgrips.Utility.IO.SharedResourceMapCheck.Engine;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Service
{
    public partial class SharedResourceCheckService : ServiceBase
    {

        private EventLog eventLog = new EventLog();
        private Task taskCheck = null;
        private CancellationTokenSource tokenSource = new CancellationTokenSource();



        public SharedResourceCheckService()
        {
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists("Application"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                     "SharedResourceCheckService", string.Empty);
            }
            eventLog.Source = "SharedResourceCheckService";
            WriteLog("Inizializzazione SharedResourceCheckService");
        }

        public void WriteLog(string msg)
        {
            WriteLog(msg, EventLogEntryType.Information);
        }

        public void WriteLog(string msg, EventLogEntryType logType, bool sendEmail = false, string messageErr = "")
        {
            string stackTrace = "";
            System.Net.Mail.MailAddress addressFrom = null;
            System.Net.Mail.MailAddress addressTo = null;
            System.Net.Mail.MailMessage message = null;
            System.Net.Mail.SmtpClient smtp = null;

            try
            {
                eventLog.WriteEntry(msg, logType);
            }
            catch (Exception ex)
            {

                sendEmail = true;
                messageErr = ex.Message;
                stackTrace = ex.StackTrace.ToString();
            }

            if (sendEmail)
            {

                try
                {
                    addressFrom = new System.Net.Mail.MailAddress("Check.Share@brt.it", String.Format("{0}", "CheckShareBRTService"));
                    addressTo = new System.Net.Mail.MailAddress("cedalert@brt.it");
                    message = new System.Net.Mail.MailMessage(addressFrom, addressTo);
                    smtp = new System.Net.Mail.SmtpClient();
                    message.Subject = String.Format("Modulo:{0} PC Name:{1}", "WriteLog", System.Environment.MachineName);
                    message.Body = string.Format("Error: {0} StackTrace: {1}", messageErr, stackTrace);
                    smtp.Send(message);
                }
                catch (Exception)
                {
                }

            }
        }


        protected override void OnStart(string[] args)
        {

            try
            {
                taskCheck = Task.Run(() =>
                {

                    WriteLog("Started SharedResourceCheckService Service");

                    EngineSharedResourceCheck engine = null;
                    try
                    {
                        engine = new EngineSharedResourceCheck();

                        WriteLog(string.Format( "EngineSharedResourceCheck Database Connected to {0}",
                            engine.DataBasePath));

                        while (true)
                        {

                            System.Threading.Thread.Sleep(200);
                            tokenSource.Token.ThrowIfCancellationRequested();

                            foreach (var item in engine.SharedResourceCheck)
                            {
                                if (item.DateTimeStart <= DateTime.Now)
                                {
                                    engine.RunTask(item);
                                    engine.UpdateDateTimeStart(item);
                                }

                            }
                        }

                    }


                    catch (OperationCanceledException ex) { }
                    catch (Exception ex)
                    {

                        WriteLog(string.Format("SharedResourceCheckService {0}", ex.StackTrace.ToString()), EventLogEntryType.Error, true, "Service Stopped for Critical Error Occurred");
                        System.ServiceProcess.ServiceController svc = new System.ServiceProcess.ServiceController(base.ServiceName);
                        svc.Stop();
                    }
                }, tokenSource.Token);
            }
            catch (Exception ex)
            {
                WriteLog(string.Format("SharedResourceCheckService {0}", ex.StackTrace.ToString()), EventLogEntryType.Error, true, "Service Stopped for Critical Error Occurred");
                System.ServiceProcess.ServiceController svc = new System.ServiceProcess.ServiceController(base.ServiceName);
                svc.Stop();

            }



        }


        private void CheckResourcesMap()
        {

        }

        protected override void OnStop()
        {
            try
            {
                tokenSource.Cancel();
                taskCheck.Wait();
            }
            catch (Exception)
            {
            }

            WriteLog("Stop SharedResourceCheckService Service");
        }
    }
}
