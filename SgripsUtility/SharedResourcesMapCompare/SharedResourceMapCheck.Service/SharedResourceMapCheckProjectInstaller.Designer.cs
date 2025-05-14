namespace Sgrips.Utility.IO.SharedResourceMapCheck.Service
{
    partial class SharedResourceMapCheckProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SharedResourceMapCheckServiceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.SharedResourceMapCheckServiceInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // SharedResourceMapCheckServiceProcessInstaller
            // 
            this.SharedResourceMapCheckServiceProcessInstaller.Password = null;
            this.SharedResourceMapCheckServiceProcessInstaller.Username = null;
            // 
            // SharedResourceMapCheckServiceInstaller
            // 
            this.SharedResourceMapCheckServiceInstaller.Description = "Verifica che le risorse mappate di 2 server siano identiche utilizzando il comand" +
    "o NET VIEW";
            this.SharedResourceMapCheckServiceInstaller.DisplayName = "SharedResourceMapCheckService";
            this.SharedResourceMapCheckServiceInstaller.ServiceName = "SharedResourceMapCheckService";
            // 
            // SharedResourceMapCheckProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.SharedResourceMapCheckServiceProcessInstaller,
            this.SharedResourceMapCheckServiceInstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller SharedResourceMapCheckServiceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller SharedResourceMapCheckServiceInstaller;
    }
}