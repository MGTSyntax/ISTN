namespace ISTNotification
{
    partial class ProjectInstaller
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
            this.istnProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.istninstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // istnProcessInstaller
            // 
            this.istnProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.istnProcessInstaller.Password = null;
            this.istnProcessInstaller.Username = null;
            // 
            // istninstaller
            // 
            this.istninstaller.Description = "Notification for taking In-Service Training";
            this.istninstaller.DisplayName = "ISTN";
            this.istninstaller.ServiceName = "ISTN";
            this.istninstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.istnProcessInstaller,
            this.istninstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller istnProcessInstaller;
        private System.ServiceProcess.ServiceInstaller istninstaller;
    }
}