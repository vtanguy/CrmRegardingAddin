using System;
using System.Drawing;
using System.Security.Principal;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    public class LoginPromptForm : Form
    {
        private TextBox _user;
        private TextBox _pass;
        private Button _ok;
        private Button _cancel;
        private CheckBox _remember;
        private CheckBox _showPwd;

        public string EnteredUserName { get { return _user.Text; } }
        public string EnteredPassword { get { return _pass.Text; } }
        public bool RememberPassword { get { return _remember.Checked; } set { _remember.Checked = value; } }

        public LoginPromptForm()
        {
            this.Text = "Connexion CRM (WS-Trust)";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ClientSize = new Size(460, 206);
            this.Font = SystemFonts.MessageBoxFont;
            this.KeyPreview = true;

            var lblInfo = new Label { Left = 12, Top = 12, Width = 430, Height = 28,
                Text = "Entrez vos identifiants CRM." };

            var lblUser = new Label { Left = 12, Top = 52, Width = 110, Text = "Identifiant :" };
            _user = new TextBox { Left = 130, Top = 48, Width = 300 };

            var lblPass = new Label { Left = 12, Top = 82, Width = 110, Text = "Mot de passe :" };
            _pass = new TextBox { Left = 130, Top = 78, Width = 300, UseSystemPasswordChar = true };

            _remember = new CheckBox { Left = 130, Top = 112, Width = 300, Text = "Se souvenir du mot de passe" };
            _showPwd  = new CheckBox { Left = 130, Top = 134, Width = 300, Text = "Afficher le mot de passe" };

            _ok = new Button { Text = "OK", Left = 268, Top = 166, Width = 80, DialogResult = DialogResult.OK };
            _cancel = new Button { Text = "Annuler", Left = 354, Top = 166, Width = 80, DialogResult = DialogResult.Cancel };

            this.Controls.Add(lblInfo);
            this.Controls.Add(lblUser);
            this.Controls.Add(_user);
            this.Controls.Add(lblPass);
            this.Controls.Add(_pass);
            this.Controls.Add(_remember);
            this.Controls.Add(_showPwd);
            this.Controls.Add(_ok);
            this.Controls.Add(_cancel);

            this.AcceptButton = _ok;
            this.CancelButton = _cancel;

            _showPwd.CheckedChanged += (s, e) => { _pass.UseSystemPasswordChar = !_showPwd.Checked; };

            // Préremplissage UPN automatiquement si possible
            this.Shown += (s, e) =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(_user.Text))
                    {
                        // 1) Via Outlook/Exchange
                        try
                        {
                            var app = Globals.ThisAddIn != null ? Globals.ThisAddIn.Application : null;
                            var ae = app?.Session?.CurrentUser?.AddressEntry;
                            var exu = ae?.GetExchangeUser();
                            var upn = exu?.PrimarySmtpAddress;
                            if (!string.IsNullOrEmpty(upn)) _user.Text = upn;
                        }
                        catch { }

                        // 2) Fallback WindowsIdentity (DOMAIN\\user -> user@domaine si suffixe connu)
                        if (string.IsNullOrWhiteSpace(_user.Text))
                        {
                            try
                            {
                                var win = WindowsIdentity.GetCurrent().Name; // DOMAIN\\user
                                var slash = win.IndexOf('\\');
                                if (slash > 0 && slash < win.Length - 1)
                                {
                                    var sam = win.Substring(slash + 1);
                                    // si vous avez un suffixe connu, remplacez-le ici
                                    // ex: _user.Text = sam + "@votre-domaine.tld";
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
                _user.Focus();
            };
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // LoginPromptForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "LoginPromptForm";
            this.Load += new System.EventHandler(this.LoginPromptForm_Load);
            this.ResumeLayout(false);

        }

        private void LoginPromptForm_Load(object sender, EventArgs e)
        {

        }
    }
}
