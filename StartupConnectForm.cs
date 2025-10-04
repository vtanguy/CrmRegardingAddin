using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;

namespace CrmRegardingAddin
{
    /// <summary>
    /// Résultat de connexion (compatible C#6, pas de tuples)
    /// </summary>
    public sealed class ConnectResult
    {
        public IOrganizationService Service;
        public string Diag;
    }

    /// <summary>
    /// Délégué exécuté par la fenêtre pour tenter la connexion.
    /// </summary>
    /// <param name="ct">CancellationToken</param>
    public delegate Task<ConnectResult> ConnectAsyncDelegate(CancellationToken ct);

    public class StartupConnectForm : Form
    {
        // UI (nom complet System.Windows.Forms.* pour éviter le conflit avec Microsoft.Xrm.Sdk.Label)
        private readonly System.Windows.Forms.Label _title = new System.Windows.Forms.Label();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly TextBox _diag = new TextBox();
        private readonly Button _btnRetry = new Button();
        private readonly Button _btnCancel = new Button();

        private CancellationTokenSource _cts;

        // Delegate à fournir depuis ThisAddIn
        public ConnectAsyncDelegate ConnectAsync;

        public IOrganizationService Service { get; private set; }

        public StartupConnectForm()
        {
            // Fenêtre
            Text = "Connexion à Dynamics 365";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 220);

            // Titre
            _title.Text = "Connexion au CRM…";
            _title.AutoSize = true;
            _title.Location = new Point(12, 12);

            // Barre de progression
            _progress.Style = ProgressBarStyle.Marquee;
            _progress.MarqueeAnimationSpeed = 30;
            _progress.Size = new Size(496, 18);
            _progress.Location = new Point(12, 40);

            // Zone diagnostic
            _diag.Multiline = true;
            _diag.ReadOnly = true;
            _diag.ScrollBars = ScrollBars.Vertical;
            _diag.Size = new Size(496, 100);
            _diag.Location = new Point(12, 68);
            _diag.Visible = false;

            // Boutons
            _btnRetry.Text = "Réessayer";
            _btnRetry.Enabled = false;
            _btnRetry.Size = new Size(90, 26);
            _btnRetry.Location = new Point(322, 180);
            _btnRetry.Click += async (s, e) => { await StartConnect(); };

            _btnCancel.Text = "Annuler";
            _btnCancel.Size = new Size(90, 26);
            _btnCancel.Location = new Point(418, 180);
            _btnCancel.Click += (s, e) =>
            {
                if (_cts != null) _cts.Cancel();
                DialogResult = DialogResult.Cancel;
                Close();
            };

            Controls.AddRange(new Control[] { _title, _progress, _diag, _btnRetry, _btnCancel });

            Shown += async (s, e) => { await StartConnect(); };
            FormClosed += (s, e) => { if (_cts != null) _cts.Cancel(); };
        }

        private async Task StartConnect()
        {
            _btnRetry.Enabled = false;
            _diag.Visible = false;
            _diag.Text = string.Empty;
            _progress.Style = ProgressBarStyle.Marquee;

            _cts = new CancellationTokenSource();

            try
            {
                ConnectResult result = null;

                if (ConnectAsync != null)
                    result = await ConnectAsync(_cts.Token);

                if (result != null && result.Service != null)
                {
                    Service = result.Service;
                    DialogResult = DialogResult.OK;
                    Close();
                    return;
                }

                // Échec : afficher le diagnostic
                _diag.Visible = true;
                _diag.Text = (result != null ? result.Diag : "Échec de la connexion (aucun résultat).");
                _btnRetry.Enabled = true;
                _progress.Style = ProgressBarStyle.Blocks;
            }
            catch (OperationCanceledException)
            {
                // Annulé par l’utilisateur
            }
            catch (System.Exception ex) // System.Exception pour éviter le conflit avec Outlook.Exception
            {
                _diag.Visible = true;
                _diag.Text = "Exception: " + ex.Message;
                _btnRetry.Enabled = true;
                _progress.Style = ProgressBarStyle.Blocks;
            }
        }
    }
}
