using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CrmRegardingAddin
{
    partial class CrmLinkPane
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnUnlink;
        
        private System.Windows.Forms.Button btnFinalize;private System.Windows.Forms.ListView lvLinks;
        private System.Windows.Forms.ColumnHeader colRole;
        private System.Windows.Forms.ColumnHeader colName;
        private System.Windows.Forms.ColumnHeader colEntity;
        private System.Windows.Forms.ColumnHeader colId;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnUnlink = new System.Windows.Forms.Button();
            this.btnFinalize = new System.Windows.Forms.Button();
            this.lvLinks = new System.Windows.Forms.ListView();
            this.colRole = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colEntity = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoEllipsis = true;
            this.lblTitle.Location = new System.Drawing.Point(10, 10);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(380, 20);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Liens CRM détectés pour ce mail :";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefresh.Location = new System.Drawing.Point(396, 6);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(85, 26);
            this.btnRefresh.TabIndex = 1;
            this.btnRefresh.Text = "Rafraîchir";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 

            // 
            // btnFinalize
            // 
            this.btnFinalize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFinalize.Location = new System.Drawing.Point(486, 6);
            this.btnFinalize.Name = "btnFinalize";
            this.btnFinalize.Size = new System.Drawing.Size(120, 26);
            this.btnFinalize.TabIndex = 4;
            this.btnFinalize.Text = "Finaliser le lien";
            this.btnFinalize.UseVisualStyleBackColor = true;
            this.btnFinalize.Click += new System.EventHandler(this.btnFinalize_Click);
                // btnUnlink
            // 
            this.btnUnlink.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUnlink.Location = new System.Drawing.Point(611, 6);
            this.btnUnlink.Name = "btnUnlink";
            this.btnUnlink.Size = new System.Drawing.Size(120, 26);
            this.btnUnlink.TabIndex = 2;
            this.btnUnlink.Text = "Annuler le lien…";
            this.btnUnlink.UseVisualStyleBackColor = true;
            this.btnUnlink.Click += new System.EventHandler(this.btnUnlink_Click);
            // 
            // lvLinks
            // 
            this.lvLinks.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvLinks.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colRole,
            this.colName,
            this.colEntity,
            this.colId});
            this.lvLinks.FullRowSelect = true;
            this.lvLinks.HideSelection = false;
            this.lvLinks.Location = new System.Drawing.Point(13, 36);
            this.lvLinks.MultiSelect = false;
            this.lvLinks.Name = "lvLinks";
            this.lvLinks.Size = new System.Drawing.Size(718, 128);
            this.lvLinks.TabIndex = 3;
            this.lvLinks.UseCompatibleStateImageBehavior = false;
            this.lvLinks.View = System.Windows.Forms.View.Details;
            this.lvLinks.SelectedIndexChanged += new System.EventHandler(this.lvLinks_SelectedIndexChanged);
            this.lvLinks.DoubleClick += new System.EventHandler(this.lvLinks_DoubleClick);
            // 
            // colRole
            // 
            this.colRole.Text = "Rôle";
            this.colRole.Width = 110;
            // 
            // colName
            // 
            this.colName.Text = "Nom";
            this.colName.Width = 360;
            // 
            // colEntity
            // 
            this.colEntity.Text = "Entité";
            this.colEntity.Width = 120;
            // 
            // colId
            // 
            this.colId.Text = "Id";
            this.colId.Width = 220;
            // 
            // CrmLinkPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lvLinks);
            this.Controls.Add(this.btnUnlink);
            this.Controls.Add(this.btnFinalize);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.lblTitle);
            this.Name = "CrmLinkPane";
            this.Size = new System.Drawing.Size(744, 177);
            this.ResumeLayout(false);

        }

        #endregion
    }
}