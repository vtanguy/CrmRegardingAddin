using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;

namespace CrmRegardingAddin
{
    public partial class SearchDialog : Form
    {
        private readonly IOrganizationService _org;
        private readonly IEnumerable<EntityReference> _initialSuggestions;
        private List<Entity> _currentResults = new List<Entity>();
        private List<ColumnSpec> _currentColumns = new List<ColumnSpec>();

        public EntityReference SelectedReference { get; private set; }

        public SearchDialog(IOrganizationService org, IEnumerable<EntityReference> initialSuggestions = null)
        {
            InitializeComponent();
            _org = org;
            _initialSuggestions = initialSuggestions ?? new List<EntityReference>();
        }

        private void SearchDialog_Load(object sender, EventArgs e)
        {
            try
            {
                var entities = CrmSearchService.GetEntityOptions();
                this.cboEntity.DisplayMember = "DisplayName";
                this.cboEntity.ValueMember = "LogicalName";
                this.cboEntity.DataSource = entities;
                int defaultIdx = entities.FindIndex(x => x.LogicalName == "contact");
                this.cboEntity.SelectedIndex = defaultIdx >= 0 ? defaultIdx : 0;

                LoadViewsForSelectedEntity();
                BuildColumnsForCurrentView();
                RunSearch();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur au chargement : " + ex.Message, "CRM",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cboEntity_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                LoadViewsForSelectedEntity();
                BuildColumnsForCurrentView();
                RunSearch();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur changement d’entité : " + ex.Message, "CRM",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cboView_SelectedIndexChanged(object sender, EventArgs e)
        {
            BuildColumnsForCurrentView();
            RunSearch();
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true; e.SuppressKeyPress = true; RunSearch();
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            RunSearch();
        }

        private void LoadViewsForSelectedEntity()
        {
            var entity = this.cboEntity.SelectedItem as SearchEntityOption;
            if (entity == null) return;

            var views = CrmSearchService.GetViews(_org, entity.LogicalName);
            this.cboView.DisplayMember = "Name";
            this.cboView.ValueMember = "Id";
            this.cboView.DataSource = views;

            int qfIndex = views.FindIndex(v => v.IsQuickFind);
            if (qfIndex >= 0) this.cboView.SelectedIndex = qfIndex;
            else if (views.Count > 0) this.cboView.SelectedIndex = 0;

            this.cboView.SelectedIndexChanged -= cboView_SelectedIndexChanged;
            this.cboView.SelectedIndexChanged += cboView_SelectedIndexChanged;
        }

        private void BuildColumnsForCurrentView()
        {
            var entity = this.cboEntity.SelectedItem as SearchEntityOption;
            var view = this.cboView.SelectedItem as ViewInfo;
            this.lvResults.BeginUpdate();
            try
            {
                this.lvResults.Columns.Clear();
                _currentColumns.Clear();

                if (view != null && !string.IsNullOrWhiteSpace(view.LayoutXml))
                {
                    var cols = CrmSearchService.ParseLayoutColumns(_org, entity.LogicalName, view.LayoutXml);
                    if (cols.Count == 0) cols.Add(new ColumnSpec { EntityLogicalName = entity.LogicalName, Attribute = "name", Width = 220, Header = "Nom" });
                    _currentColumns = cols;

                    foreach (var c in cols)
                    {
                        var ch = new ColumnHeader();
                        ch.Text = string.IsNullOrWhiteSpace(c.Header) ? c.Attribute : c.Header;
                        ch.Width = Math.Max(80, Math.Min(400, c.Width));
                        this.lvResults.Columns.Add(ch);
                    }
                }
                else
                {
                    _currentColumns.Add(new ColumnSpec { EntityLogicalName = entity.LogicalName, Attribute = "name", Width = 220, Header = "Nom" });
                    _currentColumns.Add(new ColumnSpec { EntityLogicalName = entity.LogicalName, Attribute = "emailaddress1", Width = 200, Header = "Email" });

                    var ch1 = new ColumnHeader { Text = "Nom", Width = 220 };
                    var ch2 = new ColumnHeader { Text = "Email", Width = 200 };
                    this.lvResults.Columns.AddRange(new[] { ch1, ch2 });
                }
            }
            finally
            {
                this.lvResults.EndUpdate();
            }
        }

        private void RunSearch()
        {
            var entity = this.cboEntity.SelectedItem as SearchEntityOption;
            var view = this.cboView.SelectedItem as ViewInfo;
            string term = this.txtSearch.Text ?? "";

            var results = CrmSearchService.Search(_org, entity.LogicalName, view, term, top: 200);
            _currentResults = results.Entities.ToList();

            var listView = this.lvResults;
            if (listView != null)
            {
                listView.BeginUpdate();
                try
                {
                    listView.Items.Clear();
                    foreach (var r in _currentResults)
                    {
                        var values = new List<string>();
                        foreach (var col in _currentColumns)
                        {
                            values.Add(FormatValue(r, col.EntityLogicalName, col.Attribute));
                        }
                        var item = new ListViewItem(values.ToArray());
                        item.Tag = r;
                        listView.Items.Add(item);
                    }
                }
                finally
                {
                    listView.EndUpdate();
                }
            }
        }

        private string FormatValue(Entity e, string entityLogicalName, string attributeLogicalName)
        {
            if (e == null) return "";

            // attribut direct
            if (e.Attributes.ContainsKey(attributeLogicalName))
                return FormatRawValue(entityLogicalName, attributeLogicalName, e[attributeLogicalName]);

            // alias "alias.attr" ou "entity.attr" → stocké sous forme d'AliasedValue
            var aliasKey = attributeLogicalName;
            if (!aliasKey.Contains("."))
            {
                // parfois layout sans alias → rien
                return "";
            }

            if (e.Attributes.ContainsKey(aliasKey))
            {
                var av = e[aliasKey] as Microsoft.Xrm.Sdk.AliasedValue;
                var val = av != null ? av.Value : e[aliasKey];
                // essayer de déduire l'entité depuis "entity.attr"
                var parts = aliasKey.Split('.');
                var ent = parts.Length == 2 ? parts[0] : entityLogicalName;
                return FormatRawValue(ent, parts.Length == 2 ? parts[1] : attributeLogicalName, val);
            }
            return "";
        }

        private string FormatRawValue(string entityLogicalName, string attributeLogicalName, object v)
        {
            if (v == null) return "";
            var er = v as EntityReference;
            if (er != null) return !string.IsNullOrEmpty(er.Name) ? er.Name : er.Id.ToString();
            var os = v as OptionSetValue;
            if (os != null) return MetadataCache.GetOptionLabel(_org, entityLogicalName, attributeLogicalName, os.Value);
            var money = v as Microsoft.Xrm.Sdk.Money;
            if (money != null) return money.Value.ToString("0.00");
            if (v is bool)
            {
                return MetadataCache.GetBooleanLabel(_org, entityLogicalName, attributeLogicalName, (bool)v);
            }
            var dt = v as DateTime?;
            if (dt != null) return dt.Value.ToString("g");
            return Convert.ToString(v);
        }

        private void SelectCurrent(ListViewItem selectedItem)
        {
            var ent = selectedItem != null ? selectedItem.Tag as Entity : null;
            if (ent == null) return;

            var er = new EntityReference(ent.LogicalName, ent.Id);
            if (ent.Contains("name")) er.Name = Convert.ToString(ent["name"]);
            else if (ent.Contains("fullname")) er.Name = Convert.ToString(ent["fullname"]);
            else if (ent.Contains("title")) er.Name = Convert.ToString(ent["title"]);

            this.SelectedReference = er;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void lvResults_DoubleClick(object sender, EventArgs e)
        {
            var lv = sender as ListView;
            if (lv != null && lv.SelectedItems.Count == 1)
                SelectCurrent(lv.SelectedItems[0]);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var lv = this.lvResults;
            if (lv != null && lv.SelectedItems.Count == 1)
                SelectCurrent(lv.SelectedItems[0]);
        }
    }

    internal static class ListExtensions
    {
        public static int FindIndex<T>(this IList<T> list, Func<T, bool> predicate)
        {
            for (int i = 0; i < list.Count; i++)
                if (predicate(list[i])) return i;
            return -1;
        }
    }
}
