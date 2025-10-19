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

        // Liens récents
        private List<RecentLink> _recentLinks = new List<RecentLink>();

        // Tri
        private int _sortCol = -1;
        private SortOrder _sortOrder = SortOrder.None;
        private List<string> _baseHeaders = new List<string>();

        // Empêche les recherches involontaires pendant changements d'UI
        private bool _suspendSearch = false;

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
                // ---- Peupler les liens récents ----
                try
                {
                    _recentLinks = RecentLinks.GetAll();
                    this.cboRecent.Items.Clear();
                    this.cboRecent.Items.Add("(sélectionner un lien récent)");
                    foreach (var r in _recentLinks) this.cboRecent.Items.Add(r);
                    this.cboRecent.SelectedIndex = 0;
                }
                catch { }


                // Tri au clic
                this.lvResults.ColumnClick += lvResults_ColumnClick;

                // Pas de recherche automatique ici (l'utilisateur lance la recherche)
                //RunSearch();
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
                _suspendSearch = true;

                // Reset champ et état de tri
                this.txtSearch.Clear();
                _sortCol = -1;
                _sortOrder = SortOrder.None;

                // Vider résultats de l'ancienne entité
                _currentResults.Clear();
                this.lvResults.Items.Clear();

                LoadViewsForSelectedEntity();
                BuildColumnsForCurrentView();
                // ---- Peupler les liens récents ----
                try
                {
                    _recentLinks = RecentLinks.GetAll();
                    this.cboRecent.Items.Clear();
                    this.cboRecent.Items.Add("(sélectionner un lien récent)");
                    foreach (var r in _recentLinks) this.cboRecent.Items.Add(r);
                    this.cboRecent.SelectedIndex = 0;
                }
                catch { }

                UpdateSortHeaderGlyphs();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur changement d’entité : " + ex.Message, "CRM",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _suspendSearch = false;
            }
        }

        private void cboView_SelectedIndexChanged(object sender, EventArgs e)
        {
            _suspendSearch = true;
            try
            {
                BuildColumnsForCurrentView();
                // ---- Peupler les liens récents ----
                try
                {
                    _recentLinks = RecentLinks.GetAll();
                    this.cboRecent.Items.Clear();
                    this.cboRecent.Items.Add("(sélectionner un lien récent)");
                    foreach (var r in _recentLinks) this.cboRecent.Items.Add(r);
                    this.cboRecent.SelectedIndex = 0;
                }
                catch { }

                UpdateSortHeaderGlyphs();
            }
            finally
            {
                _suspendSearch = false;
            }
            // Pas de RunSearch automatique
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
                // Mémorise les en-têtes (sans flèche) et met à jour le glyph
                _baseHeaders = this.lvResults.Columns.Cast<ColumnHeader>().Select(ch => ch.Text).ToList();
                UpdateSortHeaderGlyphs();
                this.lvResults.EndUpdate();
            }
        }


        private static bool IsInstanceEntity(string logicalName)
        {
            if (string.IsNullOrWhiteSpace(logicalName)) return false;
            logicalName = logicalName.Trim().ToLowerInvariant();
            return logicalName == "new_instance" || logicalName == "instance";
        }

        private int FindBestDateSortColumn(List<Entity> sampleEntities)
        {
            if (_currentColumns == null || _currentColumns.Count == 0) return -1;

            // 1) Try well-known date attributes by name
            string[] preferred = new[] {
                "createdon","modifiedon","overriddencreatedon",
                "scheduledstart","scheduledend","actualstart","actualend",
                "new_date","new_startdate","new_enddate","new_duedate"
            };
            for (int i = 0; i < _currentColumns.Count; i++)
            {
                var attr = _currentColumns[i].Attribute ?? "";
                var last = attr.Contains(".") ? attr.Substring(attr.LastIndexOf('.') + 1) : attr;
                foreach (var p in preferred)
                {
                    if (string.Equals(last, p, StringComparison.OrdinalIgnoreCase))
                        return i;
                }
            }

            // 2) Inspect runtime types from first non-null entity value
            if (sampleEntities != null && sampleEntities.Count > 0)
            {
                var first = sampleEntities[0];
                for (int i = 0; i < _currentColumns.Count; i++)
                {
                    var raw = GetRawValue(first, _currentColumns[i]);
                    if (GetColKind(raw) == ColKind.Date) return i;
                }
            }

            return -1;
        }

        private void RunSearch()
        {
            if (_suspendSearch) return;

            var entity = this.cboEntity.SelectedItem as SearchEntityOption;
            var view = this.cboView.SelectedItem as ViewInfo;
            string term = this.txtSearch.Text ?? "";

            int topN = IsInstanceEntity(entity.LogicalName) ? 200 : 1000;
            var results = CrmSearchService.Search(_org, entity.LogicalName, view, term, top: topN);
            _currentResults = results.Entities.ToList();

            // Tri par défaut pour l'entité 'instance' / 'new_instance' : date la plus récente en premier
            try
            {
                var entOpt = this.cboEntity.SelectedItem as SearchEntityOption;
                if (entOpt != null && IsInstanceEntity(entOpt.LogicalName))
                {
                    int dateCol = FindBestDateSortColumn(_currentResults);
                    if (dateCol >= 0)
                    {
                        _sortCol = dateCol;
                        _sortOrder = SortOrder.Descending;
                    }
                    else
                    {
                        if (_sortCol < 0) { _sortCol = 0; _sortOrder = SortOrder.Descending; }
                    }
                }
            }
            catch { }

            // Applique le tri courant, puis affiche
            ApplySort();
            RenderCurrentResults();
            UpdateSortHeaderGlyphs();
        }

        // ------------------------- Tri & rendu -------------------------

        private object GetRawValue(Entity e, ColumnSpec col)
        {
            if (e == null || col == null) return null;
            var attr = col.Attribute;
            if (string.IsNullOrWhiteSpace(attr)) return null;

            if (e.Attributes.ContainsKey(attr))
            {
                var v = e[attr];
                var av = v as Microsoft.Xrm.Sdk.AliasedValue;
                return av != null ? av.Value : v;
            }

            if (attr.Contains(".") && e.Attributes.ContainsKey(attr))
            {
                var av = e[attr] as Microsoft.Xrm.Sdk.AliasedValue;
                return av != null ? av.Value : e[attr];
            }

            if (attr.EndsWith("name", StringComparison.OrdinalIgnoreCase) && e.Attributes.ContainsKey(attr))
            {
                return e[attr];
            }

            return null;
        }

        private enum ColKind { Text, Numeric, Date }

        private ColKind GetColKind(object raw)
        {
            if (raw == null) return ColKind.Text;
            if (raw is DateTime || raw is DateTime?) return ColKind.Date;
            if (raw is Microsoft.Xrm.Sdk.Money) return ColKind.Numeric;
            if (raw is sbyte || raw is byte || raw is short || raw is ushort || raw is int || raw is uint || raw is long || raw is ulong) return ColKind.Numeric;
            if (raw is float || raw is double || raw is decimal) return ColKind.Numeric;
            return ColKind.Text;
        }

        private int CompareEntitiesForColumn(Entity a, Entity b, int colIndex)
        {
            if (colIndex < 0 || colIndex >= _currentColumns.Count) return 0;
            var col = _currentColumns[colIndex];
            var va = GetRawValue(a, col);
            var vb = GetRawValue(b, col);

            var kind = GetColKind(va ?? vb);
            int sign = (_sortOrder == SortOrder.Descending) ? -1 : 1;

            switch (kind)
            {
                case ColKind.Numeric:
                    decimal da = 0m, db = 0m;
                    try { if (va is Microsoft.Xrm.Sdk.Money) da = ((Microsoft.Xrm.Sdk.Money)va).Value; else if (va is IConvertible) da = Convert.ToDecimal(va); } catch { }
                    try { if (vb is Microsoft.Xrm.Sdk.Money) db = ((Microsoft.Xrm.Sdk.Money)vb).Value; else if (vb is IConvertible) db = Convert.ToDecimal(vb); } catch { }
                    return sign * da.CompareTo(db);

                case ColKind.Date:
                    DateTime ta = DateTime.MinValue, tb = DateTime.MinValue;
                    try { if (va != null) ta = Convert.ToDateTime(va); } catch { }
                    try { if (vb != null) tb = Convert.ToDateTime(vb); } catch { }
                    return sign * ta.CompareTo(tb);

                default:
                    var sa = (va == null) ? "" : FormatRawValue(_currentColumns[colIndex].EntityLogicalName, _currentColumns[colIndex].Attribute, va);
                    var sb = (vb == null) ? "" : FormatRawValue(_currentColumns[colIndex].EntityLogicalName, _currentColumns[colIndex].Attribute, vb);
                    return sign * StringComparer.CurrentCultureIgnoreCase.Compare(sa, sb);
            }
        }

        private void ApplySort()
        {
            if (_sortCol < 0 || _sortOrder == SortOrder.None || _currentResults == null || _currentResults.Count == 0) return;
            _currentResults = _currentResults
                .OrderBy(x => x, Comparer<Entity>.Create((x, y) => CompareEntitiesForColumn(x, y, _sortCol)))
                .ToList();
        }

        private void UpdateSortHeaderGlyphs()
        {
            if (this.lvResults == null || this.lvResults.Columns == null) return;
            if (_baseHeaders == null || _baseHeaders.Count != this.lvResults.Columns.Count)
            {
                _baseHeaders = this.lvResults.Columns.Cast<ColumnHeader>().Select(ch => ch.Text).ToList();
            }
            for (int i = 0; i < this.lvResults.Columns.Count; i++)
            {
                string baseText = _baseHeaders[i];
                string arrow = "";
                if (i == _sortCol)
                {
                    arrow = _sortOrder == SortOrder.Ascending ? " ▲" :
                            _sortOrder == SortOrder.Descending ? " ▼" : "";
                }
                this.lvResults.Columns[i].Text = baseText + arrow;
            }
        }

        private void RenderCurrentResults()
        {
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

        private void lvResults_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (_currentResults == null || _currentResults.Count == 0) return;

            if (_sortCol != e.Column)
            {
                _sortCol = e.Column;
                // Ordre par défaut selon le type
                var sample = _currentResults[0];
                var kind = GetColKind(GetRawValue(sample, _currentColumns[_sortCol]));
                _sortOrder = (kind == ColKind.Text) ? SortOrder.Ascending : SortOrder.Descending;
            }
            else
            {
                _sortOrder = (_sortOrder == SortOrder.Ascending) ? SortOrder.Descending : SortOrder.Ascending;
            }

            ApplySort();
            RenderCurrentResults();
            UpdateSortHeaderGlyphs();
        }

        // ------------------------- Formatage & sélection -------------------------

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
            string label = null;
if (ent.Contains("name")) label = Convert.ToString(ent["name"]);
else if (ent.Contains("fullname")) label = Convert.ToString(ent["fullname"]);
else if (ent.Contains("new_name")) label = Convert.ToString(ent["new_name"]); // custom entity 'new_instance' uses 'new_name'
else if (ent.Contains("title")) label = Convert.ToString(ent["title"]);
if (!string.IsNullOrEmpty(label)) er.Name = label;
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
            if (this.cboRecent != null && this.cboRecent.SelectedIndex > 0)
            {
                var rl = this.cboRecent.SelectedItem as RecentLink;
                if (rl != null)
                {
                    this.SelectedReference = new Microsoft.Xrm.Sdk.EntityReference(rl.EntityLogicalName, rl.Id) { Name = rl.Name };
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                    return;
                }
            }

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
