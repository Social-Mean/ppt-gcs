using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    // Created normally by managed code, then attached to the native ATL pane.
    // Keep all document operations in the supplied shared command callbacks.
    public sealed class ConstraintPaneControl : UserControl
    {
        private readonly Label pageLabel = new Label();
        private readonly Label messageLabel = new Label();
        private readonly ComboBox scopeBox = new ComboBox();
        private readonly ListView objects = new ListView();
        private readonly ListView relations = new ListView();
        private readonly Button addButton = new Button();
        private readonly Button editButton = new Button();
        private readonly Button enableButton = new Button();
        private readonly Button deleteButton = new Button();
        private readonly GroupBox editor = new GroupBox();
        private readonly ComboBox optionBox = new ComboBox();
        private readonly Label editorHint = new Label();
        private readonly Label valueLabel = new Label();
        private readonly Label secondLabel = new Label();
        private readonly NumericUpDown valueBox = new NumericUpDown();
        private readonly NumericUpDown secondBox = new NumericUpDown();
        private readonly Button applyButton = new Button();
        private readonly ToolTip tips = new ToolTip();
        private ConstraintPaneActions actions;
        private ConstraintPaneModel model = new ConstraintPaneModel();
        private bool rendering;
        private bool selectionQueued;
        private int selectionGeneration;
        private string editingId;
        private string selectedRelationId;
        private string listSignature;
        private string operationMessage;

        public ConstraintPaneControl()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = SystemFonts.MessageBoxFont;
            BackColor = SystemColors.Window;
            ForeColor = SystemColors.WindowText;
            MinimumSize = new Size(250, 250);
            Size = new Size(360, 720);
            AutoScroll = true;
            AccessibleName = "对象与约束窗格";
            Padding = new Padding(10);

            var content = new TableLayoutPanel {
                Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1,
                Margin = Padding.Empty, Padding = Padding.Empty
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(content);

            pageLabel.AutoSize = true;
            pageLabel.Font = new Font(Font, FontStyle.Bold);
            pageLabel.Margin = new Padding(0, 0, 0, 8);
            pageLabel.AccessibleName = "当前页面与选择";
            content.Controls.Add(pageLabel);

            var toolbar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false, Margin = new Padding(0, 0, 0, 4) };
            scopeBox.DropDownStyle = ComboBoxStyle.DropDownList;
            scopeBox.Items.AddRange(new object[] { "当前页全部", "仅所选对象" });
            scopeBox.SelectedIndex = 0;
            scopeBox.Width = 150;
            scopeBox.AccessibleName = "约束显示范围";
            scopeBox.SelectedIndexChanged += (sender, args) => RefreshLists();
            toolbar.Controls.Add(scopeBox);
            var refresh = MakeButton("刷新", 60, () => InvokeCommand(() => actions.Refresh?.Invoke()));
            refresh.AccessibleName = "刷新对象和约束";
            toolbar.Controls.Add(refresh);
            content.Controls.Add(toolbar);

            content.Controls.Add(SectionLabel("对象"));
            ConfigureList(objects, true, 122, "当前页对象");
            objects.Columns.Add("对象", 250);
            objects.Columns.Add("关系", 52);
            objects.SelectedIndexChanged += ObjectsSelected;
            content.Controls.Add(objects);

            content.Controls.Add(SectionLabel("约束 · 点击可定位参与对象"));
            ConfigureList(relations, false, 186, "对象之间的约束");
            relations.Columns.Add("关系", 100);
            relations.Columns.Add("参与对象", 142);
            relations.Columns.Add("状态", 60);
            relations.SelectedIndexChanged += RelationSelected;
            relations.DoubleClick += (sender, args) => EditSelected();
            content.Controls.Add(relations);

            var commands = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 6, 0, 6) };
            ConfigureButton(addButton, "新增…", 72, () => ShowEditor());
            ConfigureButton(editButton, "修改…", 72, EditSelected);
            ConfigureButton(enableButton, "停用", 60, ToggleSelected);
            ConfigureButton(deleteButton, "删除", 60, DeleteSelected);
            commands.Controls.AddRange(new Control[] { addButton, editButton, enableButton, deleteButton });
            tips.SetToolTip(addButton, "对画布中选中的对象建立关系");
            tips.SetToolTip(editButton, "修改当前关系的数值参数");
            tips.SetToolTip(enableButton, "暂停或恢复当前关系");
            tips.SetToolTip(deleteButton, "只删除当前关系，保留其他关系");
            content.Controls.Add(commands);

            BuildEditor();
            content.Controls.Add(editor);
            messageLabel.AutoSize = true;
            messageLabel.Dock = DockStyle.Fill;
            messageLabel.Margin = new Padding(0, 8, 0, 0);
            messageLabel.AccessibleName = "约束操作状态";
            content.Controls.Add(messageLabel);
            Resize += (sender, args) => ResizeColumns();
            RefreshLists();
        }

        public void Bind(ConstraintPaneActions callbacks)
        {
            actions = callbacks ?? throw new ArgumentNullException("callbacks");
            RefreshButtons();
        }

        public void Render(ConstraintPaneModel snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (!String.Equals(model.PageLabel, snapshot.PageLabel, StringComparison.Ordinal))
            {
                // A pane belongs to a window, so its active page can change.
                // Never carry an unfinished operation to a different page.
                editingId = null;
                selectedRelationId = null;
                editor.Visible = false;
                operationMessage = null;
                listSignature = null;
                selectionGeneration++;
                selectionQueued = false;
            }
            model = snapshot;
            int selected = model.Objects.Count(item => item.Selected);
            pageLabel.Text = (model.PageLabel ?? "当前页") + " · 已选 " + selected;
            RefreshMessage();
            if (editingId != null && !model.Relations.Any(item => item.Id == editingId))
            { editor.Visible = false; editingId = null; }
            RefreshLists();
            if (editor.Visible) UpdateEditorHint();
        }

        // The wheel's "more constraints" action can open this exact editor.
        public void ShowEditor(string kind = null)
        {
            operationMessage = null;
            RefreshMessage();
            editingId = null;
            editor.Text = "新建约束";
            optionBox.Enabled = true;
            PopulateOptions(kind);
            applyButton.Text = "创建约束";
            editor.Visible = true;
            LoadEditorValues(null);
            ScrollControlIntoView(editor);
            optionBox.Focus();
        }

        public void ShowMessage(string message)
        { operationMessage = message; RefreshMessage(); }

        private void RefreshMessage()
        { messageLabel.Text = operationMessage ?? model.Message ?? String.Empty; }

        private void BuildEditor()
        {
            editor.Text = "新建约束";
            editor.Dock = DockStyle.Fill;
            editor.AutoSize = true;
            editor.Visible = false;
            editor.Padding = new Padding(8, 14, 8, 8);
            editor.AccessibleName = "约束参数编辑";
            var fields = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 2, Margin = Padding.Empty };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            optionBox.DropDownStyle = ComboBoxStyle.DropDownList;
            optionBox.Dock = DockStyle.Fill;
            optionBox.AccessibleName = "约束类型";
            optionBox.SelectedIndexChanged += (sender, args) => { if (!rendering) LoadEditorValues(null); };
            fields.Controls.Add(optionBox, 0, 0); fields.SetColumnSpan(optionBox, 2);
            editorHint.AutoSize = true; editorHint.Dock = DockStyle.Fill;
            editorHint.Margin = new Padding(0, 6, 0, 6);
            fields.Controls.Add(editorHint, 0, 1); fields.SetColumnSpan(editorHint, 2);
            ConfigureNumeric(valueBox, "第一个约束参数");
            ConfigureNumeric(secondBox, "第二个约束参数");
            valueLabel.AutoSize = true; valueLabel.Anchor = AnchorStyles.Left;
            secondLabel.AutoSize = true; secondLabel.Anchor = AnchorStyles.Left;
            fields.Controls.Add(valueLabel, 0, 2); fields.Controls.Add(valueBox, 1, 2);
            fields.Controls.Add(secondLabel, 0, 3); fields.Controls.Add(secondBox, 1, 3);
            var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 8, 0, 0) };
            ConfigureButton(applyButton, "创建约束", 88, ApplyEditor);
            buttons.Controls.Add(applyButton);
            buttons.Controls.Add(MakeButton("取消", 60, () => { editor.Visible = false; editingId = null; }));
            fields.Controls.Add(buttons, 0, 4); fields.SetColumnSpan(buttons, 2);
            editor.Controls.Add(fields);
        }

        private void RefreshLists()
        {
            if (rendering) return;
            var selected = new HashSet<string>(model.Objects.Where(item => item.Selected).Select(item => item.Id));
            bool filtered = scopeBox.SelectedIndex == 1;
            var shownObjects = model.Objects.Where(item => !filtered || item.Selected).ToList();
            var shownRelations = model.Relations.Where(item => !filtered || item.MemberIds.Any(selected.Contains)).ToList();
            string signature = scopeBox.SelectedIndex + "|" + String.Join(";", shownObjects.Select(item => item.Id + ":" + item.Name + ":" + item.RelationCount + ":" + item.Selected)) + "|" +
                String.Join(";", shownRelations.Select(item => item.Id + ":" + RelationTitle(item) + ":" + String.Join("、", item.MemberNames) + ":" + item.Enabled + ":" + item.Status));
            if (signature != listSignature)
            {
                rendering = true;
                try
                {
                    objects.BeginUpdate(); relations.BeginUpdate();
                    objects.Items.Clear(); relations.Items.Clear();
                    foreach (var item in shownObjects)
                    {
                        var row = new ListViewItem(item.Name ?? item.Id) { Tag = item, Selected = item.Selected };
                        row.SubItems.Add(item.RelationCount.ToString(CultureInfo.InvariantCulture));
                        objects.Items.Add(row);
                    }
                    foreach (var item in shownRelations)
                    {
                        string members = String.Join(" → ", item.MemberNames);
                        var row = new ListViewItem(RelationTitle(item)) { Tag = item, Selected = item.Id == selectedRelationId };
                        row.SubItems.Add(members);
                        row.SubItems.Add(item.Enabled ? (item.Status ?? "已满足") : "已停用");
                        row.ToolTipText = RelationTitle(item) + Environment.NewLine + members + Environment.NewLine + (item.Status ?? String.Empty);
                        if (!item.Enabled) row.ForeColor = SystemColors.GrayText;
                        relations.Items.Add(row);
                    }
                    if (!shownRelations.Any(item => item.Id == selectedRelationId)) selectedRelationId = null;
                    listSignature = signature;
                }
                finally { objects.EndUpdate(); relations.EndUpdate(); rendering = false; }
            }
            RefreshButtons(); ResizeColumns();
        }

        private void ObjectsSelected(object sender, EventArgs args)
        {
            if (rendering || actions == null || selectionQueued || !IsHandleCreated) return;
            // ListView sends a deselection notification before the new row is
            // selected. Coalesce the native sequence before updating PowerPoint;
            // a synchronous round-trip would rebuild the list mid-click.
            selectionQueued = true;
            int generation = selectionGeneration;
            BeginInvoke((Action)(() => {
                if (generation != selectionGeneration) return;
                selectionQueued = false;
                if (IsDisposed || actions == null) return;
                var ids = objects.SelectedItems.Cast<ListViewItem>().Select(item => ((ConstraintPaneObject)item.Tag).Id).ToList().AsReadOnly();
                InvokeCommand(() => actions.SelectShapes?.Invoke(ids));
            }));
        }

        private void RelationSelected(object sender, EventArgs args)
        {
            if (rendering) return;
            var relation = SelectedRelation();
            selectedRelationId = relation == null ? null : relation.Id;
            RefreshButtons();
            if (relation != null && actions != null)
                InvokeCommand(() => actions.SelectShapes?.Invoke(relation.MemberIds.ToList().AsReadOnly()));
        }

        private ConstraintPaneRelation SelectedRelation()
        {
            if (relations.SelectedItems.Count == 0) return null;
            string id = ((ConstraintPaneRelation)relations.SelectedItems[0].Tag).Id;
            return model.Relations.FirstOrDefault(item => item.Id == id);
        }

        private void RefreshButtons()
        {
            var relation = SelectedRelation();
            var option = relation == null ? null : model.Options.FirstOrDefault(item => item.Kind == relation.Kind);
            addButton.Enabled = actions != null && model.Options.Count > 0;
            editButton.Enabled = actions != null && relation != null && option != null && (!String.IsNullOrEmpty(option.ValueLabel) || !String.IsNullOrEmpty(option.SecondLabel));
            enableButton.Enabled = deleteButton.Enabled = actions != null && relation != null;
            enableButton.Text = relation != null && !relation.Enabled ? "启用" : "停用";
            enableButton.AccessibleName = enableButton.Text;
        }

        private void EditSelected()
        {
            var relation = SelectedRelation();
            if (relation == null) return;
            operationMessage = null;
            RefreshMessage();
            editingId = relation.Id;
            editor.Text = "修改约束参数";
            PopulateOptions(relation.Kind);
            optionBox.Enabled = false;
            applyButton.Text = "保存修改";
            editor.Visible = true;
            LoadEditorValues(relation);
            ScrollControlIntoView(editor);
            if (valueBox.Visible) valueBox.Focus();
        }

        private void PopulateOptions(string kind)
        {
            rendering = true;
            try
            {
                optionBox.Items.Clear();
                foreach (var option in model.Options) optionBox.Items.Add(option);
                int index = model.Options.ToList().FindIndex(option => option.Kind == kind);
                optionBox.SelectedIndex = index >= 0 ? index : (optionBox.Items.Count > 0 ? 0 : -1);
            }
            finally { rendering = false; }
        }

        private ConstraintPaneOption EditorOption()
        {
            var item = optionBox.SelectedItem as ConstraintPaneOption;
            return item == null ? null : model.Options.FirstOrDefault(option => option.Kind == item.Kind);
        }

        private void LoadEditorValues(ConstraintPaneRelation relation)
        {
            var option = EditorOption();
            valueLabel.Text = option == null ? String.Empty : option.ValueLabel;
            secondLabel.Text = option == null ? String.Empty : option.SecondLabel;
            valueLabel.Visible = valueBox.Visible = option != null && !String.IsNullOrEmpty(option.ValueLabel);
            secondLabel.Visible = secondBox.Visible = option != null && !String.IsNullOrEmpty(option.SecondLabel);
            if (option != null)
            {
                valueBox.Increment = option.Kind == "AspectRatio" ? 0.01m : 1m;
                SetNumber(valueBox, option.MinimumValue, option.MaximumValue, relation == null ? option.DefaultValue : relation.Value);
                SetNumber(secondBox, option.MinimumSecond, option.MaximumSecond, relation == null ? option.DefaultSecond : relation.Second);
            }
            UpdateEditorHint();
        }

        private void UpdateEditorHint()
        {
            var option = EditorOption();
            int count = model.Objects.Count(item => item.Selected);
            bool eligible = option != null && count >= option.MinimumMembers && (option.MaximumMembers <= 0 || count <= option.MaximumMembers);
            if (option == null) editorHint.Text = "请先选择约束类型。";
            else
            {
                string required = option.MaximumMembers == option.MinimumMembers ? option.MinimumMembers.ToString() : "至少 " + option.MinimumMembers;
                editorHint.Text = (option.Description ?? String.Empty) + (editingId == null ? Environment.NewLine + "需要 " + required + " 个对象，当前已选 " + count + " 个。" : String.Empty);
            }
            applyButton.Enabled = actions != null && option != null && (editingId != null || eligible);
        }

        private void ApplyEditor()
        {
            var option = EditorOption();
            if (option == null || actions == null) return;
            double? first = String.IsNullOrEmpty(option.ValueLabel) ? (double?)null : (double)valueBox.Value;
            double? second = String.IsNullOrEmpty(option.SecondLabel) ? (double?)null : (double)secondBox.Value;
            string id = editingId;
            if (InvokeCommand(() => { if (id == null) actions.Add?.Invoke(option.Kind, first, second); else actions.Update?.Invoke(id, first, second); }))
            { editor.Visible = false; editingId = null; }
        }

        private void ToggleSelected()
        {
            var relation = SelectedRelation();
            if (relation != null && actions != null) InvokeCommand(() => actions.SetEnabled?.Invoke(relation.Id, !relation.Enabled));
        }

        private void DeleteSelected()
        {
            var relation = SelectedRelation();
            if (relation != null && actions != null && InvokeCommand(() => actions.Delete?.Invoke(relation.Id)))
            { selectedRelationId = null; editor.Visible = false; editingId = null; }
        }

        private bool InvokeCommand(Action command)
        {
            if (actions == null) return false;
            operationMessage = null;
            RefreshMessage();
            try { command(); RefreshMessage(); return true; }
            catch (Exception exception)
            {
                // Host commands refresh the model in their finally blocks and
                // the idle timer renders it again. Keep this failure separate
                // from normal page status until the next deliberate operation.
                operationMessage = exception.Message;
                RefreshMessage();
                return false;
            }
        }

        private static string RelationTitle(ConstraintPaneRelation relation)
        {
            string label = relation.Label ?? relation.Kind;
            if (relation.Value.HasValue) label += " " + relation.Value.Value.ToString("0.###", CultureInfo.CurrentCulture);
            if (relation.Second.HasValue) label += ", " + relation.Second.Value.ToString("0.###", CultureInfo.CurrentCulture);
            return label;
        }

        private void ResizeColumns()
        {
            if (objects.Columns.Count > 1) objects.Columns[0].Width = Math.Max(120, objects.ClientSize.Width - 64);
            if (relations.Columns.Count > 2) relations.Columns[1].Width = Math.Max(80, relations.ClientSize.Width - 176);
        }

        private static Label SectionLabel(string text)
        { return new Label { Text = text, AutoSize = true, Margin = new Padding(0, 7, 0, 5) }; }
        private static void ConfigureList(ListView list, bool multiple, int height, string name)
        {
            list.View = View.Details; list.FullRowSelect = true; list.HideSelection = false;
            list.MultiSelect = multiple; list.ShowItemToolTips = true; list.Dock = DockStyle.Fill;
            list.Height = height; list.MinimumSize = new Size(200, height); list.Margin = Padding.Empty;
            list.AccessibleName = name;
        }
        private static void ConfigureNumeric(NumericUpDown box, string name)
        {
            box.DecimalPlaces = 3; box.Increment = 1; box.Dock = DockStyle.Fill;
            box.Minimum = -1000000; box.Maximum = 1000000; box.AccessibleName = name;
        }
        private static void SetNumber(NumericUpDown box, double minimum, double maximum, double? value)
        {
            box.Minimum = (decimal)minimum; box.Maximum = (decimal)maximum;
            decimal number = (decimal)(value ?? 0);
            box.Value = Math.Max(box.Minimum, Math.Min(box.Maximum, number));
        }
        private static Button MakeButton(string text, int width, Action click)
        { var result = new Button(); ConfigureButton(result, text, width, click); return result; }
        private static void ConfigureButton(Button button, string text, int width, Action click)
        {
            button.Text = text; button.Width = width; button.Height = 28;
            button.Margin = new Padding(0, 0, 5, 0); button.AccessibleName = text;
            button.Click += (sender, args) => click();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                selectionGeneration++;
                selectionQueued = false;
                actions = null;
                tips.Dispose();
                if (pageLabel.Font != Font) pageLabel.Font.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
