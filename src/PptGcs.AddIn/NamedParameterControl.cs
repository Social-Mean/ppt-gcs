using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    internal sealed class NamedParameterControl : GroupBox
    {
        private readonly ComboBox library = new ComboBox();
        private readonly TextBox nameBox = new TextBox();
        private readonly NumericUpDown valueBox = new NumericUpDown();
        private readonly Label message = new Label();
        private readonly Button save = new Button();
        private readonly Button delete = new Button();
        private bool rendering;

        private ConstraintPaneParameter Selected { get { return library.SelectedItem as ConstraintPaneParameter; } }
        public event Action<string, string, double> Save;
        public event Action<string> Delete;

        public NamedParameterControl()
        {
            Text = AccessibleName = "命名参数";
            Dock = DockStyle.Fill;
            AutoSize = true;
            Padding = new Padding(8, 14, 8, 8);
            Visible = false;
            var fields = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Top, AutoSize = true };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            library.DropDownStyle = ComboBoxStyle.DropDownList;
            library.Dock = DockStyle.Fill;
            library.AccessibleName = "文稿参数";
            library.SelectedIndexChanged += (sender, args) => { if (!rendering) LoadSelection(); };
            fields.Controls.Add(library, 0, 0);
            fields.SetColumnSpan(library, 2);
            fields.Controls.Add(new Label { Text = "名称", AutoSize = true }, 0, 1);
            nameBox.AccessibleName = "参数名称";
            nameBox.MaxLength = LayoutParameter.MaximumNameLength;
            nameBox.Dock = DockStyle.Fill;
            nameBox.TextChanged += (sender, args) => UpdateButtons();
            fields.Controls.Add(nameBox, 1, 1);
            fields.Controls.Add(new Label { Text = "数值 (pt)", AutoSize = true }, 0, 2);
            valueBox.AccessibleName = "命名参数数值 (pt)";
            valueBox.Minimum = -(decimal)LayoutParameter.MaximumMagnitude;
            valueBox.Maximum = (decimal)LayoutParameter.MaximumMagnitude;
            valueBox.DecimalPlaces = 3;
            valueBox.Dock = DockStyle.Fill;
            fields.Controls.Add(valueBox, 1, 2);
            message.AutoSize = true;
            message.AccessibleName = "参数引用和操作结果";
            fields.Controls.Add(message, 0, 3);
            fields.SetColumnSpan(message, 2);
            var actions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            var create = new Button { Text = "新建…", AccessibleName = "新建命名参数", AutoSize = true };
            create.Click += (sender, args) => { Select(null); nameBox.Focus(); };
            save.AutoSize = true;
            save.Click += (sender, args) => Save?.Invoke(Selected?.Id, nameBox.Text, (double)valueBox.Value);
            delete.Text = "删除";
            delete.AccessibleName = "删除未使用的参数";
            delete.AutoSize = true;
            delete.Click += (sender, args) => { if (Selected?.References == 0) Delete?.Invoke(Selected.Id); };
            var close = new Button { Text = "关闭", AutoSize = true };
            close.Click += (sender, args) => Visible = false;
            actions.Controls.AddRange(new Control[] { create, save, delete, close });
            fields.Controls.Add(actions, 0, 4);
            fields.SetColumnSpan(actions, 2);
            Controls.Add(fields);
            fields.SizeChanged += (sender, args) => Height = fields.Bottom + Padding.Bottom;
            SizeChanged += (sender, args) => message.MaximumSize = new Size(Math.Max(150, ClientSize.Width - Padding.Horizontal - 12), 0);
        }

        public void Open(IList<ConstraintPaneParameter> parameters)
        {
            Render(parameters);
            Select(parameters.FirstOrDefault()?.Id);
            Visible = true;
        }

        public void Render(IList<ConstraintPaneParameter> parameters)
        {
            var existing = library.Items.Cast<ConstraintPaneParameter>().ToList();
            if (existing.Count == parameters.Count && existing.Zip(parameters,(a,b)=>a.Id==b.Id && a.Name==b.Name &&
                a.Value==b.Value && a.References==b.References).All(same=>same)) return;
            var previous = Selected;
            rendering = true;
            try
            {
                library.Items.Clear();
                library.Items.AddRange(parameters.Cast<object>().ToArray());
                SetSelection(previous?.Id);
            }
            finally { rendering = false; }
            if (previous?.Name != Selected?.Name || previous?.Id != Selected?.Id || previous?.Value != Selected?.Value) LoadSelection();
            else UpdateButtons();
        }

        public void Select(string id)
        {
            SetSelection(id);
            LoadSelection();
        }

        private void SetSelection(string id)
        { library.SelectedItem = library.Items.Cast<ConstraintPaneParameter>().FirstOrDefault(item => item.Id == id); }

        public void ShowError(string error) { message.Text = error; }

        private void LoadSelection()
        {
            nameBox.Text = Selected?.Name ?? String.Empty;
            valueBox.Value = (decimal)(Selected?.Value ?? 0);
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            save.Text = save.AccessibleName = Selected == null ? "添加参数" : "保存参数";
            save.Enabled = !String.IsNullOrWhiteSpace(nameBox.Text);
            delete.Enabled = Selected?.References == 0;
            message.Text = Selected == null ? "新建参数后可在约束编辑器中选用。" :
                "全篇文稿有 " + Selected.References + " 条关系引用此参数。";
        }
    }
}
