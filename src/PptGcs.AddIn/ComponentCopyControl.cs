using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    internal sealed class ComponentCopyControl : GroupBox
    {
        private readonly ComboBox destination = new ComboBox();
        private readonly CheckBox share = new CheckBox();
        private readonly Label message = new Label();
        private readonly Label rules = new Label();
        private readonly Button apply = new Button();
        private int selectedCount;

        public event Action<int, bool> Copy;

        public ComponentCopyControl()
        {
            Text = AccessibleName = "复制对象及内部约束";
            Dock = DockStyle.Fill;
            AutoSize = true;
            Padding = new Padding(8, 14, 8, 8);
            Visible = false;
            var fields = new TableLayoutPanel { ColumnCount = 1, Dock = DockStyle.Top, AutoSize = true };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fields.Controls.Add(new Label { Text = "目标幻灯片", AutoSize = true });
            destination.DropDownStyle = ComboBoxStyle.DropDownList;
            destination.Dock = DockStyle.Fill;
            destination.AccessibleName = "目标幻灯片";
            destination.SelectedIndexChanged += (sender, args) => UpdateButton();
            fields.Controls.Add(destination);
            share.Text = share.AccessibleName = "与原件共享命名参数";
            share.AutoSize = true;
            fields.Controls.Add(share);
            rules.Text = "默认给副本创建独立的命名参数。勾选共享后，修改同一参数会同时影响原件和副本。只复制参与对象全部入选的约束。";
            rules.AutoSize = message.AutoSize = true;
            fields.Controls.Add(rules);
            message.AccessibleName = "组件复制结果";
            fields.Controls.Add(message);
            var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
            apply.Text = apply.AccessibleName = "复制所选对象";
            apply.AutoSize = true;
            apply.Click += (sender, args) => {
                var slide = destination.SelectedItem as ConstraintPaneSlide;
                if (slide != null && selectedCount > 0) Copy?.Invoke(slide.Id, share.Checked);
            };
            var cancel = new Button { Text = "取消", AutoSize = true };
            cancel.Click += (sender, args) => Visible = false;
            buttons.Controls.AddRange(new Control[] { apply, cancel });
            fields.Controls.Add(buttons);
            Controls.Add(fields);
            fields.SizeChanged += (sender, args) => Height = fields.Bottom + Padding.Bottom;
            SizeChanged += (sender, args) => rules.MaximumSize = message.MaximumSize = new Size(Math.Max(150, ClientSize.Width - Padding.Horizontal - 12), 0);
        }

        public void Open(IReadOnlyList<ConstraintPaneSlide> slides, int currentId, int count)
        {
            destination.Items.Clear();
            destination.Items.AddRange(slides.Cast<object>().ToArray());
            destination.SelectedItem = slides.FirstOrDefault(item => item.Id == currentId);
            share.Checked = false;
            ShowSelection(count);
            Visible = true;
            destination.Focus();
        }

        public void ShowSelection(int count)
        {
            if (selectedCount != count || !Visible) message.Text = "已选 " + count + " 个对象。";
            selectedCount = count;
            UpdateButton();
        }

        public void ShowError(string error) { message.Text = error; }

        private void UpdateButton() { apply.Enabled = selectedCount > 0 && destination.SelectedIndex >= 0; }
    }
}
