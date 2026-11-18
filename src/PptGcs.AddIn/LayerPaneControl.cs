using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    internal sealed class LayerPaneControl : UserControl
    {
        private readonly ListView list = new ListView {
            Dock = DockStyle.Fill, View = View.Details, HeaderStyle = ColumnHeaderStyle.Nonclickable,
            FullRowSelect = true, MultiSelect = false, HideSelection = false, AllowDrop = true,
            AccessibleName = "幻灯片图层（从前到后）"
        };
        private readonly TextBox name = new TextBox { Width = 150, MaxLength = 40, AccessibleName = "图层名称" };
        private readonly Button add, rename, remove, up, down, assign, marquee;
        private readonly Label status = new Label { AutoSize = true, Dock = DockStyle.Fill, Padding = new Padding(5) };
        private bool rendering;
        private string[] ids = new string[0];
        private string signature;
        internal Action<string> ActivateLayer, AddLayer, RenameLayer;
        internal Action<string> RemoveLayer;
        internal Action<int> ReorderLayer;
        internal Action AssignSelection, ArmMarquee;

        internal LayerPaneControl()
        {
            Dock = DockStyle.Fill;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(5) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            list.Columns.Add("图层（从前到后）", 200);
            list.Columns.Add("对象", 55, HorizontalAlignment.Right);
            list.Resize += (sender, args) => list.Columns[0].Width = Math.Max(96, list.ClientSize.Width - 74);
            list.SelectedIndexChanged += (sender, args) =>
            {
                if (rendering || list.SelectedIndices.Count == 0) return;
                Perform(() => ActivateLayer(ids[list.SelectedIndices[0]]));
            };
            list.ItemDrag += (sender, args) =>
            {
                if (list.Items.Count > 1 && args.Button == MouseButtons.Left)
                    list.DoDragDrop((string)((ListViewItem)args.Item).Tag, DragDropEffects.Move);
            };
            list.DragOver += (sender, args) =>
            {
                string id = args.Data.GetData(typeof(string)) as string;
                var point = list.PointToClient(new Point(args.X, args.Y));
                ListViewItem target = list.GetItemAt(point.X, point.Y);
                bool valid = target != null && id != null && Array.IndexOf(ids, id) >= 0;
                args.Effect = valid ? DragDropEffects.Move : DragDropEffects.None;
                list.InsertionMark.Index = valid ? target.Index : -1;
                list.InsertionMark.AppearsAfterItem = valid && point.Y >= target.Bounds.Top + target.Bounds.Height / 2;
            };
            list.DragLeave += (sender, args) => list.InsertionMark.Index = -1;
            list.DragDrop += (sender, args) =>
            {
                string id = args.Data.GetData(typeof(string)) as string;
                int from = Array.IndexOf(ids, id);
                int position = list.InsertionMark.Index;
                bool after = list.InsertionMark.AppearsAfterItem;
                list.InsertionMark.Index = -1;
                if (from < 0 || position < 0) return;
                int insert = position + (after ? 1 : 0);
                Perform(() =>
                {
                    ActivateLayer(id);
                    ReorderLayer(insert > from ? insert - 1 : insert);
                });
            };

            var editing = Row();
            editing.Controls.Add(name);
            add = Button(editing, "新建图层", () => { AddLayer(name.Text); name.Clear(); });
            rename = Button(editing, "重命名", () => { RenameLayer(name.Text); name.Clear(); });
            var actions = Row();
            up = Button(actions, "上移", () => ReorderLayer(list.SelectedIndices[0] - 1));
            down = Button(actions, "下移", () => ReorderLayer(list.SelectedIndices[0] + 1));
            remove = Button(actions, "合并到相邻层", () => RemoveLayer(ids[list.SelectedIndices[0]]));
            assign = Button(actions, "选中对象移入此层", () => AssignSelection());
            marquee = Button(actions, "框选当前层", () => ArmMarquee());
            layout.Controls.Add(editing, 0, 0);
            layout.Controls.Add(actions, 0, 1);
            layout.Controls.Add(status, 0, 2);
            layout.Controls.Add(list, 0, 3);
            Controls.Add(layout);
        }

        internal void Render(SlideLayers slide, string active, bool armed)
        {
            var layers = slide == null ? new SlideLayers.Layer[0] : slide.Stack;
            var members = slide == null ? new SlideLayers.Member[0] : slide.Members();
            var selected = slide == null ? new System.Collections.Generic.HashSet<int>() : slide.Selected();
            var counts = members.ToLookup(member => member.LayerId);
            string next = String.Join("\n", layers.Select(layer => layer.Id + ":" + layer.Name + ":" + counts[layer.Id].Count())) + "\n" + active;
            if (next != signature)
            {
                rendering = true;
                list.BeginUpdate();
                try
                {
                    ids = layers.Select(layer => layer.Id).ToArray();
                    list.Items.Clear();
                    foreach (var layer in layers)
                    {
                        var item = new ListViewItem(new[] { layer.Name, counts[layer.Id].Count().ToString() }) { Tag = layer.Id };
                        list.Items.Add(item);
                        item.Selected = layer.Id == active;
                    }
                    signature = next;
                }
                finally { list.EndUpdate(); rendering = false; }
            }
            int index = Array.IndexOf(ids, active);
            add.Enabled = slide != null;
            rename.Enabled = index >= 0;
            remove.Enabled = index >= 0 && ids.Length > 1;
            up.Enabled = index > 0;
            down.Enabled = index >= 0 && index + 1 < ids.Length;
            assign.Enabled = index >= 0 && members.Any(member => member.LayerId != active && selected.Contains(member.Id));
            marquee.Enabled = index >= 0;
            marquee.Text = armed ? "取消框选" : "框选当前层";
            status.Text = slide == null ? "请切换到普通幻灯片编辑视图。"
                : armed ? "在画布上按住左键拖动。按 Esc 取消。"
                : "框选只选择当前层的对象。新对象进入最上层。";
            status.MaximumSize = new Size(Math.Max(100, Width - 15), 0);
        }

        private static FlowLayoutPanel Row()
        { return new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true }; }

        private Button Button(FlowLayoutPanel row, string label, Action command)
        {
            var button = new Button { Text = label, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            button.Click += (sender, args) => Perform(command);
            row.Controls.Add(button);
            return button;
        }

        private void Perform(Action action)
        {
            try { action(); }
            catch (Exception error) { MessageBox.Show(this, error.Message, "PPT GCS", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
    }
}
