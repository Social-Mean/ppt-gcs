using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    internal enum LayerRowState { Empty, Off, On, Mixed }

    internal sealed class LayerPaneRow
    {
        public readonly int SlideId;
        public readonly string LayerId, Title, Key;
        public readonly int? ShapeId;
        public readonly LayerRowState Visible, Locked;
        public readonly LayerPaneRow[] Children;

        public LayerPaneRow(int slideId, string layerId, int? shapeId, string title, LayerRowState visible,
            LayerRowState locked, LayerPaneRow[] children)
        {
            SlideId = slideId; LayerId = layerId; ShapeId = shapeId; Title = title;
            Visible = visible; Locked = locked; Children = children;
            Key = shapeId.HasValue ? "shape:" + shapeId.Value : "layer:" + layerId;
        }
    }

    internal sealed class LayerPaneControl : UserControl
    {
        private readonly LayerTree tree = new LayerTree {
            Dock = DockStyle.Fill, DrawMode = TreeViewDrawMode.OwnerDrawAll, FullRowSelect = true,
            HideSelection = false, ShowLines = false, AllowDrop = true, AccessibleName = "图层和对象"
        };
        private readonly TextBox name = new TextBox { Width = 136, MaxLength = 40, AccessibleName = "图层名称" };
        private readonly ToolTip tips = new ToolTip();
        private readonly Button add, rename, remove, up, down, assign;
        private readonly Label empty = new Label { Text = "请切换到普通幻灯片编辑视图。", AutoSize = true, Padding = new Padding(5) };
        private readonly Dictionary<string, TreeNode> nodes = new Dictionary<string, TreeNode>();
        private HashSet<int> selected = new HashSet<int>();
        private string[] layers = Array.Empty<string>();
        private string active, structure, dragLayer;
        private int slideId;
        private bool rendering;
        private TreeNode dropTarget;
        private bool dropAfter;
        private Font layerFont;

        internal Action<LayerPaneRow> ActivateRow;
        internal Action<LayerPaneRow, bool> ToggleState;
        internal Action<string> AddLayer, RenameLayer, RemoveLayer;
        internal Action<string, int> ReorderLayer;
        internal Action AssignSelection;

        internal LayerPaneControl()
        {
            Dock = DockStyle.Fill;
            Font = SystemFonts.MessageBoxFont;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(5) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var editing = Row();
            editing.Controls.Add(name);
            add = AddButton(editing, "新建图层", () => { AddLayer(name.Text); name.Clear(); });
            rename = AddButton(editing, "重命名", () => { RenameLayer(name.Text); name.Clear(); });
            var actions = Row();
            up = AddButton(actions, "上移", () => ReorderLayer(active, Array.IndexOf(layers, active) - 1));
            down = AddButton(actions, "下移", () => ReorderLayer(active, Array.IndexOf(layers, active) + 1));
            remove = AddButton(actions, "合并图层", () => RemoveLayer(active));
            assign = AddButton(actions, "移入此层", () => AssignSelection());
            tips.SetToolTip(remove, "将当前层合并到相邻层");
            tips.SetToolTip(assign, "将画布所选对象移入当前层");
            layout.Controls.Add(editing, 0, 0);
            layout.Controls.Add(actions, 0, 1);
            layout.Controls.Add(empty, 0, 2);
            layout.Controls.Add(tree, 0, 3);
            Controls.Add(layout);

            tree.ItemHeight = Math.Max(24, Font.Height + 9);
            tree.Indent = Math.Max(18, Font.Height + 3);
            layerFont = new Font(tree.Font, FontStyle.Bold);
            tree.DrawNode += DrawRow;
            tree.HitIcon = point =>
            {
                int column = Column(point.X);
                if (column == 0) return null;
                var node = NodeAt(point.Y);
                var row = node?.Tag as LayerPaneRow;
                if (row == null || row.Visible == LayerRowState.Empty) return null;
                return () => Perform(() => ToggleState(row, column == 1));
            };
            tree.AfterSelect += (sender, args) =>
            {
                if (!rendering && (args.Action == TreeViewAction.ByKeyboard || args.Action == TreeViewAction.ByMouse))
                    Perform(() => ActivateRow((LayerPaneRow)args.Node.Tag));
            };
            tree.MouseUp += (sender, args) =>
            {
                if (args.Button != MouseButtons.Left) return;
                TreeNode node = NodeAt(args.Y);
                if (node == null) return;
                var row = (LayerPaneRow)node.Tag;
                if (Column(args.X) == 0 && (tree.HitTest(args.Location).Location & TreeViewHitTestLocations.PlusMinus) == 0)
                {
                    if (active == row.LayerId && (!row.ShapeId.HasValue || selected.Contains(row.ShapeId.Value))) return;
                    tree.SelectedNode = node;
                    Perform(() => ActivateRow(row));
                }
            };
            tree.MouseMove += (sender, args) =>
            {
                var node = NodeAt(args.Y);
                int column = Column(args.X);
                var row = node?.Tag as LayerPaneRow;
                string hint = row == null || row.Visible == LayerRowState.Empty || column == 0 ? "" :
                    column == 1 ? (row.Visible == LayerRowState.Off ? "显示" : "隐藏") + "（Ctrl+Shift+S）" :
                    (row.Locked == LayerRowState.On ? "解锁" : "锁定") + "（Ctrl+Shift+L）";
                if (tips.GetToolTip(tree) != hint) tips.SetToolTip(tree, hint);
            };
            tree.KeyDown += (sender, args) =>
            {
                var row = tree.SelectedNode?.Tag as LayerPaneRow;
                if (row == null || row.Visible == LayerRowState.Empty || !args.Control || !args.Shift ||
                    (args.KeyCode != Keys.S && args.KeyCode != Keys.L)) return;
                args.SuppressKeyPress = true;
                Perform(() => ToggleState(row, args.KeyCode == Keys.S));
            };
            tree.ItemDrag += (sender, args) =>
            {
                var row = ((TreeNode)args.Item).Tag as LayerPaneRow;
                if (layers.Length < 2 || row == null || row.ShapeId.HasValue || Column(tree.PointToClient(MousePosition).X) != 0 ||
                    args.Button != MouseButtons.Left) return;
                dragLayer = row.LayerId;
                try { tree.DoDragDrop(dragLayer, DragDropEffects.Move); }
                finally { dragLayer = null; ClearDrop(); }
            };
            tree.DragOver += (sender, args) =>
            {
                var point = tree.PointToClient(new Point(args.X, args.Y));
                var target = NodeAt(point.Y);
                bool allowed = target != null && target.Parent == null && dragLayer != null;
                args.Effect = allowed ? DragDropEffects.Move : DragDropEffects.None;
                dropTarget = allowed ? target : null;
                dropAfter = allowed && point.Y >= target.Bounds.Top + tree.ItemHeight / 2;
                tree.Invalidate();
            };
            tree.DragLeave += (sender, args) => ClearDrop();
            tree.DragDrop += (sender, args) =>
            {
                if (dropTarget == null || dragLayer == null) return;
                int from = Array.IndexOf(layers, dragLayer);
                int insert = dropTarget.Index + (dropAfter ? 1 : 0);
                if (from >= 0) Perform(() => ReorderLayer(dragLayer, insert > from ? insert - 1 : insert));
                ClearDrop();
            };
        }

        internal void Render(SlideLayers slide, string activeLayer)
        {
            var members = slide?.Members() ?? Array.Empty<SlideLayers.Member>();
            var roots = slide?.PaneRows(members) ?? Array.Empty<LayerPaneRow>();
            var ordered = Flatten(roots).ToArray();
            layers = roots.Select(row => row.LayerId).ToArray();
            int currentSlide = slide?.SlideId ?? 0;
            var nextSelection = slide?.Selected(true) ?? new HashSet<int>();
            string signature = String.Join("|", ordered.Select(row => row.Key + "/" + row.Children.Length));
            bool repaint = active != activeLayer || !selected.SetEquals(nextSelection);
            active = activeLayer;
            selected = nextSelection;
            rendering = true;
            try
            {
                if (slideId != currentSlide || structure != signature)
                {
                    if (dragLayer != null) return;
                    var collapsed = slideId == currentSlide ? new HashSet<string>(nodes.Where(pair => pair.Value.Nodes.Count > 0 &&
                        !pair.Value.IsExpanded).Select(pair => pair.Key)) : new HashSet<string>();
                    var expanded = slideId == currentSlide ? new HashSet<string>(nodes.Where(pair => pair.Value.IsExpanded)
                        .Select(pair => pair.Key)) : new HashSet<string>();
                    string focused = slideId == currentSlide ? (tree.SelectedNode?.Tag as LayerPaneRow)?.Key : null;
                    slideId = currentSlide;
                    structure = signature;
                    tree.BeginUpdate();
                    try
                    {
                        tree.Nodes.Clear();
                        nodes.Clear();
                        foreach (var row in roots) AddNode(tree.Nodes, row, collapsed, expanded);
                        TreeNode node;
                        if (focused != null && nodes.TryGetValue(focused, out node)) tree.SelectedNode = node;
                        else if (active != null && nodes.TryGetValue("layer:" + active, out node)) tree.SelectedNode = node;
                    }
                    finally { tree.EndUpdate(); }
                    repaint = true;
                }
                foreach (var row in ordered)
                {
                    var node = nodes[row.Key];
                    var previous = (LayerPaneRow)node.Tag;
                    if (row.Title != node.Text) { node.Text = row.Title; repaint = true; }
                    repaint |= row.Visible != previous.Visible || row.Locked != previous.Locked;
                    node.Tag = row;
                }
            }
            finally { rendering = false; }
            if (repaint) tree.Invalidate();
            int index = Array.IndexOf(layers, active);
            empty.Visible = slide == null;
            add.Enabled = slide != null;
            rename.Enabled = index >= 0;
            remove.Enabled = index >= 0 && layers.Length > 1;
            up.Enabled = index > 0;
            down.Enabled = index >= 0 && index < layers.Length - 1;
            var selectedRoots = slide?.Selected();
            assign.Enabled = index >= 0 && members.Any(member => member.LayerId != active && selectedRoots.Contains(member.Id));
        }

        private void AddNode(TreeNodeCollection collection, LayerPaneRow row, HashSet<string> collapsed, HashSet<string> expanded)
        {
            var node = collection.Add(row.Title);
            node.Tag = row;
            nodes.Add(row.Key, node);
            foreach (var child in row.Children) AddNode(node.Nodes, child, collapsed, expanded);
            if (row.ShapeId.HasValue ? expanded.Contains(row.Key) : !collapsed.Contains(row.Key)) node.Expand();
        }

        private static IEnumerable<LayerPaneRow> Flatten(IEnumerable<LayerPaneRow> rows)
        {
            foreach (var row in rows)
            {
                yield return row;
                foreach (var child in Flatten(row.Children)) yield return child;
            }
        }

        private void DrawRow(object sender, DrawTreeNodeEventArgs args)
        {
            if (args.Bounds.Height == 0) return;
            var row = args.Node.Tag as LayerPaneRow;
            if (row == null) return;
            var area = new Rectangle(0, args.Bounds.Top, tree.ClientSize.Width, tree.ItemHeight);
            bool layer = !row.ShapeId.HasValue;
            bool isActive = layer ? row.LayerId == active : selected.Contains(row.ShapeId.Value);
            bool isFocused = args.Node == tree.SelectedNode && tree.Focused;
            Color background = isFocused ? Color.FromArgb(219, 235, 252) : isActive ? Color.FromArgb(232, 240, 249) : tree.BackColor;
            using (var brush = new SolidBrush(background)) args.Graphics.FillRectangle(brush, area);
            var color = row.Visible == LayerRowState.Off ? SystemColors.GrayText : tree.ForeColor;
            int left = args.Node.Bounds.Left;
            int iconWidth = IconWidth;
            int right = area.Right - 2 * iconWidth - 3;
            if (args.Node.Nodes.Count > 0)
            {
                int centerY = area.Top + area.Height / 2, centerX = left - tree.Indent / 2;
                using (var pen = new Pen(color, 1.5f))
                {
                    var points = args.Node.IsExpanded
                        ? new[] { new Point(centerX - 3, centerY - 1), new Point(centerX, centerY + 2), new Point(centerX + 3, centerY - 1) }
                        : new[] { new Point(centerX - 1, centerY - 3), new Point(centerX + 2, centerY), new Point(centerX - 1, centerY + 3) };
                    args.Graphics.DrawLines(pen, points);
                }
            }
            var label = new Rectangle(left, area.Top, Math.Max(0, right - left), area.Height);
            TextRenderer.DrawText(args.Graphics, row.Title, layer ? layerFont : tree.Font, label, color,
                TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
            DrawIcon(args.Graphics, right + iconWidth / 2 + 3, area.Top + area.Height / 2, row.Visible, true);
            DrawIcon(args.Graphics, right + iconWidth * 3 / 2 + 3, area.Top + area.Height / 2, row.Locked, false);
            if (dropTarget == args.Node)
            {
                int y = dropAfter ? area.Bottom - 1 : area.Top;
                using (var pen = new Pen(SystemColors.Highlight, 2)) args.Graphics.DrawLine(pen, 0, y, area.Right, y);
            }
        }

        private void DrawIcon(Graphics graphics, int x, int y, LayerRowState state, bool eye)
        {
            if (state == LayerRowState.Empty) return;
            var saved = graphics.Save();
            try
            {
                graphics.TranslateTransform(x, y);
                float scale = tree.ItemHeight / 24f;
                graphics.ScaleTransform(scale, scale);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(state == LayerRowState.Off && eye ? SystemColors.GrayText : tree.ForeColor, 1.25f))
                {
                    pen.StartCap = pen.EndCap = LineCap.Round;
                    if (eye)
                    {
                        using (var outline = new GraphicsPath())
                        {
                            outline.AddBezier(-7, 0, -3, -5.6f, 3, -5.6f, 7, 0);
                            outline.AddBezier(7, 0, 3, 5.6f, -3, 5.6f, -7, 0);
                            graphics.DrawPath(pen, outline);
                        }
                        if (state != LayerRowState.Off) graphics.DrawEllipse(pen, -1.6f, -1.6f, 3.2f, 3.2f);
                        else graphics.DrawLine(pen, -6, 5, 6, -5);
                    }
                    else
                    {
                        graphics.DrawRectangle(pen, -5, -0.5f, 10, 6);
                        if (state == LayerRowState.On)
                        {
                            graphics.DrawArc(pen, -3.1f, -6.7f, 6.2f, 6.2f, 180, 180);
                            graphics.DrawLine(pen, -3.1f, -3.6f, -3.1f, -0.5f);
                            graphics.DrawLine(pen, 3.1f, -3.6f, 3.1f, -0.5f);
                        }
                        else
                        {
                            graphics.DrawArc(pen, 2, -6.7f, 6.2f, 6.2f, 180, 180);
                            graphics.DrawLine(pen, 2, -3.6f, 2, -0.5f);
                        }
                    }
                    if (state == LayerRowState.Mixed) graphics.DrawLine(pen, -2, 8, 2, 8);
                }
            }
            finally { graphics.Restore(saved); }
        }

        private int IconWidth { get { return tree.ItemHeight + 2; } }
        private int Column(int x)
        {
            int right = tree.ClientSize.Width;
            return x >= right - IconWidth ? 2 : x >= right - 2 * IconWidth ? 1 : 0;
        }

        private TreeNode NodeAt(int y)
        {
            for (TreeNode node = tree.TopNode; node != null; node = node.NextVisibleNode)
            {
                if (y < node.Bounds.Top) return null;
                if (y < node.Bounds.Top + tree.ItemHeight) return node;
            }
            return null;
        }

        private void ClearDrop() { dropTarget = null; tree.Invalidate(); }

        private static FlowLayoutPanel Row()
        { return new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true }; }

        private Button AddButton(FlowLayoutPanel row, string label, Action command)
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

        protected override void Dispose(bool disposing)
        {
            if (disposing) { tips.Dispose(); layerFont.Dispose(); }
            base.Dispose(disposing);
        }

        private sealed class LayerTree : TreeView
        {
            public Func<Point, Action> HitIcon;

            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x201 || message.Msg == 0x203)
                {
                    long packed = message.LParam.ToInt64();
                    var point = new Point(unchecked((short)packed), unchecked((short)(packed >> 16)));
                    var action = HitIcon?.Invoke(point);
                    if (action != null)
                    {
                        // The Office pane host can keep the mouse-up used to
                        // activate it. Dispatch the icon click after this
                        // mouse-down instead of capturing the pointer.
                        BeginInvoke(action);
                        return;
                    }
                }
                base.WndProc(ref message);
            }
        }
    }
}
