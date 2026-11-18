using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class PaneTests
{
    private static int passed;

    [STAThread]
    private static int Main()
    {
        try
        {
            using (var pane = new ConstraintPaneControl()) Run(pane);
            using (var pane = new ConstraintPaneControl()) RunGrid(pane);
            Console.WriteLine("PASS: " + passed + " pane tests");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static void RunGrid(ConstraintPaneControl pane)
    {
        pane.CreateControl();
        var model=Model("网格测试","shape:256/");
        model.Options=ConstraintCatalog.Options(new List<ShapeGeometry>());
        model.Objects[1].Selected=true;
        string created=null;
        pane.Bind(new ConstraintPaneActions {Add=(kind,value,second)=>created=kind});
        pane.Render(model);
        pane.ShowEditor("SameWidth");
        var grid=Find<TableLayoutPanel>(pane,"约束类型按钮");
        var viewport=Find<Panel>(pane,"约束类型");
        var apply=Find<Button>(pane,"创建约束");
        var buttons=grid.Controls.OfType<Button>().ToList();
        Check(buttons.Count==20 && grid.ColumnCount==4 && grid.RowCount==5,
            "twenty constraint buttons use four columns and five rows");
        Check(model.Options.Select((option,index)=>new {option,index}).All(item=>{
            var button=buttons.Single(control=>(string)control.Tag==item.option.Kind);
            var cell=grid.GetPositionFromControl(button);
            return cell.Column==item.index%4 && cell.Row==item.index/4 && button.TabIndex==item.index;
        }),"catalog order fills each row from left to right and defines keyboard tab order");
        Check(!Descendants(pane).OfType<ComboBox>().Any(control=>control.AccessibleName=="约束类型"),
            "constraint type no longer uses a dropdown");
        var width=Find<Button>(pane,"约束类型：同宽");
        var height=Find<Button>(pane,"约束类型：同高");
        Check(width.FlatAppearance.BorderSize==2 && width.BackColor==(SystemInformation.HighContrast?SystemColors.Highlight:Color.FromArgb(220,235,252)),
            "the active constraint type has a visible selected state");
        height.PerformClick();
        Check(height.FlatAppearance.BorderSize==2 && width.FlatAppearance.BorderSize==1,
            "clicking a constraint button switches the single selected type");
        apply.PerformClick();
        Check(created=="SameHeight","button choice invokes the shared create command");

        pane.ShowEditor("FixedWidth");
        Check(!Find<Button>(pane,"约束类型：固定宽度").Enabled && !apply.Enabled,
            "buttons and create action reflect maximum selection size");
        model.Objects[1].Selected=false;
        pane.Render(model);
        var number=Find<NumericUpDown>(pane,"第一个约束参数");
        number.Value=137;
        for(int i=0;i<4;i++)pane.Render(model);
        Check(number.Value==137 && Find<Button>(pane,"约束类型：固定宽度").Enabled,
            "selection refresh updates button eligibility without replacing entered parameters");

        model.Objects[0].UnsupportedReason="旋转对象暂不支持";
        pane.Render(model);
        Check(grid.Controls.OfType<Button>().All(button=>!button.Enabled) && !apply.Enabled,
            "unsupported selected objects disable all creation choices");
        model.Objects[0].UnsupportedReason=null;
        pane.Render(model);
        var relations=Find<ListView>(pane,"对象之间的约束");
        relations.Items[1].Selected=true;
        Find<Button>(pane,"修改…").PerformClick();
        Check(grid.Controls.OfType<Button>().All(button=>!button.Enabled) && apply.Enabled,
            "editing an existing relation locks its type while allowing parameter changes");

        pane.Width=250;pane.PerformLayout();
        var editor=Find<GroupBox>(pane,"约束参数编辑");
        var fields=editor.Controls.OfType<TableLayoutPanel>().Single();
        fields.PerformLayout();editor.PerformLayout();pane.PerformLayout();
        Check(fields.Bottom+editor.Padding.Bottom<=editor.Height,
            "wrapped editor hints keep the parameter and apply controls inside the group");
        Check(viewport.AutoScroll && viewport.AutoScrollMinSize.Width==grid.Width && grid.Width>viewport.ClientSize.Width,
            "narrow panes preserve readable button columns with horizontal scrolling");
        Check(grid.Controls.OfType<Button>().All(button=>!String.IsNullOrEmpty(button.AccessibleName) && button.AccessibleName.StartsWith("约束类型：")),
            "all type buttons expose meaningful accessible names");
    }

    private static void Run(ConstraintPaneControl pane)
    {
        // A UserControl can create its child handles without a visible Form.
        // No PowerPoint instance, global input, screenshot or desktop file is used.
        pane.CreateControl();
        var objects = Find<ListView>(pane, "当前页对象");
        var relations = Find<ListView>(pane, "对象之间的约束");
        var number = Find<NumericUpDown>(pane, "第一个约束参数");
        var message = Find<Label>(pane, "约束操作状态");
        var editor = Find<GroupBox>(pane, "约束参数编辑");
        var apply = Find<Button>(pane, "创建约束");
        var edit = Find<Button>(pane, "修改…");
        var enable = Find<Button>(pane, "停用");
        var delete = Find<Button>(pane, "删除");
        var objectHandle = objects.Handle;
        var relationHandle = relations.Handle;

        var model = Model("第 1 页", "shape:256/");
        string driver="AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA";
        var geometry=new[]{new ShapeGeometry(driver,20,30,72,30)};
        var lockedWidth=new[]{new LayoutConstraint("BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB",256,ConstraintKind.FixedWidth,new[]{driver},72)};
        model.Objects[0].Freedom=LayoutDiagnostics.AnalyzeShape(geometry,lockedWidth,driver);
        string[] selection = null;
        int selectionCalls = 0;
        string addKind = null, updatedId = null, deletedId = null;
        double? addValue = null, updatedValue = null;
        bool? enabled = null;
        pane.Bind(new ConstraintPaneActions {
            SelectShapes = ids => { selection = ids.ToArray(); selectionCalls++; },
            Add = (kind, value, second) => { addKind = kind; addValue = value; },
            Update = (id, value, second) => { updatedId = id; updatedValue = value; },
            Delete = id => deletedId = id,
            SetEnabled = (id, value) => enabled = value
        });

        pane.Render(model);
        Check(objects.Items.Count == 3 && relations.Items.Count == 2,
            "objects and multi-member relations render");
        Check(selectionCalls == 0, "render does not trigger document selection callbacks");
        var freedom=Find<Label>(pane,"所选对象可编辑方向");
        Check(freedom.Visible && freedom.Text.Contains("横向可移动") && freedom.Text.Contains("宽度固定") && !freedom.Text.Contains("矩形 B"),
            "pane reports editable directions and fixed dimensions for the selected object");

        pane.ShowEditor("FixedWidth");
        number.Value = 72;
        pane.Render(model);
        Check(number.Value == 72, "refresh preserves pending parameter input");
        apply.PerformClick();
        Check(addKind == "FixedWidth" && addValue == 72,
            "create invokes shared command with the entered numeric parameter");

        // A native ListView sends deselection and selection separately. Only the
        // final choice may be forwarded; the empty intermediate selection would
        // otherwise refresh and rebuild the list while a click is in progress.
        objects.Items[0].Selected = false;
        objects.Items[1].Selected = true;
        Application.DoEvents();
        Check(selectionCalls == 1 && selection.SequenceEqual(new[] { "shape:256/5" }),
            "object selection notifications coalesce to the final native selection");

        relations.Items[1].Selected = true;
        Application.DoEvents();
        Check(selection.SequenceEqual(new[] { "shape:256/4" }),
            "relation selection locates only its members");
        Check(enable.Text == "启用", "disabled relation offers enable action");
        Check(enable.AccessibleName == enable.Text, "accessible action matches the current enable command");
        enable.PerformClick();
        Check(enabled == true, "enable forwards the requested relation state");
        edit.PerformClick();
        number.Value = 96;
        apply.PerformClick();
        Check(updatedId == "R2" && updatedValue == 96,
            "parameter edit targets the selected relation with its new value");
        delete.PerformClick();
        Check(deletedId == "R2", "delete targets only the selected relation");

        int callsBeforeCanvasActivation = selectionCalls;
        pane.ActivateRelation("R1", true);
        Check(relations.SelectedItems.Count == 1 && relations.SelectedItems[0].Text.StartsWith("同宽") && !editor.Visible,
            "activating a parameterless relation selects its row without opening an editor");
        pane.ActivateRelation("R2", true);
        Check(pane.EditorKind == "FixedWidth" && number.Value == 60 && selectionCalls == callsBeforeCanvasActivation,
            "canvas activation loads the addressed relation without reselecting objects");
        number.Value = 88;
        apply.PerformClick();
        Check(updatedId == "R2" && updatedValue == 88, "canvas editor applies the named relation through the shared action");

        pane.ShowEditor("SameWidth");
        Check(!apply.Enabled, "insufficient selected objects disable creation");

        model.Message = "已选 1 个对象 · 当前页 2 条关系";
        pane.Bind(new ConstraintPaneActions {
            Add = (kind, value, second) => {
                // The host refreshes from its finally block before propagating
                // the command error. A later idle refresh must not erase it.
                pane.Render(model);
                throw new InvalidOperationException("约束冲突：固定宽度不一致");
            }
        });
        pane.ShowEditor("FixedWidth");
        apply.PerformClick();
        for (int index = 0; index < 5; index++) pane.Render(model);
        Check(message.Text == "约束冲突：固定宽度不一致" && editor.Visible,
            "background renders preserve command failure feedback and pending editor");
        pane.ShowEditor("FixedWidth");
        Check(message.Text == model.Message,
            "the next deliberate editor operation clears stale failure feedback");

        pane.Bind(new ConstraintPaneActions { SelectShapes = ids => selectionCalls++ });
        objects.Items[0].Selected = false;
        objects.Items[2].Selected = true;
        int callsBeforePageChange = selectionCalls;
        pane.Render(Model("第 2 页", "shape:257/"));
        Application.DoEvents();
        Check(selectionCalls == callsBeforePageChange,
            "page switch cancels queued selection from the previous page");
        Check(!editor.Visible && relations.SelectedItems.Count == 0,
            "page switch closes the parameter editor and clears relation selection");
    }

    private static ConstraintPaneModel Model(string page, string prefix)
    {
        return new ConstraintPaneModel {
            PageLabel = page,
            Objects = new List<ConstraintPaneObject> {
                new ConstraintPaneObject { Id=prefix+"4", Name="矩形 A", Selected=true, RelationCount=2 },
                new ConstraintPaneObject { Id=prefix+"5", Name="矩形 B", RelationCount=1 },
                new ConstraintPaneObject { Id=prefix+"6", Name="矩形 C", RelationCount=0 }
            },
            Relations = new List<ConstraintPaneRelation> {
                new ConstraintPaneRelation { Id="R1", Kind="SameWidth", Label="同宽",
                    MemberIds=new[] {prefix+"4",prefix+"5"}, MemberNames=new[] {"矩形 A","矩形 B"}, Status="已满足" },
                new ConstraintPaneRelation { Id="R2", Kind="FixedWidth", Label="固定宽度",
                    MemberIds=new[] {prefix+"4"}, MemberNames=new[] {"矩形 A"}, Value=60, Enabled=false }
            },
            Options = new List<ConstraintPaneOption> {
                new ConstraintPaneOption { Kind="SameWidth", Label="同宽", MinimumMembers=2 },
                new ConstraintPaneOption { Kind="FixedWidth", Label="固定宽度", MinimumMembers=1, MaximumMembers=1,
                    ValueLabel="宽度 (pt)", MinimumValue=1, DefaultValue=60, Description="宽度保持为指定值" }
            }
        };
    }

    private static T Find<T>(Control parent, string accessibleName) where T : Control
    {
        var matches = Descendants(parent).OfType<T>().Where(item => item.AccessibleName == accessibleName).ToList();
        if (matches.Count != 1)
            throw new InvalidOperationException("Expected one accessible control: " + accessibleName);
        return matches[0];
    }

    private static IEnumerable<Control> Descendants(Control parent)
    {
        foreach (Control child in parent.Controls)
        {
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }

    private static void Check(bool condition, string description)
    {
        if (!condition) throw new InvalidOperationException(description);
        passed++;
        Console.WriteLine("PASS " + description);
    }
}
