using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PptGcs.AddIn;

internal static class PaneTests
{
    private static int passed;

    [STAThread]
    private static int Main()
    {
        try
        {
            using (var pane = new ConstraintPaneControl()) Run(pane);
            Console.WriteLine("PASS: " + passed + " pane tests");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
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
