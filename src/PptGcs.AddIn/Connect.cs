using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    [ComVisible(true), Guid("1B728F4B-5AF0-41FD-8589-86E44A7A5501"), ProgId("PptGcs.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : StandardOleMarshalObject, IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic app;
        private Timer timer;
        private ConstrainedDragController controller;
        private bool busy;
        private DateTime nextRefresh;
        private string activeKey;
        private string lastError;
        private int constraintCount;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PptGcs", "layout-v2.log");

        public void OnConnection(object application, ext_ConnectMode mode, object addIn, ref Array custom)
        {
            app = application;
            ((dynamic)addIn).Object = this;
            controller = new ConstrainedDragController(application, Log);
            timer = new Timer { Interval = 25 };
            timer.Tick += Tick;
            timer.Start();
            Log("connected", json.Serialize(new { version = "0.2.0", powerpoint = (string)app.Version, processId = Process.GetCurrentProcess().Id }));
        }
        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { Stop(); }
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        private void Stop()
        {
            if (timer == null && controller == null) return;
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (controller != null) { controller.Dispose(); controller = null; }
            Log("disconnected", "{}"); app = null;
        }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui""><ribbon><tabs>
<tab id=""PptGcsTab"" label=""PPT GCS"">
<group id=""Alignment"" label=""位置约束"">
<button id=""SameX"" label=""同 X（左边对齐）"" onAction=""LinkSameX""/>
<button id=""SameY"" label=""同 Y（顶部对齐）"" onAction=""LinkSameY""/>
</group><group id=""Dimensions"" label=""尺寸约束"">
<button id=""SameWidth"" label=""同宽"" onAction=""LinkSameWidth""/>
<button id=""SameHeight"" label=""同高"" onAction=""LinkSameHeight""/>
</group><group id=""Spacing"" label=""等间距（至少三个）"">
<button id=""EqualHorizontalGap"" label=""水平等间距"" onAction=""LinkHorizontalGap""/>
<button id=""EqualVerticalGap"" label=""垂直等间距"" onAction=""LinkVerticalGap""/>
</group><group id=""Manage"" label=""管理"">
<button id=""Remove"" label=""移除选中对象的约束"" onAction=""UnlinkSelected""/>
<button id=""Demo"" label=""新建六组演示"" onAction=""CreateDemo""/>
<button id=""Status"" label=""状态"" onAction=""ShowStatus""/>
</group></tab></tabs></ribbon></customUI>";
        }
        public void LinkSameX(object control) { LinkSelection(ConstraintKind.SameX); }
        public void LinkSameY(object control) { LinkSelection(ConstraintKind.SameY); }
        public void LinkSameWidth(object control) { LinkSelection(ConstraintKind.SameWidth); }
        public void LinkSameHeight(object control) { LinkSelection(ConstraintKind.SameHeight); }
        public void LinkHorizontalGap(object control) { LinkSelection(ConstraintKind.EqualHorizontalGap); }
        public void LinkVerticalGap(object control) { LinkSelection(ConstraintKind.EqualVerticalGap); }

        private void LinkSelection(ConstraintKind kind)
        {
            RunCommand(() => {
                LayoutContext context = PowerPointLayout.Read((object)app);
                dynamic selection = ((dynamic)context.Window).Selection;
                int minimum = IsGap(kind) ? 3 : 2;
                if ((int)selection.Type != 2 || (int)selection.ShapeRange.Count < minimum)
                    throw new InvalidOperationException("请在同一页选中至少 " + minimum + " 个矩形。");
                var shapes = new List<object>();
                for (int i = 1; i <= (int)selection.ShapeRange.Count; i++) shapes.Add((object)selection.ShapeRange[i]);
                AddConstraint(context, shapes, kind);
            });
        }

        private static bool IsGap(ConstraintKind kind)
        {
            return kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.EqualVerticalGap;
        }

        private void AddConstraint(LayoutContext context, List<object> selected, ConstraintKind kind)
        {
            foreach (var shape in selected) PowerPointLayout.ValidateShape(shape);
            if (IsGap(kind))
                selected = selected.OrderBy(s => kind == ConstraintKind.EqualHorizontalGap ? (double)((dynamic)s).Left : (double)((dynamic)s).Top).ToList();
            var ids = new List<string>();
            var savedTags = new Dictionary<object, string>();
            foreach (var shape in selected)
            {
                string existing = (string)((dynamic)shape).Tags[PowerPointLayout.TagName];
                string id = PowerPointLayout.ReadId(shape) ?? Guid.NewGuid().ToString("D").ToUpperInvariant();
                dynamic slide = context.Slide;
                for (int i=1; i <= (int)slide.Shapes.Count; i++)
                {
                    dynamic other = slide.Shapes[i];
                    if ((int)other.Id != (int)((dynamic)shape).Id && PowerPointLayout.ReadId((object)other) == id)
                        throw new InvalidOperationException("此页存在重复的矩形标识，请重新创建复制的矩形后再建立约束。");
                }
                if (ids.Contains(id) || (context.Shapes.ContainsKey(id) && (int)((dynamic)context.Shapes[id]).Id != (int)((dynamic)shape).Id))
                    throw new InvalidOperationException("矩形 GUID 重复，请先移除复制对象的标识或重新创建矩形。");
                ids.Add(id); savedTags.Add(shape, existing);
                if (!context.Shapes.ContainsKey(id))
                {
                    context.Shapes.Add(id, shape);
                    context.Geometry.Add(PowerPointLayout.Geometry(shape, id));
                }
            }
            ConstraintDocument original = context.Document;
            var all = original.Constraints.ToList();
            var relation = new LayoutConstraint(Guid.NewGuid().ToString("D"), context.SlideId, kind, ids);
            all.Add(relation);
            var proposed = new ConstraintDocument(all);
            var relations = proposed.Constraints.Where(c => c.SlideId == context.SlideId).ToList();
            var target = context.Geometry.Single(g => g.Id == ids[0]);
            // Solve before changing the document so incompatible additions are atomic.
            var solved = LayoutSolver.Solve(context.Geometry, relations, target.Id, target);
            app.StartNewUndoEntry();
            try
            {
                for (int i = 0; i < selected.Count; i++) ((dynamic)selected[i]).Tags.Add(PowerPointLayout.TagName, ids[i]);
                PowerPointLayout.WriteDocument(context.Presentation, proposed);
                PowerPointLayout.Apply(context, solved);
                var readback = PowerPointLayout.Snapshot(context);
                if (LayoutSolver.MaxResidual(readback, relations) > LayoutSolver.PositionTolerance)
                    throw new InvalidOperationException("PowerPoint 的自动尺寸设置阻止了约束写回。");
                activeKey = null;
                Log("constraint-added", json.Serialize(new { id = relation.Id, kind = kind.ToString(), slideId = context.SlideId, shapeIds = ids }));
            }
            catch
            {
                PowerPointLayout.Apply(context, context.Geometry);
                PowerPointLayout.WriteDocument(context.Presentation, original);
                foreach (var entry in savedTags)
                {
                    if (String.IsNullOrEmpty(entry.Value)) ((dynamic)entry.Key).Tags.Delete(PowerPointLayout.TagName);
                    else ((dynamic)entry.Key).Tags.Add(PowerPointLayout.TagName, entry.Value);
                }
                throw;
            }
            finally { app.StartNewUndoEntry(); }
        }

        public void UnlinkSelected(object control)
        {
            RunCommand(() => {
                dynamic window = app.ActiveWindow;
                dynamic selection = window.Selection;
                if ((int)selection.Type != 2) throw new InvalidOperationException("请先选中矩形。");
                var ids = new HashSet<string>();
                for (int i = 1; i <= (int)selection.ShapeRange.Count; i++)
                {
                    string id = PowerPointLayout.ReadId((object)selection.ShapeRange[i]);
                    if (id != null) ids.Add(id);
                }
                int slideId = (int)window.View.Slide.SlideID;
                ConstraintDocument current = PowerPointLayout.ReadDocument((object)window.Presentation);
                var remaining = current.Constraints.Where(c => c.SlideId != slideId || !c.ShapeIds.Any(ids.Contains)).ToList();
                app.StartNewUndoEntry();
                PowerPointLayout.WriteDocument((object)window.Presentation, new ConstraintDocument(remaining));
                app.StartNewUndoEntry();
                activeKey = null;
            });
        }

        public void CreateDemo(object control)
        {
            RunCommand(() => {
                dynamic presentation = app.Presentations.Add();
                presentation.PageSetup.SlideWidth = 960f; presentation.PageSetup.SlideHeight = 540f;
                var kinds = new[] { ConstraintKind.SameY, ConstraintKind.SameX, ConstraintKind.SameWidth,
                    ConstraintKind.SameHeight, ConstraintKind.EqualHorizontalGap, ConstraintKind.EqualVerticalGap };
                string[] titles = { "同 Y · 顶部对齐", "同 X · 左边对齐", "同宽 · 拉伸联动", "同高 · 拉伸联动", "水平等间距 · 按边缘距离", "垂直等间距 · 按边缘距离" };
                string[] hints = { "从内部上下拖动 A，再上下拖动 B。", "从内部左右拖动 A，再左右拖动 B。",
                    "先选中 A，拖动右边中点缩放；再对 B 做同样操作。", "先选中 A，拖动底边中点缩放；再对 B 做同样操作。",
                    "从内部左右拖动 A，再左右拖动 C。", "从内部上下拖动 A，再上下拖动 C。" };
                for (int index = 0; index < kinds.Length; index++)
                {
                    ConstraintKind kind = kinds[index];
                    dynamic slide = presentation.Slides.Add(index+1, 12);
                    slide.Name = kind.ToString(); slide.Tags.Add("PPTGCS_CASE", kind.ToString());
                    AddText(slide, (index+1) + "  " + titles[index], 28, 18, 900, 38, 24);
                    AddText(slide, hints[index], 28, 57, 900, 30, 15);
                    var shapes = new List<object>();
                    if (kind == ConstraintKind.SameY)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 180,190,170,100));
                        shapes.Add(Rectangle(slide, kind, "B", 580,260,170,100));
                    }
                    else if (kind == ConstraintKind.SameX)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 200,150,150,90));
                        shapes.Add(Rectangle(slide, kind, "B", 600,350,170,90));
                    }
                    else if (kind == ConstraintKind.SameWidth)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 180,200,160,100));
                        shapes.Add(Rectangle(slide, kind, "B", 580,200,220,140));
                    }
                    else if (kind == ConstraintKind.SameHeight)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 180,200,180,100));
                        shapes.Add(Rectangle(slide, kind, "B", 580,200,220,160));
                    }
                    else if (kind == ConstraintKind.EqualHorizontalGap)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 120,220,120,90));
                        shapes.Add(Rectangle(slide, kind, "B", 380,220,160,90));
                        shapes.Add(Rectangle(slide, kind, "C", 700,220,100,90));
                    }
                    else
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 420,100,140,60));
                        shapes.Add(Rectangle(slide, kind, "B", 420,245,180,80));
                        shapes.Add(Rectangle(slide, kind, "C", 420,420,120,55));
                    }
                    app.ActiveWindow.View.GotoSlide(index+1);
                    AddConstraint(PowerPointLayout.Read((object)app), shapes, kind);
                }
                app.ActiveWindow.View.GotoSlide(1);
                activeKey = null;
                Log("demo-created", json.Serialize(new { slides = 6, processId = Process.GetCurrentProcess().Id }));
            });
        }

        private static void AddText(dynamic slide, string text, float left, float top, float width, float height, float size)
        {
            dynamic label = slide.Shapes.AddTextbox(1, left, top, width, height);
            label.TextFrame.TextRange.Text = text; label.TextFrame.TextRange.Font.Size = size;
            label.TextFrame.TextRange.Font.Color.RGB = 0x403020;
        }
        private static object Rectangle(dynamic slide, ConstraintKind kind, string letter, float left, float top, float width, float height)
        {
            dynamic shape = slide.Shapes.AddShape(1,left,top,width,height);
            shape.Name = "GCS_" + kind + "_" + letter;
            shape.LockAspectRatio = 0; shape.TextFrame.AutoSize = 0;
            shape.Fill.ForeColor.RGB = letter == "B" ? 0x8B6542 : 0xD88928;
            shape.TextFrame.TextRange.Text = letter;
            shape.TextFrame.TextRange.Font.Size = 26f;
            return shape;
        }

        public string GetSnapshot()
        {
            dynamic presentation = app.ActivePresentation;
            ConstraintDocument document = PowerPointLayout.ReadDocument((object)presentation);
            var slides = new List<object>();
            for (int i=1; i <= (int)presentation.Slides.Count; i++)
            {
                dynamic slide = presentation.Slides[i];
                int slideId = (int)slide.SlideID;
                var geometry = new List<object>();
                for (int j=1; j <= (int)slide.Shapes.Count; j++)
                {
                    dynamic shape = slide.Shapes[j];
                    string id = PowerPointLayout.ReadId((object)shape);
                    if (id == null) continue;
                    geometry.Add(new { Id=id, Name=(string)shape.Name, Left=(double)shape.Left, Top=(double)shape.Top, Width=(double)shape.Width, Height=(double)shape.Height });
                }
                var constraints = document.Constraints.Where(c => c.SlideId == slideId).Select(c => new { c.Id,c.SlideId,Kind=c.Kind.ToString(),c.ShapeIds }).ToList();
                slides.Add(new { slideId, name=(string)slide.Name, kind=(string)slide.Tags["PPTGCS_CASE"], geometry, constraints });
            }
            return json.Serialize(new { presentation=(string)presentation.FullName, slides });
        }
        public void ShowStatus(object control) { MessageBox.Show(GetStatus(),"PPT GCS"); }
        public string GetStatus() { return "几何约束 v0.2.0\r\n已载入约束：" + constraintCount + "\r\n状态：" + (lastError ?? "运行中") + "\r\n日志：" + logPath; }
        public string GetDiagnostics() { return controller.InspectTargets(); }
        public void Reload() { activeKey = null; }

        // Integration harness only: this is programmatic geometry replay, never
        // evidence of real mouse input. Restrict it to the six-case test fixture.
        public string ReplayGestureForTest(string driverId, string mode, string offsetsJson)
        {
            if (busy || controller.IsDragging) throw new InvalidOperationException("The add-in is busy.");
            dynamic presentation = app.ActivePresentation;
            string[] cases = { "SameY", "SameX", "SameWidth", "SameHeight", "EqualHorizontalGap", "EqualVerticalGap" };
            if ((int)presentation.Slides.Count != cases.Length)
                throw new InvalidOperationException("Replay requires the six-slide test fixture.");
            for (int i=1; i <= cases.Length; i++)
                if ((string)presentation.Slides[i].Tags["PPTGCS_CASE"] != cases[i-1])
                    throw new InvalidOperationException("Replay requires the six-slide test fixture.");
            double[][] offsets = json.Deserialize<double[][]>(offsetsJson);
            busy = true;
            try
            {
                controller.ReplayGestureForTest(driverId, mode, offsets);
                return GetSnapshot();
            }
            finally { busy = false; }
        }

        private void Tick(object sender, EventArgs args)
        {
            if (busy || app == null) return;
            busy = true;
            try
            {
                controller.Tick();
                if (controller.IsDragging || (activeKey != null && DateTime.UtcNow < nextRefresh)) return;
                nextRefresh = DateTime.UtcNow.AddMilliseconds(200);
                if ((int)app.Presentations.Count == 0) { activeKey=null;constraintCount=0;return; }
                dynamic window = app.ActiveWindow;
                if (window == null || (int)window.ViewType != 9) return;
                LayoutContext context = PowerPointLayout.Read((object)app);
                string key = (string)((dynamic)context.Presentation).FullName + "/" + context.SlideId + "/" + context.Document.ToXml();
                constraintCount = context.Document.Constraints.Count;
                if (key != activeKey)
                {
                    activeKey=key;
                    Log("document-loaded",json.Serialize(new { presentation=(string)((dynamic)context.Presentation).FullName,slideId=context.SlideId,constraints=constraintCount }));
                }
                lastError=null;
            }
            catch (COMException) { }
            catch (Exception ex) { if (lastError!=ex.Message) Log("host-error",json.Serialize(new { error=ex.Message }));lastError=ex.Message; }
            finally { busy=false; }
        }
        private void RunCommand(Action action)
        {
            if (busy || (controller != null && controller.IsDragging)) return;
            busy=true;
            try { action(); }
            catch (Exception ex) { Log("command-error",json.Serialize(new { error=ex.ToString() }));MessageBox.Show(ex.Message,"PPT GCS"); }
            finally { busy=false; }
        }
        private void Log(string type,string data)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(logPath));File.AppendAllText(logPath,DateTime.UtcNow.ToString("O")+" "+type+" "+data+Environment.NewLine); }
            catch(IOException) { }
        }
    }
}
