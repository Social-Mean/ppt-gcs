using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PptGcs.Core;
using Extensibility;
using Microsoft.Office.Core;

namespace PptGcs.AddIn
{
    [ComVisible(true), Guid("1B728F4B-5AF0-41FD-8589-86E44A7A5501"), ProgId("PptGcs.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : StandardOleMarshalObject, IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic app;
        private Timer timer;
        private ConstrainedDragController dragController;
        private bool busy;
        private string activeKey;
        private readonly Dictionary<string, double> lastTop = new Dictionary<string, double>();
        private ConstraintDocument document = new ConstraintDocument(new SameTopConstraint[0]);
        private const string TagName = "PPTGCS_ID";
        private const string XmlNamespace = "urn:ppt-gcs:v1";
        private DateTime nextRefresh;
        private string lastError;
        private readonly string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PptGcs", "addin.log");

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            app = application;
            ((dynamic)addInInst).Object = this;
            dragController = new ConstrainedDragController(application, Log);
            timer = new Timer { Interval = 25 };
            timer.Tick += Tick;
            timer.Start();
            Log("connected", "PowerPoint=" + app.Version);
        }
        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { Stop(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        private void Stop()
        {
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (dragController != null) { dragController.Dispose(); dragController = null; }
            Log("disconnected", ""); app = null;
        }

        public string GetCustomUI(string ribbonId)
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"PptGcsTab\" label=\"PPT GCS\"><group id=\"PptGcsGroup\" label=\"几何约束 · 实验版\"><button id=\"GcsLink\" label=\"顶部联动\" size=\"large\" imageMso=\"ObjectAlignTop\" onAction=\"LinkSelected\" screentip=\"选择两个未旋转矩形，建立双向顶部约束\"/><button id=\"GcsUnlink\" label=\"移除联动\" onAction=\"UnlinkSelected\"/><button id=\"GcsDemo\" label=\"新建演示\" onAction=\"CreateDemo\"/><button id=\"GcsStatus\" label=\"状态\" onAction=\"ShowStatus\"/></group></tab></tabs></ribbon></customUI>";
        }

        public void CreateDemo(object control)
        {
            RunCommand(() => {
                dynamic p = app.Presentations.Add();
                p.PageSetup.SlideWidth = 960f; p.PageSetup.SlideHeight = 540f;
                dynamic slide = p.Slides.Add(1, 12);
                dynamic a = slide.Shapes.AddShape(1, 180f, 180f, 200f, 110f);
                dynamic b = slide.Shapes.AddShape(1, 580f, 180f, 200f, 110f);
                a.Name = "GCS_Rectangle_A"; b.Name = "GCS_Rectangle_B";
                a.Fill.ForeColor.RGB = 0xD88928; b.Fill.ForeColor.RGB = 0x8B6542;
                a.TextFrame.TextRange.Text = "A"; b.TextFrame.TextRange.Text = "B";
                a.TextFrame.TextRange.Font.Size = 32f; b.TextFrame.TextRange.Font.Size = 32f;
                AddPair(p, slide, a, b);
                a.Select();
                Log("demo-created", "slide=" + slide.SlideID);
            });
        }
        public void LinkSelected(object control)
        {
            RunCommand(() => {
                dynamic selection = app.ActiveWindow.Selection;
                if ((int)selection.Type != 2 || (int)selection.ShapeRange.Count != 2) throw new InvalidOperationException("请先选中两个未旋转的普通矩形。");
                dynamic a = selection.ShapeRange[1]; dynamic b = selection.ShapeRange[2];
                ValidateRectangle(a); ValidateRectangle(b);
                app.StartNewUndoEntry();
                AddPair(app.ActivePresentation, app.ActiveWindow.View.Slide, a, b);
            });
        }
        private void ValidateRectangle(dynamic shape)
        {
            if ((int)shape.Type != 1 || (int)shape.AutoShapeType != 1 || Math.Abs((double)shape.Rotation) > 0.01)
                throw new InvalidOperationException("当前实验仅支持未旋转的普通矩形。");
        }
        private string GetOrCreateId(dynamic shape)
        {
            string id = (string)shape.Tags[TagName];
            Guid parsed;
            if (!Guid.TryParse(id, out parsed) || parsed == Guid.Empty) parsed = Guid.NewGuid();
            return parsed.ToString("D").ToUpperInvariant();
        }
        private void AddPair(dynamic presentation, dynamic slide, dynamic a, dynamic b)
        {
            ConstraintDocument current = ReadDocument(presentation);
            string aid = GetOrCreateId(a), bid = GetOrCreateId(b);
            if (aid == bid) bid = Guid.NewGuid().ToString("D").ToUpperInvariant();
            if (current.Constraints.Any(c => c.ShapeA == aid || c.ShapeB == aid || c.ShapeA == bid || c.ShapeB == bid))
                throw new InvalidOperationException("该矩形已有联动。实验版每个矩形只支持一组配对，请先移除原联动。");
            var list = current.Constraints.ToList();
            list.Add(new SameTopConstraint(Guid.NewGuid().ToString("D"), (int)slide.SlideID, aid, bid));
            a.Tags.Add(TagName, aid); b.Tags.Add(TagName, bid);
            WriteDocument(presentation, new ConstraintDocument(list));
            b.Top = a.Top;
            activeKey = null;
        }
        public void UnlinkSelected(object control)
        {
            RunCommand(() => {
                dynamic selection = app.ActiveWindow.Selection;
                if ((int)selection.Type != 2) throw new InvalidOperationException("请选中要移除约束的矩形。");
                var ids = new HashSet<string>();
                for (int i = 1; i <= (int)selection.ShapeRange.Count; i++) ids.Add(((string)selection.ShapeRange[i].Tags[TagName]).ToUpperInvariant());
                dynamic p = app.ActivePresentation;
                ConstraintDocument current = ReadDocument(p);
                app.StartNewUndoEntry();
                WriteDocument(p, new ConstraintDocument(current.Constraints.Where(c => !ids.Contains(c.ShapeA) && !ids.Contains(c.ShapeB))));
                activeKey = null;
            });
        }
        public void ShowStatus(object control) { MessageBox.Show(GetStatus(), "PPT GCS"); }
        public string GetStatus() { return "双向顶部联动实验：拖动矩形内部\r\n已载入约束：" + document.Constraints.Count + "\r\n状态：" + (lastError ?? "运行中") + "\r\n日志：" + logPath; }
        public string GetDiagnostics() { return dragController.InspectTargets(); }
        public void Reload() { activeKey = null; }

        private ConstraintDocument ReadDocument(dynamic presentation)
        {
            dynamic parts = presentation.CustomXMLParts.SelectByNamespace(XmlNamespace);
            if ((int)parts.Count == 0) return new ConstraintDocument(new SameTopConstraint[0]);
            if ((int)parts.Count != 1) throw new InvalidOperationException("发现重复的 PPT GCS 约束数据，已停止联动。");
            return ConstraintDocument.Parse((string)parts[1].XML);
        }
        private void WriteDocument(dynamic presentation, ConstraintDocument value)
        {
            dynamic parts = presentation.CustomXMLParts.SelectByNamespace(XmlNamespace);
            if ((int)parts.Count == 0) presentation.CustomXMLParts.Add(value.ToXml());
            else if ((int)parts.Count == 1) { if (!(bool)parts[1].LoadXML(value.ToXml())) throw new InvalidOperationException("约束数据写入失败。"); }
            else throw new InvalidOperationException("发现重复约束数据。");
        }
        private void Tick(object sender, EventArgs args)
        {
            if (busy || app == null) return;
            busy = true;
            try
            {
                dragController.Tick();
                if (dragController.IsDragging) return;
                if (activeKey != null && DateTime.UtcNow < nextRefresh) return;
                nextRefresh = DateTime.UtcNow.AddMilliseconds(200);
                dragController.SetTargets(new DragPair[0]);
                if ((int)app.Presentations.Count == 0) { activeKey = null; return; }
                dynamic window = app.ActiveWindow;
                if (window == null || (int)window.ViewType != 9) return;
                dynamic p = app.ActivePresentation; dynamic slide = window.View.Slide;
                int slideId = (int)slide.SlideID;
                string key = (string)p.FullName + "/" + slideId;
                ConstraintDocument loaded = ReadDocument(p);
                if (key != activeKey || loaded.ToXml() != document.ToXml())
                {
                    document = loaded; activeKey = key; lastTop.Clear();
                    Log("document-loaded", "constraints=" + document.Constraints.Count + " key=" + key);
                }
                var pairs = new List<DragPair>();
                foreach (var constraint in document.Constraints.Where(c => c.SlideId == slideId))
                {
                    dynamic a = null, b = null; int ac = 0, bc = 0;
                    for (int i = 1; i <= (int)slide.Shapes.Count; i++)
                    {
                        dynamic shape = slide.Shapes[i];
                        string id = ((string)shape.Tags[TagName]).ToUpperInvariant();
                        if (id == constraint.ShapeA) { a = shape; ac++; }
                        if (id == constraint.ShapeB) { b = shape; bc++; }
                    }
                    if (ac != 1 || bc != 1) continue;
                    double at = (double)a.Top, bt = (double)b.Top;
                    pairs.Add(new DragPair((object)a, (object)b, constraint));
                    double oldA, oldB;
                    if (lastTop.TryGetValue(constraint.ShapeA, out oldA) && lastTop.TryGetValue(constraint.ShapeB, out oldB))
                    {
                        bool ca = Math.Abs(at-oldA)>0.01, cb = Math.Abs(bt-oldB)>0.01;
                        if (ca || cb)
                        {
                            // Observe native Undo/Redo without writing back or replaying the move.
                            Log("observed", string.Format(CultureInfo.InvariantCulture,"a={0:F3} b={1:F3}",at,bt));
                        }
                    }
                    lastTop[constraint.ShapeA]=at; lastTop[constraint.ShapeB]=bt;
                }
                dragController.SetTargets(pairs);
                lastError = null;
            }
            catch (COMException) { /* PowerPoint can reject reads while editing a dialog. */ }
            catch (Exception ex) { if (lastError != ex.Message) Log("tick-error",ex.ToString()); lastError = ex.Message; }
            finally { busy=false; }
        }
        private void RunCommand(Action action)
        {
            if (busy || (dragController != null && dragController.IsDragging)) return; busy=true;
            try { action(); } catch(Exception ex) { Log("command-error",ex.ToString()); MessageBox.Show(ex.Message,"PPT GCS"); }
            finally { busy=false; }
        }
        private void Log(string type,string message)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(logPath)); File.AppendAllText(logPath,DateTime.UtcNow.ToString("O")+" "+type+" "+message+Environment.NewLine); } catch(IOException) { }
        }
    }
}
