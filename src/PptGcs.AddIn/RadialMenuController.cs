using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Receives PowerPoint's own thread-hook messages. This class installs no
    // hooks and preview never writes a PowerPoint shape or constraint.
    public sealed class RadialMenuController : IDisposable
    {
        private const int MouseMove = 0x0200, MiddleDown = 0x0207, MiddleUp = 0x0208;
        private const int MouseWheel = 0x020A, HorizontalWheel = 0x020E, MiddleKey = 0x04;
        private readonly Func<ConstraintPaneModel> getModel;
        private readonly Func<string, int, Image> getIcon;
        private readonly Func<string, IList<ShapeGeometry>> preview;
        private readonly Action<string> execute;
        private readonly Action<bool> showPane;
        private readonly Func<object> activeWindow;
        private readonly Action normalize;
        private readonly Action<string, string> log;
        private readonly uint uiThread;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly Slot[] slots = {
            new Slot("SameX", "左边对齐", 2), new Slot("SameY", "顶部对齐", 2),
            new Slot("SameWidth", "同宽", 2), new Slot("SameHeight", "同高", 2),
            new Slot("EqualHorizontalGap", "水平等距", 3), new Slot("EqualVerticalGap", "垂直等距", 3),
            new Slot("more", "更多约束", 0), new Slot("view", "查看约束", 0)
        };
        private bool open, opening, processing, disposed, swallowMiddleUp, acceptedStart;
        private NativePoint? pendingMove, pendingRelease;
        private string pendingCancellation;
        private IntPtr canvasWindow, rootWindow, documentIdentity;
        private object documentWindow;
        private Rectangle canvasBounds;
        private PointF center;
        private NativePoint startPoint;
        private double originX, originY, scaleX, scaleY;
        private float dpiScale = 1, outerRadius, innerRadius;
        private int currentSlot = -1, selectedCount;
        private string selectionSignature, documentKey, gestureId, lastChoice, lastOutcome, lastError;
        private DateTime nextContextCheck, releasedSince;
        private RadialHud hud;
        private readonly List<RectangleF> previewBounds = new List<RectangleF>();

        public RadialMenuController(Func<ConstraintPaneModel> getModel,
            Func<string, IList<ShapeGeometry>> preview, Action<string> execute,
            Action<bool> showPane, Func<object> activeWindow, Action normalize, Action<string, string> log,
            Func<string, int, Image> getIcon = null)
        {
            this.getIcon = getIcon;
            this.getModel = getModel ?? throw new ArgumentNullException("getModel");
            this.preview = preview ?? throw new ArgumentNullException("preview");
            this.execute = execute ?? throw new ArgumentNullException("execute");
            this.showPane = showPane ?? throw new ArgumentNullException("showPane");
            this.activeWindow = activeWindow ?? throw new ArgumentNullException("activeWindow");
            this.normalize = normalize ?? throw new ArgumentNullException("normalize");
            this.log = log ?? throw new ArgumentNullException("log");
            uiThread = GetCurrentThreadId();
        }

        public bool IsOpen { get { return open || opening; } }

        public bool HandleMouse(int message, int screenX, int screenY, IntPtr window)
        {
            RequireUiThread();
            if (disposed) return false;
            var point = new NativePoint { X = screenX, Y = screenY };
            if (processing && IsOpen) return QueuePointer(message, point);
            if (!IsOpen && swallowMiddleUp && message == MiddleUp)
            { swallowMiddleUp = false; return true; }
            if (!IsOpen && message != MiddleDown) return false;
            bool owned = IsOpen;
            acceptedStart = false;
            processing = true;
            try
            {
                if (!IsOpen)
                {
                    swallowMiddleUp = false;
                    if (!Begin(point, window)) return false;
                    owned = true;
                }
                else if (!QueuePointer(message, point)) return false;
                DrainPointer();
                return true;
            }
            catch (Exception exception)
            {
                lastError = exception.Message;
                SafeLog("radial-error", new { gestureId, source = "middle-pointer", error = exception.ToString() });
                Close("error");
                return owned || acceptedStart;
            }
            finally { processing = false; }
        }

        public bool HandleKeyboard(int key, bool isDown)
        {
            RequireUiThread();
            if (disposed || !IsOpen) return false;
            if (key == 0x1B && isDown)
            {
                pendingCancellation = "escape";
                if (!processing) Close("escape");
            }
            return true;
        }

        private bool QueuePointer(int message, NativePoint point)
        {
            if (message == MouseWheel || message == HorizontalWheel)
            {
                pendingCancellation = "wheel-scroll";
                if (!opening) Close("wheel-scroll");
                return false;
            }
            if (message == MouseMove)
            {
                if (!pendingRelease.HasValue) pendingMove = point;
                return true;
            }
            if (message == MiddleUp)
            {
                if (!pendingRelease.HasValue) pendingRelease = point;
                swallowMiddleUp = false;
                return true;
            }
            // Left/right clicks must not change the selection behind the wheel.
            return message >= 0x0200 && message <= 0x020D;
        }

        private void DrainPointer()
        {
            while (IsOpen)
            {
                if (pendingCancellation != null) { Close(pendingCancellation); return; }
                if (pendingRelease.HasValue)
                {
                    var release = pendingRelease.Value;
                    pendingMove = null; pendingRelease = null;
                    Choose(release);
                    if (!IsOpen) return;
                    string choice = currentSlot >= 0 && slots[currentSlot].Enabled ? slots[currentSlot].Key : null;
                    lastChoice = choice;
                    SafeLog("radial-release", new { gestureId, source = "middle-pointer", x = release.X, y = release.Y,
                        choice, middleDown = KeyDown(MiddleKey), previewCount = previewBounds.Count });
                    IntPtr owner = rootWindow;
                    Close(choice == null ? "cancelled" : "selected");
                    if (choice == null) return;
                    try
                    {
                        if (choice == "more" || choice == "view") showPane(choice == "more");
                        else execute(choice);
                        lastOutcome = choice == "more" || choice == "view" ? "pane-opened" : "applied";
                        SafeLog("radial-command", new { gestureId, source = "middle-pointer", choice, outcome = lastOutcome });
                    }
                    catch (Exception exception)
                    {
                        lastOutcome = "command-failed"; lastError = exception.Message;
                        SafeLog("radial-command-error", new { gestureId, choice, error = exception.ToString() });
                        MessageBox.Show(new WindowOwner(owner), exception.Message, "PPT GCS", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    return;
                }
                if (!pendingMove.HasValue) return;
                NativePoint point = pendingMove.Value; pendingMove = null;
                Choose(point);
            }
        }

        private bool Begin(NativePoint point, IntPtr window)
        {
            uint process;
            IntPtr root = GetAncestor(window, 2);
            if (WindowClass(window) != "mdiClass" || root == IntPtr.Zero || root != GetAncestor(GetForegroundWindow(), 2) ||
                GetWindowThreadProcessId(window, out process) != uiThread || process != GetCurrentProcessId() || GetCapture() != IntPtr.Zero)
                return false;
            object currentWindow = activeWindow();
            if (currentWindow == null || (int)((dynamic)currentWindow).ViewType != 9) return false;
            NativeRect bounds;
            if (!GetWindowRect(window, out bounds)) throw new InvalidOperationException("无法读取 PowerPoint 画布范围。");
            Rectangle rectangle = Rectangle.FromLTRB(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
            if (!rectangle.Contains(point.X, point.Y)) return false;
            uint dpi = GetDpiForWindow(window);
            if (dpi == 0) throw new InvalidOperationException("无法读取 PowerPoint 画布 DPI。");
            opening = true;
            acceptedStart = true;
            pendingCancellation = null; pendingMove = null; pendingRelease = null;
            try
            {
                normalize();
                ActivateSlidePane(currentWindow);
                var model = getModel();
                if (model == null) throw new InvalidOperationException("无法读取当前对象选择。");
                canvasWindow = window; rootWindow = root; canvasBounds = rectangle;
                documentWindow = currentWindow; documentIdentity = ComIdentity(currentWindow);
                documentKey = WindowKey(currentWindow); selectionSignature = SelectionKey(model);
                selectedCount = model.Objects.Count(item => item.Selected);
                string unsupported = model.Objects.Where(item => item.Selected).Select(item => item.UnsupportedReason).FirstOrDefault(reason => reason != null);
                dpiScale = dpi / 96f;
                outerRadius = Math.Min(136 * dpiScale, Math.Min(rectangle.Width, rectangle.Height) / 2f - 12 * dpiScale);
                if (outerRadius < 70 * dpiScale) throw new InvalidOperationException("画布空间不足，请扩大窗口后使用轮盘。");
                innerRadius = 38 * dpiScale;
                center = new PointF(Clamp(point.X - rectangle.Left, outerRadius + 8 * dpiScale, rectangle.Width - outerRadius - 8 * dpiScale),
                    Clamp(point.Y - rectangle.Top, outerRadius + 8 * dpiScale, rectangle.Height - outerRadius - 8 * dpiScale));
                startPoint = point; currentSlot = -1; previewBounds.Clear();
                gestureId = Guid.NewGuid().ToString("D"); lastChoice = null; lastOutcome = "open"; lastError = null;
                foreach (Slot slot in slots)
                {
                    var option = model.Options.FirstOrDefault(item => item.Kind == slot.Key);
                    int minimum = option == null ? slot.Minimum : option.MinimumMembers;
                    int maximum = option == null ? 0 : option.MaximumMembers;
                    slot.Enabled = slot.Minimum == 0 || (unsupported == null && selectedCount >= minimum && (maximum == 0 || selectedCount <= maximum));
                    slot.Reason = slot.Enabled ? null : unsupported ?? "请先选中至少 " + minimum + " 个对象";
                    slot.Previewed = slot.Minimum == 0;
                    slot.Rectangles.Clear();
                }
                dynamic projection = currentWindow;
                originX = (double)projection.PointsToScreenPixelsX(0f);
                originY = (double)projection.PointsToScreenPixelsY(0f);
                scaleX = ((double)projection.PointsToScreenPixelsX(100f) - originX) / 100;
                scaleY = ((double)projection.PointsToScreenPixelsY(100f) - originY) / 100;
                if (scaleX <= 0 || scaleY <= 0 || Double.IsNaN(scaleX) || Double.IsNaN(scaleY))
                    throw new InvalidOperationException("幻灯片坐标映射无效。");
                hud = new RadialHud(this) { Bounds = rectangle };
                open = true;
                hud.Show(new WindowOwner(root));
                SetCapture(window);
                if (GetCapture() != window) throw new InvalidOperationException("无法捕获中键手势。");
                releasedSince = DateTime.MinValue;
                nextContextCheck = DateTime.UtcNow.AddMilliseconds(100);
                SafeLog("radial-open", new { gestureId, source = "middle-pointer", middleDown = KeyDown(MiddleKey), selectedCount,
                    canvas = window.ToInt64(), x = point.X, y = point.Y, dpi,
                    center = new { x = rectangle.Left + center.X, y = rectangle.Top + center.Y } });
                return true;
            }
            catch { Close("open-failed"); throw; }
            finally { opening = false; }
        }

        private void Choose(NativePoint point)
        {
            if (!open) return;
            int index = HitSlot(point);
            if (index == currentSlot) return;
            currentSlot = index; previewBounds.Clear();
            if (index >= 0)
            {
                Slot slot = slots[index];
                if (slot.Enabled && !slot.Previewed)
                {
                    try
                    {
                        var changed = preview(slot.Key) ?? new List<ShapeGeometry>();
                        if (!open) return;
                        foreach (var shape in changed)
                            slot.Rectangles.Add(new RectangleF((float)(originX + shape.Left * scaleX - canvasBounds.Left),
                                (float)(originY + shape.Top * scaleY - canvasBounds.Top),
                                (float)(shape.Width * scaleX), (float)(shape.Height * scaleY)));
                    }
                    catch (Exception exception)
                    {
                        slot.Enabled = false; slot.Reason = exception.Message;
                        SafeLog("radial-preview-rejected", new { gestureId, choice = slot.Key, reason = slot.Reason });
                    }
                    slot.Previewed = true;
                }
                if (slot.Enabled) previewBounds.AddRange(slot.Rectangles);
            }
            if (hud != null) hud.Invalidate();
            SafeLog("radial-choice", new { gestureId, source = "middle-pointer", choice = index < 0 ? null : slots[index].Key,
                enabled = index >= 0 && slots[index].Enabled, previewCount = previewBounds.Count, middleDown = KeyDown(MiddleKey) });
        }

        private int HitSlot(NativePoint point)
        {
            if (!canvasBounds.Contains(point.X, point.Y)) return -1;
            double travelX = point.X - startPoint.X, travelY = point.Y - startPoint.Y;
            if (travelX * travelX + travelY * travelY < 36 * dpiScale * dpiScale) return -1;
            double dx = point.X - canvasBounds.Left - center.X, dy = point.Y - canvasBounds.Top - center.Y;
            if (dx * dx + dy * dy < innerRadius * innerRadius) return -1;
            double angle = Math.Atan2(dy, dx) * 180 / Math.PI + 112.5;
            if (angle < 0) angle += 360;
            return ((int)(angle / 45)) % 8;
        }

        public void Tick()
        {
            RequireUiThread();
            if (disposed || !open || processing) return;
            try
            {
                if (!IsWindow(canvasWindow) || GetCapture() != canvasWindow) { Close("capture-lost"); return; }
                if (GetAncestor(GetForegroundWindow(), 2) != rootWindow) { Close("focus-lost"); return; }
                if (!KeyDown(MiddleKey))
                {
                    if (releasedSince == DateTime.MinValue) releasedSince = DateTime.UtcNow;
                    else if ((DateTime.UtcNow - releasedSince).TotalMilliseconds >= 500) { Close("release-message-lost"); return; }
                }
                else releasedSince = DateTime.MinValue;
                if (DateTime.UtcNow < nextContextCheck) return;
                nextContextCheck = DateTime.UtcNow.AddMilliseconds(100);
                NativeRect bounds;
                object current = activeWindow();
                if (!GetWindowRect(canvasWindow, out bounds) || Rectangle.FromLTRB(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom) != canvasBounds ||
                    current == null || ComIdentity(current) != documentIdentity || WindowKey(current) != documentKey ||
                    SelectionKey(getModel()) != selectionSignature)
                    Close("context-changed");
            }
            catch (Exception exception)
            { lastError = exception.Message; SafeLog("radial-context-error", new { gestureId, error = exception.ToString() }); Close("context-error"); }
        }

        public string GetState()
        {
            RequireUiThread();
            return json.Serialize(new {
                open = IsOpen, source = "middle-pointer", gestureId, selectedCount,
                currentChoice = IsOpen && currentSlot >= 0 ? slots[currentSlot].Key : null,
                middleDown = KeyDown(MiddleKey), lastChoice, lastOutcome, lastError,
                canvasWindow = canvasWindow.ToInt64(), rootWindow = rootWindow.ToInt64(),
                center = new { x = canvasBounds.Left + center.X, y = canvasBounds.Top + center.Y }, innerRadius, outerRadius,
                previewCount = previewBounds.Count,
                slots = slots.Select((slot, index) => new {
                    index, key = slot.Key, label = slot.Label, enabled = slot.Enabled, reason = slot.Reason,
                    x = canvasBounds.Left + center.X + Math.Cos((index * 45 - 90) * Math.PI / 180) * outerRadius * 0.72,
                    y = canvasBounds.Top + center.Y + Math.Sin((index * 45 - 90) * Math.PI / 180) * outerRadius * 0.72
                }).ToArray()
            });
        }

        private void Close(string reason)
        {
            bool wasOpen = IsOpen;
            open = false; opening = false;
            pendingMove = null; pendingRelease = null; pendingCancellation = null;
            if (GetCapture() == canvasWindow && canvasWindow != IntPtr.Zero) ReleaseCapture();
            if (hud != null) { hud.Hide(); hud.Dispose(); hud = null; }
            documentWindow = null; previewBounds.Clear();
            swallowMiddleUp = KeyDown(MiddleKey);
            if (wasOpen)
            {
                lastOutcome = reason;
                SafeLog("radial-close", new { gestureId, source = "middle-pointer", reason, middleDown = KeyDown(MiddleKey) });
            }
        }

        public void Dispose()
        { RequireUiThread(); if (disposed) return; Close("dispose"); disposed = true; }

        private static float Clamp(float value, float minimum, float maximum)
        { return Math.Max(minimum, Math.Min(maximum, value)); }
        private static string SelectionKey(ConstraintPaneModel model)
        { return String.Join("|", model.Objects.Where(item => item.Selected).Select(item => item.Id).OrderBy(id => id, StringComparer.Ordinal)); }
        private static string WindowKey(object windowObject)
        { dynamic window = windowObject; return (string)window.Presentation.FullName + "/" + (int)window.View.Slide.SlideID; }
        private static IntPtr ComIdentity(object value)
        { IntPtr identity = Marshal.GetIUnknownForObject(value); Marshal.Release(identity); return identity; }
        private static void ActivateSlidePane(object windowObject)
        {
            dynamic window = windowObject;
            if ((int)window.ActivePane.ViewType == 1) return;
            for (int i = 1; i <= (int)window.Panes.Count; i++) if ((int)window.Panes[i].ViewType == 1) { window.Panes[i].Activate(); break; }
            if ((int)window.ActivePane.ViewType != 1) throw new InvalidOperationException("请切换到幻灯片编辑画布。");
        }
        private static string WindowClass(IntPtr window)
        { var name = new StringBuilder(128); GetClassName(window, name, name.Capacity); return name.ToString(); }
        private static bool KeyDown(int key) { return (GetAsyncKeyState(key) & 0x8000) != 0; }
        private void RequireUiThread()
        { if (GetCurrentThreadId() != uiThread) throw new InvalidOperationException("The radial menu must run on PowerPoint's UI thread."); }
        private void SafeLog(string kind, object value)
        { try { log(kind, json.Serialize(value)); } catch { } }

        private sealed class Slot
        {
            public readonly string Key, Label;
            public readonly int Minimum;
            public bool Enabled, Previewed;
            public string Reason;
            public readonly List<RectangleF> Rectangles = new List<RectangleF>();
            public Slot(string key, string label, int minimum) { Key = key; Label = label; Minimum = minimum; }
        }

        private sealed class WindowOwner : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public WindowOwner(IntPtr handle) { Handle = handle; }
        }

        private sealed class RadialHud : Form
        {
            private readonly RadialMenuController owner;
            private readonly Font labelFont, detailFont, centerFont;
            private readonly Image[] icons;
            private static readonly Color ClearColor = Color.Magenta;
            public RadialHud(RadialMenuController owner)
            {
                this.owner = owner;
                // Office supplies native 16/32 px artwork; the former 24 px
                // request interpolated even at 100% display scaling.
                int imageSize=Math.Max(16,Math.Min(128,(int)Math.Round(32*owner.dpiScale)));
                icons=owner.slots.Select(slot=>owner.getIcon?.Invoke(slot.Key,imageSize)).ToArray();
                FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual; AutoScaleMode = AutoScaleMode.None;
                BackColor = ClearColor; TransparencyKey = ClearColor;
                AccessibleName = "PPT GCS 中键约束轮盘";
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
                labelFont = new Font(SystemFonts.MessageBoxFont.FontFamily, 14 * owner.dpiScale, FontStyle.Regular, GraphicsUnit.Pixel);
                detailFont = new Font(SystemFonts.MessageBoxFont.FontFamily, 12 * owner.dpiScale, FontStyle.Regular, GraphicsUnit.Pixel);
                centerFont = new Font(SystemFonts.MessageBoxFont.FontFamily, 12 * owner.dpiScale, FontStyle.Bold, GraphicsUnit.Pixel);
            }
            protected override bool ShowWithoutActivation { get { return true; } }
            protected override CreateParams CreateParams
            {
                get { var parameters = base.CreateParams; parameters.ExStyle |= 0x08000000 | 0x00000080 | 0x00000020; return parameters; }
            }
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x0021) { message.Result = new IntPtr(3); return; }
                if (message.Msg == 0x0084) { message.Result = new IntPtr(-1); return; }
                base.WndProc(ref message);
            }
            protected override void OnPaint(PaintEventArgs args)
            {
                Graphics graphics = args.Graphics;
                graphics.Clear(ClearColor); graphics.SmoothingMode = SmoothingMode.AntiAlias;
                float scale = owner.dpiScale, outer = owner.outerRadius, inner = owner.innerRadius;
                PointF center = owner.center;
                using (var previewPen = new Pen(Color.FromArgb(25, 102, 202), 2 * scale) { DashStyle = DashStyle.Dash })
                    foreach (var box in owner.previewBounds) graphics.DrawRectangle(previewPen, box.X, box.Y, box.Width, box.Height);
                RectangleF outside = new RectangleF(center.X - outer, center.Y - outer, outer * 2, outer * 2);
                RectangleF inside = new RectangleF(center.X - inner, center.Y - inner, inner * 2, inner * 2);
                for (int index = 0; index < owner.slots.Length; index++)
                {
                    Slot slot = owner.slots[index];
                    bool hover = owner.currentSlot == index;
                    using (var path = new GraphicsPath())
                    using (var fill = new SolidBrush(!slot.Enabled ? Color.FromArgb(237, 239, 242) : hover ? Color.FromArgb(216, 234, 255) : Color.FromArgb(250, 251, 253)))
                    using (var outline = new Pen(hover && slot.Enabled ? Color.FromArgb(25, 102, 202) : Color.FromArgb(147, 157, 171), hover ? 2 * scale : scale))
                    {
                        float start = index * 45 - 112.5f;
                        path.AddArc(outside, start, 45); path.AddArc(inside, start + 45, -45); path.CloseFigure();
                        graphics.FillPath(fill, path); graphics.DrawPath(outline, path);
                    }
                    double angle = (index * 45 - 90) * Math.PI / 180;
                    float anchorX=(float)(center.X+Math.Cos(angle)*outer*0.72);
                    float anchorY=(float)(center.Y+Math.Sin(angle)*outer*0.72);
                    Image icon=icons[index];
                    if(icon!=null)
                    {
                        int iconX=(int)(anchorX-icon.Width/2f), iconY=(int)(anchorY-28*scale);
                        if(slot.Enabled) graphics.DrawImageUnscaled(icon,iconX,iconY);
                        else ControlPaint.DrawImageDisabled(graphics,icon,iconX,iconY,Color.FromArgb(237,239,242));
                    }
                    var label = new Rectangle((int)(anchorX-41*scale),
                        (int)(anchorY+(icon==null?-18:7)*scale), (int)(82*scale), (int)(30*scale));
                    TextRenderer.DrawText(graphics, slot.Label, labelFont, label, slot.Enabled ? Color.FromArgb(25, 34, 47) : Color.FromArgb(130, 136, 145),
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
                }
                using (var fill = new SolidBrush(Color.White)) graphics.FillEllipse(fill, inside);
                var centerTop = new Rectangle((int)(center.X - inner + 2 * scale), (int)(center.Y - 18 * scale), (int)(2 * inner - 4 * scale), (int)(20 * scale));
                TextRenderer.DrawText(graphics, owner.selectedCount + " 个对象", centerFont, centerTop, Color.FromArgb(36, 44, 56), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                centerTop.Y += (int)(22 * scale);
                TextRenderer.DrawText(graphics, "中心取消", detailFont, centerTop, Color.FromArgb(105, 113, 126), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                string title = "按住中键，移向命令后松开";
                string detail = "返回中心或按 Esc 取消";
                if (owner.currentSlot >= 0)
                {
                    Slot slot = owner.slots[owner.currentSlot]; title = slot.Label;
                    detail = !slot.Enabled ? slot.Reason : slot.Minimum == 0 ? "松开中键打开约束窗格" :
                        "预览 " + owner.previewBounds.Count + " 个对象的变化 · 松开建立约束";
                }
                float width = Math.Min(430 * scale, ClientSize.Width - 12 * scale), height = 62 * scale;
                float x = Clamp(center.X - width / 2, 6 * scale, ClientSize.Width - width - 6 * scale);
                float y = center.Y + outer + 10 * scale;
                if (y + height > ClientSize.Height - 6 * scale) y = Math.Max(6 * scale, center.Y - outer - height - 10 * scale);
                var information = new RectangleF(x, y, width, height);
                using (var fill = new SolidBrush(Color.White)) graphics.FillRectangle(fill, information);
                using (var pen = new Pen(Color.FromArgb(147, 157, 171))) graphics.DrawRectangle(pen, x, y, width, height);
                var text = Rectangle.Round(new RectangleF(x + 8 * scale, y + 5 * scale, width - 16 * scale, 21 * scale));
                TextRenderer.DrawText(graphics, title, labelFont, text, Color.FromArgb(25, 34, 47), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                text.Y += (int)(24 * scale); text.Height = (int)(28 * scale);
                TextRenderer.DrawText(graphics, detail, detailFont, text, Color.FromArgb(74, 83, 97), TextFormatFlags.HorizontalCenter | TextFormatFlags.WordBreak | TextFormatFlags.EndEllipsis);
            }
            protected override void Dispose(bool disposing)
            {
                if (disposing) { labelFont.Dispose(); detailFont.Dispose(); centerFont.Dispose(); }
                base.Dispose(disposing);
            }
        }

        [StructLayout(LayoutKind.Sequential)] private struct NativePoint { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetCapture();
        [DllImport("user32.dll")] private static extern IntPtr SetCapture(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out NativeRect rectangle);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
    }
}
