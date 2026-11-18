using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;

namespace PptGcs.AddIn
{
    // The images are borrowed by the views and owned by this application-lifetime cache.
    // GetImageMso must run inside PowerPoint: IPictureDisp does not marshal reliably
    // to an external automation process. No Office calls occur while painting.
    public sealed class OfficeIconProvider : IDisposable
    {
        private readonly object application;
        private readonly Action<string, string> report;
        private readonly Dictionary<string, Image> images = new Dictionary<string, Image>();
        private readonly Dictionary<string, string> errors = new Dictionary<string, string>();
        private bool disposed;

        // Names from OfficeDev/office-fluent-ui-command-identifiers, PowerPoint.
        // Equal-size glyphs use Office's equal-column/equal-row commands.
        private static readonly Dictionary<string, string> identifiers = new Dictionary<string, string> {
            {"SameX", "ObjectsAlignLeftSmart"}, {"RightAligned", "ObjectsAlignRightSmart"},
            {"CenterXAligned", "ObjectsAlignCenterHorizontalSmart"}, {"SameY", "ObjectsAlignTopSmart"},
            {"BottomAligned", "ObjectsAlignBottomSmart"}, {"CenterYAligned", "ObjectsAlignMiddleVerticalSmart"},
            {"CentersCoincident", "DiagramTargetInsertClassic"},
            {"SameWidth", "TableColumnsDistribute"}, {"SameHeight", "TableRowsDistribute"},
            {"EqualHorizontalGap", "AlignDistributeHorizontally"}, {"EqualVerticalGap", "AlignDistributeVertically"},
            {"FixedWidth", "ShapeWidth"}, {"FixedHeight", "ObjectSizeAndPositionDialog"},
            {"FixedPosition", "LockObject"}, {"AspectRatio", "PowerPointPageSetup"},
            {"FixedHorizontalGap", "AlignDistributeHorizontally"}, {"FixedVerticalGap", "AlignDistributeVertically"},
            {"Offset", "ObjectSizeAndPositionDialog"},
            {"SymmetricVertical", "ObjectFlipHorizontal"}, {"SymmetricHorizontal", "ObjectFlipVertical"},
            {"PageLeft", "ObjectsAlignLeftSmart"}, {"ContainerLeft", "ObjectsAlignLeftSmart"},
            {"PageCenterX", "ObjectsAlignCenterHorizontalSmart"}, {"ContainerCenterX", "ObjectsAlignCenterHorizontalSmart"},
            {"PageRight", "ObjectsAlignRightSmart"}, {"ContainerRight", "ObjectsAlignRightSmart"},
            {"PageTop", "ObjectsAlignTopSmart"}, {"ContainerTop", "ObjectsAlignTopSmart"},
            {"PageCenterY", "ObjectsAlignMiddleVerticalSmart"}, {"ContainerCenterY", "ObjectsAlignMiddleVerticalSmart"},
            {"PageBottom", "ObjectsAlignBottomSmart"}, {"ContainerBottom", "ObjectsAlignBottomSmart"},
            {"more", "ObjectAlignMenu"}, {"view", "SelectionPane"}, {"pageSize", "PowerPointPageSetup"},
            {"copyComponent", "Copy"}, {"parameters", "ObjectSizeAndPositionDialog"},
            {"remove", "Clear"}, {"demo", "SlideNew"}, {"status", "FileProperties"}
        };

        public OfficeIconProvider(object application, Action<string, string> report)
        { this.application = application ?? throw new ArgumentNullException("application"); this.report = report; }

        public static IEnumerable<string> Commands { get { return identifiers.Keys; } }
        public static string ImageMso(string command)
        { string id; return command != null && identifiers.TryGetValue(command, out id) ? id : null; }

        public Image Get(string command, int size)
        {
            if (disposed) throw new ObjectDisposedException("OfficeIconProvider");
            if (size < 16 || size > 128) throw new ArgumentOutOfRangeException("size");
            string id = ImageMso(command);
            if (id == null) return null;
            string key = id + ":" + size;
            Image result;
            if (images.TryGetValue(key, out result)) return result;
            if (errors.ContainsKey(key)) return null;
            object picture = null;
            try
            {
                picture = ((dynamic)application).CommandBars.GetImageMso(id, size, size);
                result = CopyPicture(picture);
                images.Add(key, result);
                return result;
            }
            catch (Exception exception)
            {
                // Labels and commands remain available if Office cannot supply a glyph.
                errors.Add(key, exception.Message);
                report?.Invoke(id, exception.Message);
                return null;
            }
            finally
            {
                if (picture != null && Marshal.IsComObject(picture)) Marshal.ReleaseComObject(picture);
            }
        }

        public object Diagnostics()
        {
            return identifiers.Select(pair => {
                Image image = Get(pair.Key, 24);
                string error; errors.TryGetValue(pair.Value + ":24", out error);
                return new { command = pair.Key, imageMso = pair.Value, loaded = image != null,
                    width = image == null ? 0 : image.Width, height = image == null ? 0 : image.Height, error };
            }).ToArray();
        }

        private static Image CopyPicture(object picture)
        {
            dynamic source = picture;
            var handle = new IntPtr(unchecked((long)(uint)(int)source.Handle));
            int type = (int)source.Type;
            if (type == 3)
            { using (var icon = Icon.FromHandle(handle)) return icon.ToBitmap(); }
            if (type != 1) throw new InvalidOperationException("Office returned an unsupported picture type.");
            NativeBitmap native;
            if (GetObject(handle, Marshal.SizeOf(typeof(NativeBitmap)), out native) == 0)
                throw new InvalidOperationException("Cannot read the Office icon bitmap.");
            int width = native.Width, height = Math.Abs(native.Height);
            var information = new BitmapInfoHeader { Size = 40, Width = width, Height = -height,
                Planes = 1, BitCount = 32, Compression = 0 };
            var pixels = new byte[checked(width * height * 4)];
            IntPtr dc = GetDC(IntPtr.Zero);
            try
            {
                if (GetDIBits(dc, handle, 0, (uint)height, pixels, ref information, 0) != height)
                    throw new InvalidOperationException("Cannot copy the Office icon pixels.");
            }
            finally { if (dc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, dc); }
            // GetImageMso returns straight BGRA, including RGB values greater
            // than alpha on antialiased edges. Marking those bytes as PArgb
            // brightens the edges and corrupts compositing onto the pane.
            // A source format with no alpha channel is opaque by definition.
            if (native.BitsPixel < 32) for (int i = 3; i < pixels.Length; i += 4) pixels[i] = 255;
            var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            try
            {
                var bits = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                try
                {
                    for (int y = 0; y < height; y++)
                        Marshal.Copy(pixels, y * width * 4, IntPtr.Add(bits.Scan0, y * bits.Stride), width * 4);
                }
                finally { bitmap.UnlockBits(bits); }
                return bitmap;
            }
            catch { bitmap.Dispose(); throw; }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            foreach (var image in images.Values) image.Dispose();
            images.Clear(); errors.Clear();
        }

        [StructLayout(LayoutKind.Sequential)] private struct NativeBitmap
        { public int Type, Width, Height, WidthBytes; public ushort Planes, BitsPixel; public IntPtr Bits; }
        [StructLayout(LayoutKind.Sequential)] private struct BitmapInfoHeader
        { public uint Size; public int Width, Height; public ushort Planes, BitCount; public uint Compression, SizeImage; public int XPelsPerMeter, YPelsPerMeter; public uint ClrUsed, ClrImportant; }
        [DllImport("gdi32.dll", EntryPoint="GetObjectW")] private static extern int GetObject(IntPtr bitmap, int size, out NativeBitmap value);
        [DllImport("gdi32.dll")] private static extern int GetDIBits(IntPtr dc, IntPtr bitmap, uint start, uint lines, byte[] bits, ref BitmapInfoHeader information, uint usage);
        [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr window);
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);
    }
}
