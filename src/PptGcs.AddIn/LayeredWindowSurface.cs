using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace PptGcs.AddIn
{
    // UpdateLayeredWindow consumes a top-down, premultiplied 32-bit BGRA DIB.
    // Keep the DC/DIB and transfer buffer for the lifetime of this HUD surface.
    internal sealed class LayeredWindowSurface : IDisposable
    {
        private readonly Size size;
        private readonly byte[] pixels;
        private IntPtr dc, bitmap, previousBitmap, bits;
        private bool disposed;

        public LayeredWindowSurface(Size size)
        {
            if (size.Width <= 0 || size.Height <= 0) throw new ArgumentOutOfRangeException("size");
            this.size = size;
            pixels = new byte[checked(size.Width * size.Height * 4)];
            try
            {
                dc = CreateCompatibleDC(IntPtr.Zero);
                if (dc == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
                var header = new BitmapInfoHeader { Size = 40, Width = size.Width, Height = -size.Height, Planes = 1, BitCount = 32 };
                bitmap = CreateDIBSection(dc, ref header, 0, out bits, IntPtr.Zero, 0);
                if (bitmap == IntPtr.Zero || bits == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
                previousBitmap = SelectObject(dc, bitmap);
                if (previousBitmap == IntPtr.Zero || previousBitmap == new IntPtr(-1))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            catch { Dispose(); throw; }
        }

        public void Present(IntPtr window, Bitmap image, Point location)
        {
            if (disposed) throw new ObjectDisposedException("LayeredWindowSurface");
            if (image == null || image.Size != size || image.PixelFormat != PixelFormat.Format32bppPArgb)
                throw new ArgumentException("The layered surface requires a matching premultiplied bitmap.", "image");
            var data = image.LockBits(new Rectangle(Point.Empty, size), ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
            try
            {
                int rowBytes = size.Width * 4;
                if (data.Stride == rowBytes) Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
                else for (int y = 0; y < size.Height; y++) Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), pixels, y * rowBytes, rowBytes);
            }
            finally { image.UnlockBits(data); }
            Marshal.Copy(pixels, 0, bits, pixels.Length);
            var destination = new NativePoint { X = location.X, Y = location.Y };
            var source = new NativePoint();
            var extent = new NativeSize { Width = size.Width, Height = size.Height };
            var blend = new BlendFunction { SourceConstantAlpha = 255, AlphaFormat = 1 };
            if (!UpdateLayeredWindow(window, IntPtr.Zero, ref destination, ref extent, dc, ref source, 0, ref blend, 2))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            if (dc != IntPtr.Zero && previousBitmap != IntPtr.Zero && previousBitmap != new IntPtr(-1)) SelectObject(dc, previousBitmap);
            if (bitmap != IntPtr.Zero) DeleteObject(bitmap);
            if (dc != IntPtr.Zero) DeleteDC(dc);
            dc = bitmap = previousBitmap = bits = IntPtr.Zero;
        }

        [StructLayout(LayoutKind.Sequential)] private struct NativePoint { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct NativeSize { public int Width, Height; }
        [StructLayout(LayoutKind.Sequential, Pack = 1)] private struct BlendFunction
        { public byte BlendOperation, BlendFlags, SourceConstantAlpha, AlphaFormat; }
        [StructLayout(LayoutKind.Sequential)] private struct BitmapInfoHeader
        { public uint Size; public int Width, Height; public ushort Planes, BitCount; public uint Compression, SizeImage; public int XPelsPerMeter, YPelsPerMeter; public uint ColorsUsed, ColorsImportant; }
        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr CreateCompatibleDC(IntPtr dc);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr CreateDIBSection(IntPtr dc, ref BitmapInfoHeader info, uint usage, out IntPtr pixels, IntPtr section, uint offset);
        [DllImport("gdi32.dll", SetLastError = true)] private static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);
        [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr dc);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool UpdateLayeredWindow(IntPtr window, IntPtr destinationDc,
            ref NativePoint destination, ref NativeSize size, IntPtr sourceDc, ref NativePoint source, uint colorKey, ref BlendFunction blend, uint flags);
    }
}
