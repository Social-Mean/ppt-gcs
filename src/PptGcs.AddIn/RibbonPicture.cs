using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace PptGcs.AddIn
{
    // Transfer a straight-alpha DIB to an owning IPictureDisp. AxHost's bitmap
    // conversion blends partial-alpha pixels into a system background first.
    internal static class RibbonPicture
    {
        internal static object Create(Bitmap image)
        {
            var header=new BitmapHeader {Size=40,Width=image.Width,Height=-image.Height,Planes=1,BitCount=32};
            IntPtr bits;
            IntPtr handle=CreateDIBSection(IntPtr.Zero,ref header,0,out bits,IntPtr.Zero,0);
            if(handle==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());
            try
            {
                if(bits==IntPtr.Zero)throw new Win32Exception(Marshal.GetLastWin32Error());
                var data=image.LockBits(new Rectangle(Point.Empty,image.Size),ImageLockMode.ReadOnly,PixelFormat.Format32bppArgb);
                try
                {
                    var row=new byte[checked(image.Width*4)];
                    for(int y=0;y<image.Height;y++)
                    {
                        Marshal.Copy(IntPtr.Add(data.Scan0,y*data.Stride),row,0,row.Length);
                        Marshal.Copy(row,0,IntPtr.Add(bits,y*row.Length),row.Length);
                    }
                }
                finally {image.UnlockBits(data);}
                var description=new PictureDescription {Size=Marshal.SizeOf(typeof(PictureDescription)),Type=1,Bitmap=handle};
                var iid=new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
                object picture;
                OleCreatePictureIndirect(ref description,ref iid,true,out picture);
                handle=IntPtr.Zero; // OLE now owns the DIB.
                return picture;
            }
            finally {if(handle!=IntPtr.Zero)DeleteObject(handle);}
        }

        [StructLayout(LayoutKind.Sequential)] private struct PictureDescription
        {public int Size,Type;public IntPtr Bitmap;public int Reserved1,Reserved2;}
        [StructLayout(LayoutKind.Sequential)] private struct BitmapHeader
        {public uint Size;public int Width,Height;public ushort Planes,BitCount;public uint Compression,SizeImage;public int XPelsPerMeter,YPelsPerMeter;public uint ColorsUsed,ColorsImportant;}
        [DllImport("gdi32.dll",SetLastError=true)] private static extern IntPtr CreateDIBSection(IntPtr dc,ref BitmapHeader header,uint usage,out IntPtr bits,IntPtr section,uint offset);
        [DllImport("gdi32.dll")] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool DeleteObject(IntPtr handle);
        [DllImport("oleaut32.dll",PreserveSig=false)] private static extern void OleCreatePictureIndirect(ref PictureDescription description,ref Guid iid,[MarshalAs(UnmanagedType.Bool)] bool own,[MarshalAs(UnmanagedType.IDispatch)] out object picture);
    }
}
