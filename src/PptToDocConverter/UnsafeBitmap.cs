using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;

namespace PptToDocConverter
{
    public unsafe class UnsafeBitmap : IDisposable
    {
        Bitmap bitmap;

        // three elements used for MakeGreyUnsafe
        int width;
        BitmapData bitmapData = null;
        Byte* pBase = null;

        public UnsafeBitmap(Image bitmap)
        {
            this.bitmap = new Bitmap(bitmap);
        }

        public UnsafeBitmap(int width, int height)
        {
            this.bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
        }

        public void Dispose()
        {
            bitmap.Dispose();
        }

        public Bitmap Bitmap
        {
            get
            {
                return (bitmap);
            }
        }

        private Point PixelSize
        {
            get
            {
                GraphicsUnit unit = GraphicsUnit.Pixel;
                RectangleF bounds = bitmap.GetBounds(ref unit);

                return new Point((int)bounds.Width, (int)bounds.Height);
            }
        }

        private bool isLocked;

        public void LockBitmap()
        {
            isLocked = true;

            GraphicsUnit unit = GraphicsUnit.Pixel;
            RectangleF boundsF = bitmap.GetBounds(ref unit);
            Rectangle bounds = new Rectangle((int)boundsF.X,
          (int)boundsF.Y,
          (int)boundsF.Width,
          (int)boundsF.Height);

            // Figure out the number of bytes in a row
            // This is rounded up to be a multiple of 4
            // bytes, since a scan line in an image must always be a multiple of 4 bytes
            // in length. 
            width = (int)boundsF.Width * sizeof(PixelData);
            if (width % 4 != 0)
            {
                width = 4 * (width / 4 + 1);
            }
            bitmapData =
          bitmap.LockBits(bounds, ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

            pBase = (Byte*)bitmapData.Scan0.ToPointer();
        }

        public PixelData GetPixel(int x, int y)
        {
            if (isLocked)
            {
                PixelData returnValue = *PixelAt(x, y);
                return returnValue;
            }
            else
            {
                Color c = bitmap.GetPixel(x, y);
                PixelData data = new PixelData()
                {
                    R = c.R,
                    G = c.G,
                    B = c.B
                };
                return data;

            }
        }

        public void SetPixel(int x, int y, PixelData colour)
        {
            if (isLocked)
            {
                PixelData* pixel = PixelAt(x, y);
                *pixel = colour;
            }
            else
            {
                Color c = Color.FromArgb(colour.R, colour.G, colour.B);
                bitmap.SetPixel(x, y, c);
            }
        }

        public void UnlockBitmap()
        {
            isLocked = false;

            if (bitmapData != null)
            {
                bitmap.UnlockBits(bitmapData);
                bitmapData = null;
                pBase = null;
            }
        }
        public PixelData* PixelAt(int x, int y)
        {
            return (PixelData*)(pBase + y * width + x * sizeof(PixelData));
        }

        void IDisposable.Dispose()
        {
            if (bitmap != null)
            {
                bitmap.Dispose();
            }
        }
    }

}