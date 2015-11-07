using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace PptToDocConverter
{
    /// <summary>
    /// Represents an ARGB pixel array of an image
    /// </summary>
    public class MemImage
    {
        public int Width { get; private set; }
        public int Height { get; private set; }

        private PixelData[][] array;

        public MemImage(int width, int height)
        {
            this.Width = width;
            this.Height = height;

            array = new PixelData[width][];
            for (int i = 0; i < width; i++)
                array[i] = new PixelData[height];
        }

        public MemImage(Image img)
            : this(img.Width, img.Height)
        {
            using (UnsafeBitmap bmp = new UnsafeBitmap(img))
            {
                bmp.LockBitmap();

                Parallel.For(0, Height, j =>
                {
                    //for (int j = 0; j < Height; j++)
                    //{

                    for (int i = 0; i < Width; i++)
                    {
                        PixelData pixelInput = bmp.GetPixel(i, j);
                        array[i][j] = pixelInput;
                    }
                    //}
                });
                bmp.UnlockBitmap();
            }
        }

        private MemImage(int width, int height, PixelData[][] array)
        {
            this.Width = width;
            this.Height = height;

            this.array = new PixelData[width][];
            for (int i = 0; i < width; i++)
            {
                this.array[i] = new PixelData[height];
                Array.Copy(array[i], this.array[i], height);
            }
        }

        public PixelData GetPixel(int x, int y)
        {
            return array[x][y];
        }

        //public Color GetPixelColor(int x, int y)
        //{
        //    return var pd = array[x][y];
        //    return Color.FromArgb(pd.R, pd.G, pd.B);
        //}

        public void SetPixel(int x, int y, PixelData c)
        {
            array[x][y] = c;
        }

        public void SetPixel(int x, int y, Color c)
        {
            PixelData pd = new PixelData()
            {
                R = c.R,
                G = c.G,
                B = c.B
            };
            array[x][ y] = pd;
        }

        public static MemImage Crop(MemImage input, int x, int y, int width, int height)
        {
            if (width <= 0 || height <= 0 || x < 0 || y < 0)
                return null;

            int iWidth = input.Width;
            int iHeight = input.Height;

            MemImage output = new MemImage(width, height);

            int jMax = (y + height > iHeight ? iHeight : y + height);
            int iMax = (x + width > iWidth ? iWidth : x + width);

            for (int j = y; j < jMax; j++)
            {
                for (int i = x; i < iMax; i++)
                {
                    PixelData pixelInput = input.GetPixel(i, j);
                    output.SetPixel(i - x, j - y, pixelInput);
                }
            }

            return output;
        }

        public Image ToImage()
        {
            Bitmap bmp = new Bitmap(Width, Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            int iWidth = Width;
            int iHeight = Height;
            UnsafeBitmap output = new UnsafeBitmap(iWidth, iHeight);
            output.LockBitmap();


            for (int j = 0; j < iHeight; j++)
            //   Parallel.For(0, iHeight, j =>
            {

                for (int i = 0; i < iWidth; i++)
                {
                    PixelData pixelInput = array[i][j];
                    output.SetPixel(i, j, pixelInput);
                }
                //});
            }


            output.UnlockBitmap();
            return output.Bitmap;

            //for (int j = 0; j < Height; j++)
            //{
            //    for (int i = 0; i < Width; i++)
            //    {
            //        bmp.SetPixel(i, j, array[i, j].ToColor());
            //    }
            //}
            //return bmp;
        }

        public MemImage Clone()
        {
            MemImage clone = new MemImage(Width, Height, this.array);
            return clone;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PixelData
    {
        public byte B;
        public byte G;
        public byte R;


        public static bool operator ==(PixelData p1, PixelData p2)
        {
            return p1.R == p2.R && p1.B == p2.B && p1.G == p2.G;
        }

        public static bool operator !=(PixelData p1, PixelData p2)
        {
            return p1.R != p2.R || p1.B != p2.B || p1.G != p2.G;
        }

        public static PixelData operator -(PixelData p1)
        {
            return new PixelData()
            {
                R = (byte)(255 - p1.R),
                G = (byte)(255 - p1.G),
                B = (byte)(255 - p1.B)
            };
        }

        public Color ToColor()
        {
            return Color.FromArgb(R, G, B);
        }
    }

}
