using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;
using DOC = Microsoft.Office.Interop.Word;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace PptToDocConverter
{
    static class Program
    {
        private static Options options;
        [STAThread]
        static void Main(string[] args)
        {
            ReadAndValidateOptions(args);

            if (options != null)
            {
                try
                {
                    ConvertPPTToDoc(options);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("Error: " + ex.GetType().FullName + " - " + ex.Message);
                }
            }

            if (System.Diagnostics.Debugger.IsAttached)
                Console.Read();

        }

        private static void ReadAndValidateOptions(string[] args)
        {
            options = new Options(args);
            if (string.IsNullOrEmpty(options.InPath) || !System.IO.File.Exists(options.InPath))
            {
                Console.WriteLine("Powerpoint file not specified or file does not exist (-in)");
                Options.PrintUsage();
                options = null;
                return;
            }

            if (string.IsNullOrEmpty(options.Outpath))
            {
                // if no output is specified use the same name but with a .docx extension
                var infile = new FileInfo(options.InPath).FullName;
                var outfile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(infile), System.IO.Path.GetFileNameWithoutExtension(infile) + ".docx");
                options.Outpath = outfile;
            }
        }


        private static void ConvertPPTToDoc(Options options)
        {
            PPT.ApplicationClass pptApp = new PPT.ApplicationClass();
            pptApp.DisplayAlerts = PPT.PpAlertLevel.ppAlertsNone;

            // make copy so we don't change the original
            string tempPptFilePath = System.IO.Path.Combine(System.IO.Path.GetTempFileName());
            System.IO.File.Copy(options.InPath, tempPptFilePath, true);

            PPT.Presentation pptPresentation = pptApp.Presentations.Open(new FileInfo(tempPptFilePath).FullName, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

            DOC.Application docApp = new DOC.Application();
            var doc = docApp.Documents.Add();

            Console.WriteLine("Converting powerpoint to word");
            ProgressBar progress = new ProgressBar();

            try
            {
                string currentTitle = "";
                string currentSubTitle = "";

                for (int slideNr = 1; slideNr <= pptPresentation.Slides.Count; slideNr++)
                {
                    progress.Report(slideNr / (float)pptPresentation.Slides.Count, "Slide " + slideNr + "/" + pptPresentation.Slides.Count);
                    CopySlide(pptPresentation, doc, options, ref currentTitle, ref currentSubTitle, slideNr);
                }
            }
            finally
            {
                pptPresentation.Close();

                if (System.IO.File.Exists(tempPptFilePath))
                    System.IO.File.Delete(tempPptFilePath);
            }
            //docApp.Visible = true;
            doc.SaveAs(new FileInfo(options.Outpath).FullName);
            doc.Close();

            progress.Dispose();

            Console.WriteLine("Conversion complete, word file written to " + new FileInfo(options.Outpath).FullName);

            try
            {
                docApp.Quit();
                pptApp.Quit();
            }
            catch (Exception)
            {
            }
        }


        private static void CopySlide(PPT.Presentation pptPresentation, DOC.Document doc, Options options, ref string currentTitle, ref string currentSubTitle, int slideNr)
        {
            var slide = pptPresentation.Slides[slideNr];

            string title = "";
            if (slide.Shapes.HasTitle != Microsoft.Office.Core.MsoTriState.msoFalse && slide.Shapes.Title.TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoFalse)
                title = slide.Shapes.Title.TextFrame.TextRange.Text;

            string newTitle;
            string newSubtitle;
            if (title.Contains(options.TitleDelimiter))
            {
                newTitle = title.Substring(0, title.IndexOf(options.TitleDelimiter));
                newSubtitle = title.Substring(title.IndexOf(options.TitleDelimiter) + 1).Trim();

                // case subtitle
                if (newSubtitle.Length > 0 && char.IsLower(newSubtitle[0]))
                    newSubtitle = char.ToUpper(newSubtitle[0]) + newSubtitle.Substring(1);
            }
            else
            {
                newTitle = title;
                newSubtitle = "";
            }



            if (options.Headings)
            {
                if (!string.IsNullOrEmpty(newTitle) && newTitle.ToLower() != currentTitle.ToLower())
                {
                    // main title has changed, insert a new heading 1
                    currentTitle = newTitle;
                    var par = doc.Paragraphs.Add();
                    par.Range.Select();
                    par.Range.Text = title;
                    par.set_Style(DOC.WdBuiltinStyle.wdStyleHeading1);
                    par.Range.InsertParagraphAfter();
                }

                if (!string.IsNullOrEmpty(newSubtitle) && newSubtitle.ToLower() != currentSubTitle.ToLower())
                {
                    // subtitle has changed, insert a new heading 2
                    currentSubTitle = newSubtitle;
                    var par = doc.Paragraphs.Add();
                    par.Range.Select();
                    par.Range.Text = currentSubTitle;
                    par.set_Style(DOC.WdBuiltinStyle.wdStyleHeading2);
                    par.Range.InsertParagraphAfter();
                }
            }

            if (options.Slides)
            {
                if (options.RemoveTheme)
                {
                    // delete everything from the master slide to remove the theme
                    if (slide.Master.Shapes.Count > 0)
                        slide.Master.Shapes.Range().Delete();
                }

                if (options.RemoveSlideNumbers)
                {
                    // search for all shapes that act as placeholder for a slide number and delete them
                    for (int i = slide.Shapes.Count; i >= 1; i--)
                    {
                        if (slide.Shapes[i].HasTextFrame != Microsoft.Office.Core.MsoTriState.msoFalse && slide.Shapes[i].TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoFalse)
                        {
                            try
                            {
                                if (slide.Shapes[i].PlaceholderFormat.Type == PPT.PpPlaceholderType.ppPlaceholderSlideNumber)
                                    slide.Shapes[i].Delete();
                            }
                            catch (Exception)
                            {
                                // silently fail if it's not a placeholder
                            }
                        }
                    }
                }

                // export the slide to an temp image
                var imgPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "slide.png");
                slide.Export(imgPath, "PNG", 1600, 1200);

                // crop if required
                if (options.CropWidth || options.CropHeight)
                    CropSlideImage(imgPath, options);

                // insert the image in the word document
                var parImg = doc.Paragraphs.Add();
                parImg.Range.Select();
                parImg.Alignment = DOC.WdParagraphAlignment.wdAlignParagraphCenter;

                var pic = parImg.Range.InlineShapes.AddPicture(imgPath);
                pic.ScaleWidth = 35f;
                pic.ScaleHeight = 35f;

                parImg.Range.InsertParagraphAfter();
            }

            if (options.Notes && HasNotes(slide))
            {
                // copy paste notes through the clipboard
                var parNotes = doc.Paragraphs.Add();
                for (int i = 1; i <= slide.NotesPage.Shapes.Count; i++)
                {
                    if (slide.NotesPage.Shapes[i].TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoFalse)
                    {
                        if (slide.NotesPage.Shapes[i].TextFrame.TextRange.Text.Length > 3)
                        {
                            slide.NotesPage.Shapes[i].TextFrame.TextRange.Copy();
                            parNotes.Range.Paste();
                        }
                    }
                }
                parNotes.Range.InsertParagraphAfter();
            }
        }

        private static void CropSlideImage(string imgPath, Options options)
        {
            Rectangle bounds;
            using (MemoryStream ms = new MemoryStream(System.IO.File.ReadAllBytes(imgPath)))
            {
                var img = new MemImage(System.Drawing.Bitmap.FromStream(ms));

                int left = 0;
                int top = 0;
                int right = img.Width - 1;
                int bottom = img.Height - 1;
                if (options.CropWidth)
                {
                    // scan left side until we encounter a non white pixel
                    for (int i = 0; i < img.Width; i++)
                    {
                        if (!Enumerable.Range(0, img.Height).All(h => { var pd = img.GetPixel(i, h); return pd.R == 255 && pd.G == 255 && pd.B == 255; }))
                        {
                            left = i;
                            break;
                        }
                    }
                    
                    // scan right side until we encounter a non white pixel
                    for (int i = img.Width - 1; i >= 0; i--)
                    {
                        if (!Enumerable.Range(0, img.Height).All(h => { var pd = img.GetPixel(i, h); return pd.R == 255 && pd.G == 255 && pd.B == 255; }))
                        {
                            right = i;
                            break;
                        }
                    }
                    left -= options.CropPadding;
                    right += options.CropPadding;
                    if (left < 0) left = 0;
                    if (right > img.Width - 1) right = img.Width - 1;

                }

                if (options.CropHeight)
                {
                    // scan top side until we encounter a non white pixel
                    for (int i = 0; i < img.Height; i++)
                    {
                        if (!Enumerable.Range(0, img.Width).All(w => { var pd = img.GetPixel(w, i); return pd.R == 255 && pd.G == 255 && pd.B == 255; }))
                        {
                            top = i;
                            break;
                        }
                    }

                    // scan bottom side until we encounter a non white pixel
                    for (int i = img.Height - 1; i >= 0; i--)
                    {
                        if (!Enumerable.Range(0, img.Width).All(w => { var pd = img.GetPixel(w, i); return pd.R == 255 && pd.G == 255 && pd.B == 255; }))
                        {
                            bottom = i;
                            break;
                        }
                    }
                    top -= options.CropPadding;
                    bottom += options.CropPadding;
                    if (top < 0) top = 0;
                    if (bottom > img.Height - 1) bottom = img.Height - 1;
                }
                bounds = Rectangle.FromLTRB(left, top, right, bottom);

                if (bounds.Left == 0 && bounds.Top == 0 && bounds.Width == img.Width - 1 && bounds.Bottom == img.Height - 1)
                {
                    // nothing to crop
                }
                else
                {
                    // crop & overwrite the original image
                    img = MemImage.Crop(img, bounds.X, bounds.Y, bounds.Width, bounds.Height);
                    using (var bmp = img.ToImage())
                        bmp.Save(imgPath, ImageFormat.Png);

                }
            }
        }

        private static bool HasNotes(PPT.Slide slide)
        {
            // check if there are any shapes with text in the notespage that has more than 3 characters of text
            for (int i = 1; i <= slide.NotesPage.Shapes.Count; i++)
            {
                if (slide.NotesPage.Shapes[i].TextFrame.HasText != Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    if (slide.NotesPage.Shapes[i].TextFrame.TextRange.Text.Length > 3)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
