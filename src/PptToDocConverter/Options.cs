using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PptToDocConverter
{
    public class Options
    {
        public Options(params string[] args)
        {
            TitleDelimiter = ':';
            CropPadding = 3;

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToLower() == "-in") InPath = GetArgValue(args, ref i);
                else if (args[i].ToLower() == "-out") Outpath = GetArgValue(args, ref i);
                else if (args[i].ToLower() == "-headings") Headings = true;
                else if (args[i].ToLower() == "-titledelimiter") TitleDelimiter = GetArgValue(args, ref i)[0];
                else if (args[i].ToLower() == "-removetheme") RemoveTheme = true;
                else if (args[i].ToLower() == "-removeslidenumbers") RemoveSlideNumbers = true;
                else if (args[i].ToLower() == "-slides") Slides = true;
                else if (args[i].ToLower() == "-crop")
                {
                    string value = GetArgValue(args, ref i);
                    if (value.ToLower() == "h")
                        CropHeight = true;
                    else if (value.ToLower() == "w")
                        CropWidth = true;
                    else
                    {
                        CropWidth = true;
                        CropHeight = true;
                    }
                }
                else if (args[i].ToLower() == "-notes") Notes = true;
            }
        }

        private static string GetArgValue(string[] args, ref int i)
        {
            if (i + 1 >= args.Length)
                return "";

            if (args[i + 1].StartsWith("-") && args[i + 1] != "-")
                return "";

            return args[++i];
        }

        public static void PrintUsage()
        {
            Console.WriteLine("USAGE: ");
            Console.WriteLine(System.IO.Path.GetFileName(System.Reflection.Assembly.GetEntryAssembly().Location) + " -in <ppt(x)file> [-out <doc(x)file>] [OPTIONS...]");
            Console.WriteLine(
@"OPTIONS:
    -slides: Include slides as images
    -crop <w/h/wh>: Removes the whitespace horizontally (w), vertically (h)
                    or both (wh). Keeping the whitespace horizontally keeps 
                    the slides nicely aligned under each other
    -removetheme: Removes the master slides from each slide before 
                  converting to an image
    -removeslidenumbers: Removes all shapes that contain slide numbers 
                         before converting to an image

    -notes: Include notes as text

    -headings: Convert slide titles to headings
    -titledelimiter <delimiter>: Split character for splitting the slide 
                                 title into heading 1 and heading 2 
                                 (by default ':')

Note: The conversion uses COM Office Interop (>= v12 or Office 2007)
      to read the powerpoint and write the word document. 
      It also uses the .Copy() and .Paste functionality of TextRanges
      to copy the notes to the a paragraph in word and to ensure 
      the formatting is kept. This means that during conversion you
      should refrain from changing the clipboard.
");


        }



        public string InPath { get; set; }
        public string Outpath { get; set; }

        public bool Headings { get; set; }

        public bool RemoveTheme { get; set; }

        public bool Slides { get; set; }
        public bool CropWidth { get; set; }
        public bool CropHeight { get; set; }
        public int CropPadding { get; set; }

        public bool Notes { get; set; }

        public char TitleDelimiter { get; set; }

        public bool RemoveSlideNumbers { get; set; }
    }
}
