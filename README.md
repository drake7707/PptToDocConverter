# Ppt To Doc Converter
Converts powerpoints to word documents using COM Interop to provide more options than just printing with notes

USAGE:
PptToDocConverter.exe -in <ppt(x)file> [-out <doc(x)file>] [OPTIONS...]
OPTIONS:
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
