using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace ExcelPain
{
    public class App
    {
        public static void Main()
        {
            // settings
            int renderW = 64;
            int renderH = 64;
            string vbaFrameSleep = "0.5";

            // get frames directory path
            Console.WriteLine("Enter Frames directory (generate with ffmpeg, format .png): ");
            string framesDir = Console.ReadLine();

            // prepare vba
            StringBuilder vbaBody = new StringBuilder();
            StringBuilder mainFn = new StringBuilder();
            mainFn.AppendLine("Sub DrawMain()");

            // enumerate all pngs
            int frameNo = 0;
            foreach (string img in Directory.EnumerateFiles(framesDir, "*.png", SearchOption.TopDirectoryOnly))
                using (Bitmap bmp = new Bitmap(img).Resize(renderW, renderH))
                {
                    Console.WriteLine($"Prepare frame {frameNo}...");

                    // create frame from image
                    Frame frame = new Frame(bmp).SwapPrimaryAndSecondary();

                    // generate vba function
                    vbaBody.AppendLine(GenerateVBAForFrame(frame, new Rectangle(0, 0, renderW, renderH), frameNo++, mainFn));

                    //add sleep statement
                    mainFn.AppendLine($"Sleep {vbaFrameSleep}");
                }

            // finish and write vba
            mainFn.AppendLine("End Sub");
            vbaBody.AppendLine(mainFn.ToString());
            vbaBody.AppendLine(ExcelSleepFn());

            File.WriteAllText(@"", vbaBody.ToString());
        }

        /// <summary>
        /// Generate excel VBA code for a single frame
        /// </summary>
        /// <param name="frame"></param>
        /// <param name="fullFrame"></param>
        /// <param name="frameNo"></param>
        /// <param name="main"></param>
        /// <returns></returns>
        static string GenerateVBAForFrame(Frame frame, Rectangle fullFrame, int frameNo, StringBuilder main)
        {
            // generate rects 
            List<Rectangle> rects = frame.GetAsRectangles();
            Color primary = frame.GetPrimary();
            Color secondary = frame.GetSecondary();

            // generate function name
            string fnName = $"Frame{frameNo:000}";

            // prepare function for this frame
            StringBuilder frameFn = new StringBuilder();
            frameFn.AppendLine($"Sub {fnName}()");

            // clear frame first
            frameFn.AppendLine(@$"Range(""{fullFrame.ToExcelRange()}"").Interior.Color={secondary.ToExcelRGB()}");

            // add call for every rect
            foreach (Rectangle r in rects)
                frameFn.AppendLine(@$"Range(""{r.ToExcelRange()}"").Interior.Color={primary.ToExcelRGB()}");

            // end the frame function and add call to main
            frameFn.AppendLine("End Sub");
            main.AppendLine(fnName);
            return frameFn.ToString();
        }

        /// <summary>
        /// https://stackoverflow.com/a/53392427
        /// </summary>
        /// <returns></returns>
        static string ExcelSleepFn()
        {
            return @"
Sub Sleep(vSeconds As Variant)
    Dim t0 As Single, t1 As Single
    t0 = Timer
    Do
        t1 = Timer
        If t1 < t0 Then t1 = t1 + 86400
        DoEvents
    Loop Until t1 - t0 >= vSeconds
End Sub
";
        }

    }
}
