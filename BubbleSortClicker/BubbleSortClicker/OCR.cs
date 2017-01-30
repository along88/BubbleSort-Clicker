using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;


static class OCR
{
    public static string textfile; //textfile name we save to, may be useless now

    /// <summary>
    /// Reads text on screen at the given coordinates
    /// </summary>
    /// <param name="xy1"></param>
    /// <param name="xy2"></param>
    public static void ReadScreen(string xy1, string xy2)
    {
        using (Process ocr = new Process())
        {
            textfile = "OCRBlock.txt";
            ocr.StartInfo.FileName = "cmd.exe";
            ocr.StartInfo.WorkingDirectory = @"C:\Users\along\Downloads\New Folder\Capture2Text";
            ocr.StartInfo.Arguments = string.Format("/C Capture2Text.exe {0} {1} {2} ", xy1, xy2, textfile);
            ocr.StartInfo.CreateNoWindow = false;
            ocr.Start();
            ocr.WaitForExit(100000);

        }
    }
}

