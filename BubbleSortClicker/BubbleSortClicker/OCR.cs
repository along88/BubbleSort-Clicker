using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;


public class OCR
{
    public string TextFile { get; private set; } //textfile name we save to, may be useless now

    

    /// <summary>
    /// Reads text on screen at the given coordinates
    /// </summary>
    /// <param name="xy1"></param>
    /// <param name="xy2"></param>
    public void ReadScreen(string xy1, string xy2)
    {
        using (Process ocr = new Process())
        {
            ///<summary>
            ///TODO: default working directory to Capture2Text executable for user
            ///</summary>
            
            TextFile = "OCRBlock.txt";
            try
            {
                ocr.StartInfo.FileName = "cmd.exe";
                ocr.StartInfo.WorkingDirectory = @"C:\Users\along\Downloads\New Folder\Capture2Text";
                ocr.StartInfo.Arguments = string.Format("/C Capture2Text.exe {0} {1} {2} ", xy1, xy2, TextFile);
                ocr.StartInfo.CreateNoWindow = false;
                ocr.Start();
                ocr.WaitForExit(100000);
                Console.WriteLine(ocr.StartInfo.WorkingDirectory);
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                ReadScreen(xy1,xy2);
            }
            
        }
    }

    
}

