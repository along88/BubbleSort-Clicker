using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace BubbleSortClicker
{
    class Program
    {
        static void Main(string[] args)
        {
            ///<summary>
            ///TODO:
            ///Add Spacebar repositioning and firing off OCR process from start to until we reach bottom of form and all items are sorted
            ///Add Check for what state we are working with
            ///Look into Auto Launching a screen coordinate finder tool for user
            ///can possible grab a reference to the coordinates so they dont have to manually enter them, only confirm(wishful thinking)
            ///Try other sorting algorithm
            /// 
            /// Procedure:
            /// The program will launch a OCR executable tool called Capture2Text which is preinstalled
            /// Capture2Text launches through cmd prompt and OCR your monitor at the specified coordinates
            /// any text capture in those coordinates gets saved to a standard .txt file
            /// NOTE: Capture2Text does a decent job of maintaining linebreaks and has fairly decent accuracy on reads
            /// using a compnay provided excel spread sheet as reference, our code essentially cleans up the .txt file
            /// compares it to the spread sheet, using indexes then decides what order the items in the policy form window
            /// need to be moved. Using mouse click events and some decent but constant measurements of line spacing and button
            /// positiosn within the policy form window, the program sends a series of mouse movements and clicks to high-
            /// light the item that needs to be moved, and moves it with the policy form window 'move buttons" using a
            /// bubble sort algorithm.
            /// 
            ///</ summary >

            #region fields
            OCR ocr = new OCR();
            Data data = new Data();
            AITask aiTask = new AITask();
            //ScrollbarPOs can be uncommented when ready to implment scrollbar movement logic
            //int scrollbarPOs = 120; 
            #endregion


            ConsoleColor defaultColor = Console.ForegroundColor;
            Console.WriteLine("Please wait, Recognizing Policy Form Window...");
            
            //OCR the Policy Forms - Ordering Window
            ocr.ReadScreen("1288 43", "2437 990"); //change these xy ,xy2 coordinates to match your screen

            
            Console.Write("One moment, I'm tyring to make sense of what I read on the screen");
            Thread.Sleep(1000);
            Console.WriteLine("....");

            //Grab the OCR's output text file and reformat then trim it
            data.FormatTxt(ocr.TextFile);

            Console.WriteLine("Comparing what I read to whats in the spreadsheet, one moment please!");
            Thread.Sleep(1000);
            
            //Trim the text we've formated so the form number is first
            data.FormatedList = data.cleanList(data.StringList);
            
            //Populate the excel spreadsheet with the approiate spreadsheet
            data.LoadExcel();
            
            Console.WriteLine("I'm Ready to sort now!");
            Thread.Sleep(2000);
            Console.WriteLine("Once this Process Begins it MUST NOT be interrupted");
            Thread.Sleep(1000);
            Console.Write("Please the ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("Enter");
            Console.ForegroundColor = defaultColor;
            Console.WriteLine("key to begin");
            Console.ReadLine();

            
            Console.ForegroundColor = ConsoleColor.Red; 
            Console.WriteLine("Sorting items please do not touch this machine until process is complete...");
            Console.ForegroundColor = defaultColor;

            //Begin sorting proccess
            aiTask.BubbleSort(data.CompareData());

            Console.WriteLine("Sorting finished");

            //After sorting the max 72 items, send click event to move screendown 
            Console.ReadLine();
        }
    }
}
