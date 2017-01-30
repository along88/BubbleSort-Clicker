using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BubbleSortClicker
{
    class Program
    {
        #region Properties
        static string state;
        public static int scrollbarPOs = 120;

        #endregion

        static void Main(string[] args)
        {

            //Add Check for what state we are working with
            //Note: Fun idea I had, maybe we can have the OCR get the state for us as well
            //Console.WriteLine("Please select your state")

            Console.WriteLine("Please wait, Recognizing Policy Form Window...");

            //OCR the Policy Forms - Ordering Window
            OCR.ReadScreen("1288 43", "2437 990"); //change these xy ,xy2 coordinates to match your screen

            //Grab the OCR's output text file and reformat then trim it  
            Data.FormatTxt(OCR.textfile);

            //Trim the text we've formated so the form number is first
            Data.formatedList = Data.cleanList(Data.stringList); //Jonathan working on this function currently

            //Populate the excel spreadsheet with the approiate spreadsheet
            //Console.Write("Loading Excel Spreadsheet for State:{0}...",state);
            Data.LoadExcel();


            Console.WriteLine("Sorting items please do not touch this machine until process is complete...");
            AITask.BubbleSort(Data.CompareData());
            Console.WriteLine("Sorting finished");


            //After sorting the max 72 items, send click event to move screendown 
            Console.ReadLine();
        }
    }
}
