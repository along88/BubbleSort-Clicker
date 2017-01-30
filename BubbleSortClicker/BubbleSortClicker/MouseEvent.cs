using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Threading;


    class MouseEvent
    {
        #region Mouse Event Setup
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")] // establish the mouse cursor position
        private static extern bool SetCursorPos(int x, int y); //extern event to take in mouse pos parameters 
        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwflgs, int dx, int dy, int cButtons, int dwExtraInfo); //mouse click parameters
        private const int MOUSEEVENTF_LEFTDOWN = 0X02; //register mouse left click down
        private const int MOUSEEVENTF_LEFTUP = 0X04; //register mouse left click up(simulates releasing the click)
        #endregion

        /// <summary>
        /// Sends a mouse left click to the desired screen coordinates
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public static void LeftClick(int x, int y)
        {

            SetCursorPos(x, y); //ensure the mouse moves to this position before clicking
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0); //send the mouse click at those same consistent coordinates
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);//release the mouse click at those same consistent coordinates
            Thread.Sleep(1000); //wait one second before exiting to allow additional mouse clicks or other operations


        }
        /// <summary>
        /// Sends X amount of mouse left clicks to the desired screen coordinates
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        public static void LeftClick(int x, int y, int numberOfClicks)
        {

            SetCursorPos(x, y);
            for (int i = 0; i < numberOfClicks; i++)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
                Thread.Sleep(1000);
            }

        }

        /// <summary>
        /// this function will send the click event based on the comparison bubble sort data return
        /// after reviewing the OCR compared to the Excel
        /// </summary>
        public static void clickAlgorithm()
        {
            int mismatchIndex = 4;

            //step one
            //we will click the item we wish to move
            //the item we first click will be determined by
            //the item that first returned out of order
            //for now we will test this by feeding in the coordinates for
            //the item at index 5
            //we can move the click to the position but using the knowledge that the
            //general distance from one index to the next on screen is about 14 pixels apart
            //using the position of index one as a constant we will move our position down
            //by 14 pixels for every index until we hit the 5th one
            int indexOneYPosition = 62;
            LeftClick(1541, indexOneYPosition + (14 * mismatchIndex), 1);
            //Move Item up(will write helper method to run script as many times as needed based on
            //how many units it needs to be moved up in relation to the index its suppose to be)
            //Click(2500,   107);

            //Move Item Down for shits and giggles
            //Click(2505, 195);
            LeftClick(2505, 195, 5);

            LeftClick(1541, indexOneYPosition + (14 * 63), 1);
            LeftClick(2503, 139);
            LeftClick(2503, 101, 10);



        }

        /// <summary>
        /// Sends a left click hold and drag event for positioning scroll bars
        /// xy are scrollbars starting coordinates, y2 are ending positions
        /// </summary>
        /// <param name="scrollbar"></param>
        /// <param name="y"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        public static void ScrollBarDown (int lastPOS)
        {
                
                lastPOS =+ 98;
                int[] scrollbar = new int[2];
                scrollbar[0] = 2443;
                scrollbar[1] = lastPOS;
                int y2 = scrollbar[1] + (98);
                SetCursorPos(scrollbar[0], scrollbar[1]); //ensure the mouse moves to this position before clicking
                mouse_event(MOUSEEVENTF_LEFTDOWN, scrollbar[0], scrollbar[1], 0, 0); //send the mouse click at those same consistent coordinates
                SetCursorPos(scrollbar[0], y2);
                mouse_event(MOUSEEVENTF_LEFTUP, scrollbar[0], y2, 0, 0);
                Thread.Sleep(1000);
                
            
            
        }
    }

//public static void insertBlock(string textFile)
//Old Dorment code that was originally used for automating creating blocks in ABBYY FLEXILAYOUT
//Keeping code commented out should there be valuable insight or use it may one day provide
//{
//    StreamReader readBlocksList = new StreamReader(string.Format("C:\\Users\\along\\Downloads\\New Folder\\Capture2Text\\{0}", textFile));

//    string line;
//    while ((line = readBlocksList.ReadLine()) != null)
//    {
//        //for each line we read, if that line is equal to the work Block
//        //we will then execute our moduel 
//        Console.WriteLine(line.TrimEnd() + "Debug No Trim");
//        if (line.TrimEnd('k') == "Block")
//        {
//            //Entry point for executing Pointer Logic
//            line = line.TrimEnd('k');
//            Console.WriteLine(line + "Debug Trim");
//        }
//        readBlocksList.ReadLine();
//        }
//    }




//insert automate mouse click method here