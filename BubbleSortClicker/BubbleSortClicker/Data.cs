using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;


    public class Data
    {
        #region Properties
        private List<string> stringList = new List<string>(); //where we reference the OCR content
        private Application excelApp = new Application(); //Starting a new excelapp
        private Workbook ExcelWB; // the workbook the application is using
        private Worksheet ExcelWS; //the worksheet the application is using
        private List<string> ExcelSpreadSheet = new List<string>(); //where we will reference the worksheets spreadsheet
        private List<string> formatedList = new List<string>();

        public List<string> FormatedList
        {
            get
            {
                return formatedList;
            }
            set
            {
                formatedList = value;
            }
        }
        public List<string> StringList
        {
            get
            {
                return stringList;
            }
            set
            {
                stringList = value;
            }
        }
    #endregion

        /// <summary>
        /// Formats the copied text data and stores it in memory
        /// </summary>
        /// <param name="OCRtextFile"></param>
        public void FormatTxt(string OCRtextFile)
        {
            StreamReader myTextReader = new StreamReader(string.Format(@"C:\Users\along\Downloads\New Folder\Capture2Text\{0}", OCRtextFile));
            string line;
            //int counter = 0;
            while ((line = myTextReader.ReadLine()) != null)
            {
                if (line.Length == 0)
                    continue;
                else if (line[line.Length - 1] == ')')
                {
                    stringList.Add(line + '\n');
                    Console.WriteLine(line + '\n');
                }
                else
                {
                    stringList.Add(line);
                    Console.WriteLine(line);
                }
                
            }
            ////<summary>
            ///DEBUG CODE! 
            ///This Will open the text file for inspection/review after formating
            //Process run = new Process();
            //run.StartInfo.WorkingDirectory = @"C:\Users\along\Downloads\New Folder\Capture2Text\";
            //run.StartInfo.FileName = "FormatBlocks.txt";
            //run.Start();
            ////END OF DEBUG</summary>




        }

        /// <summary>
        /// Loads the given excel spreadsheet column range for reference
        /// </summary>
        /// <param name="columnOne"></param>
        /// <param name="columnTwo"></param>
        /// <returns></returns>
        public void LoadExcel()
        {
        ///<summary>
        ///TODO: promp user to select excel only files from a file dialog window instead of file string</summary>
            char columnOne = 'A';
            char columnTwo = 'A';
            string file = @"C:\Users\along\Source\Repos\NewRepo\BlockCreator\BlockCreator\Copy of AKOrder.xlsx";

            ExcelWB = (Workbook)excelApp.Workbooks.Open(file);
            ExcelWS = ExcelWB.ActiveSheet;
            Console.WriteLine("Getting form number order from Excel Worksheet....");

            //this col value should be grabbing the values of the Column A and B which
            //are the form name and form number respectively
            var col = ExcelWS.UsedRange.Columns[columnOne + ":" + columnTwo, Type.Missing];

            //add each row in order taking its index(1,2,3,4,etc..) and it's form number value
            //keep in mind the form number in excel may need to be split where the form
            //number ends and the " (1/14)" begins
            foreach (Range item in col.Cells)
            {

                ExcelSpreadSheet.Add(Convert.ToString(item.Value));
                //Console.WriteLine(item.Value);
            }

        }

        /// <summary>
        /// Looks at the policy's current order and returns desired order as a indexed list of intergers
        /// </summary>
        /// <param name="policyCurrentOrder"></param>
        /// <param name="spreadSheet"></param>
        /// <returns></returns>
        public List<int> CompareData()
        {
            List<int> PolicyFormList = new List<int>();

        //compare list here
        try
        {
            foreach (string item in formatedList)
            {

                for (int i = 0; i < formatedList.Count; i++)
                {
                    //if the form number matches the form number in spread sheet
                    //assign this item to the dictionary along with the spreadsheets index number
                    //represented by i
                    //Make sure the spreadsheet index's value we are evaluationg is the form number
                    //Note: "Item" value should be compared as "item.Substring(x,x)" 
                    //since we are only looking to compare the form numbers here

                    if (item.ToUpper() == ExcelSpreadSheet[i].ToUpper())
                    {
                        //add the item to the new dictionary and assign it's key
                        //value equal to the spreadsheets index, which should
                        //accurately be represented by i
                        PolicyFormList.Add(i);
                        Console.WriteLine(item + "Index:" + "{0}", i.ToString());
                        break;
                    }
                }
            }
        }
        catch (Exception ex)
        {

            Console.WriteLine(ex.Message);
        }
           Console.WriteLine(PolicyFormList.Count);
            return PolicyFormList;
        }

        public List<string> cleanList(List<string> ocrItems)
        {
            string temp;
            List<string> trimmed = new List<string>();
            List<string> final = new List<string>();

            foreach (string i in ocrItems)
            {
                int dash = i.IndexOf('-');

                if (dash <= 0)
                    continue;
                else if (i[dash + 1] == ' ')
                    temp = i.Substring(dash + 2, i.Length - (dash + 2));
                else if (i[dash + 1] != ' ')
                    temp = i.Substring(dash + 1, i.Length - (dash + 1));
                else
                    temp = i.Substring(dash, i.Length - dash);

                trimmed.Add(temp);
            }

            foreach (string i in trimmed)
            {
                int dash = i.IndexOf('-');
                string formNum;

                if (i.Substring(0, dash - 1) == "Manuscript" || i.Substring(0, dash) == "Manuscript")
                {
                    final.Add(i);
                }

                else if (i[dash - 1] == ' ')
                {
                    formNum = i.Substring(0, dash - 1);
                    final.Add(formNum);
                }
                else
                {
                    formNum = i.Substring(0, dash);
                    final.Add(formNum);
                }

            }
            return final;
        }
    }

