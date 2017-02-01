using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


    public class AITask
    {
        private MouseEvent mouseEvent = new MouseEvent();

        /// <summary>
        /// BubbleSort algorithm that carries out the necessary mouse clicks and moves to sort the data
        /// on the screen in comparison to the order of the excel sheet
        /// </summary>
        /// <param name="index"></param>
        public void BubbleSort(List<int> index)
        {
            int indexOneYPosition = 62;
            //coming into this codeblock
            //we have one dictionary
            //we need to sort the policy form so that the index are in order
            //each step that is used to move one item in relation to another
            //needs to be translated into a click algorithm
            //we can either fire it off as the sort happens
            //or try to record these inputs
            int n = index.Capacity;

            //bool flag = true;

            for (int i = 0; i < n; i++)
            {

                for (int j = 0; j < n - 1; j++)
                {
                    Console.WriteLine(index[j]);
                    //j represents our index order in the Policy form
                    //we need to access the dictionaries key as the thing we check against
                    if (index[j + 1] > index[j])
                    {
                        //click the item we want to move
                        mouseEvent.LeftClick(1541, indexOneYPosition + (14 * index[j + 1]), 1);

                        //click on the move down button once
                        mouseEvent.LeftClick(2505, 195);
                    }
                }
                //Position scroll bar to capture next 72 items
                //MouseEvent.ScrollBarDown(Program.scrollbarPOs);
                //Program.scrollbarPOs += 98;
            }

        }
    }

