using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BubbleSortClicker
{
    class AITask
    {
        /// <summary>
        /// carries out the necessary mouse clicks and moves to sort the data
        /// on the screen in a Bubble Sort Fashion
        /// </summary>
        /// <param name="index"></param>
        public static void BubbleSort(List<int> index)
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

            for (int i = 0; i < n - 1; i++)
            {

                for (int j = 0; j < n - 1; j++)
                {
                    Console.WriteLine(index[j]);
                    //j represents our index order in the Policy form
                    //we need to access the dictionaries key as the thing we check against
                    if (index[j + 1] > index[j])
                    {
                        //click the item we want to move
                        MouseEvent.LeftClick(1541, indexOneYPosition + (14 * index[j + 1]), 1);

                        //click on the move down button once
                        MouseEvent.LeftClick(2505, 195);
                    }
                }
                //Position scroll bar to capture next 72 items
                //MouseEvent.ScrollBarDown(Program.scrollbarPOs);
                //Program.scrollbarPOs += 98;
            }

        }
    }
}
