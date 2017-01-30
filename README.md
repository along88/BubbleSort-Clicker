Automated Policy Form Window ordering bot
Alpha Version 1
README
Requirements
Capture2Text(OCR Software)
Screen Coordinate finder(I included a exe to Mofiki's Coordinate Finder)
Visual Studio 2015 or greater
Demo Directions
Open up the ExampleFormOrderingWindow.JPG , this photo is a screen capture of what a actual policy form ordering window looks like and the ocr will capture the photo the same it would the actual form window Note: Make sure you only capture the white part of the window for the OCR coordinates
Open the .sln explorer and make changes to the following lines of code
Program.Cs>Line 35> "Change these xy, xy2 coordinates to match your screen(Can use the included Mofiki's Coordinate Finder tool to get these coordinates of the winform window
Data.cs>Line 32> myTextReader filepath needs to be the path you have Capture2Text installed
Data.cs>Line106> This file path string needs to be where ever this repository is saved on your machine
OCR.cs>Line 25> file path should point to where ever your copy of Capture2Text is
make sure the ExampleFormOrderingWindow.JPG is unobstructed when the program begins and click start. The program will eventually fire off a series of mice clicks that would simulate a bubble sort algorithm for sorting the items in the picture(if they could be moved)

Directions
Before starting the program first open your desired Policy Forms ordering window.
Open the source code and ensure all files paths are reflecting YOUR directories
Make sure Capture2Text is installed
Maxmimize the screen of the form window in order to capture the largest possible area of the window
Now run the application
Note: this application has some very specific constraints in its current form
the windows policy form needs to be on the far right monitor and maxmimized
Screen resolution needs to be 1280 x 1024
Capture2Text must be installed
A Screen position tool was used to find and hard code the position of the windows
moving the windows will result in unforeseen errors unless the coordinates are properly changed in the code
Needs
architecture refactoring
Try/Catch blocks
Condition checks for what state we are working with
Excel workbook separated by states with unique forms and the desired orders for each state
Bubble sorting algorithm should only sort the first count of items and stop when this count is reached
it then needs to look back over the list until all items are in order
it then needs to scroll down to the next 75 items and start the OCR sorting process again
Additional features Coming:
A more user friendly interface so application can be deployed on multiple stations
What is this repository for?
This application will read the Policy Forms- ordering window and begin to sort the items in that window according to the excel sheet loaded in. In the background the code will perform the following task OCR the policy window format the text it reads so we get just the form number load the desired excel spreadsheet to be used as a reference for our desired order compare the two list to recognizing the order we have vs the desired order ** with this desired order in memory we send our mouse click event to the screen to carry out the sorting algorithm in real time
Who do I talk to?
Alfred or Jonathan
