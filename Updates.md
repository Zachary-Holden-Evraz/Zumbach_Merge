# Zumbach_Merge
This .exe file will merge each piece of Steel (with several .csv files) into one excel file.

Improvements made on 2/14/23:
  Fixed a bug that sometimes caused subfolders to not be deleted, stopping the program from continuing

Improvements made on 2/8/23:
  Fixed a bug that caused columns B-F on the excel file to be copied incorrectly.
  Changed the way that the files are sorted in the program.  

Improvements made on 1/20/23:
  Added an autosave system.  Every 100 pieces, the program will save the excel file and create a backup.

Improvements made on 12/16/22:
  Finally got it running faster.  It will not save without hitting the stop button, but every piece of steel should take less than 10 seconds.
  The program will now know when there are no more pieces to copy and display text accordingly.
  Starting and stopping the program still takes time proportional to the amount of information in the excel file.  I do not believe I can do anythig about that.

  The problem the program had before is that it had to save and reopen the excel file for each piece of steel.  This allowed it to save the file after each piece, which is safer, but took an incredible amount
    of time the more information was there.  If you are doing several hundred pieces at a time, it may be beneficial to stop and start it after about one or two hundred.  Just make sure to use the buttons for doing
    so and not just close the program. 

Improvements made on 12/9/22:
  An error was found and corrected.  The error caused columns B-F on the excel file to be copied incorrectly whenever the program was stopped and restarted later.

Improvements made on 11/28/22:
Implemented a start/stop button
  This will allow the user to stop the macro at any time without the risk of corrupting the main file
  This also improves the performance of the macro since it the start button puts it onto its own thread
  The start button must be pressed before the macro will begin

Quick run down of improvements to the macro as of 11/21/22: 

I sped it up a bit, though it is still not fast.  I am still working on improving the speed of the macro.

I also fixed a problem that I found today with it copying incorrect data if the macro is restarted.

The way it works now, the macro will consolidate every piece in the folder into its own sub-folder.  This speeds it up considerably at higher volumes of files.  Then it will work with each piece, deleting the sub-folders as it goes and saving after every piece.  

If the program ever needs to be stopped, make sure the excel file is working correctly afterwards, as I have noticed that it can get broken if the macro is stopped while it is saving, which takes longer the more data is in it.  

You should be able to download it from here as a OneDrive attachment; I tested that, and it worked ok.  Since we may not be able to use flash drives much longer as per an email from IT last week.  Just click the hyperlink, it should take you to a OneDrive URL, and click download.  I will also include a link below to a GitHub Repository below in case your computer doesn't like it for whatever reason.
