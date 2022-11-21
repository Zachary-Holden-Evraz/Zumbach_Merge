# Zumbach_Merge
This .exe file will merge each piece of Steel (with several .csv files) into one excel file.

Quick run down of improvements to the macro as of 11/21/22: 

I sped it up a bit, though it is still not fast.  I am still working on improving the speed of the macro.

I also fixed a problem that I found today with it copying incorrect data if the macro is restarted.

The way it works now, the macro will consolidate every piece in the folder into its own sub-folder.  This speeds it up considerably at higher volumes of files.  Then it will work with each piece, deleting the sub-folders as it goes and saving after every piece.  

If the program ever needs to be stopped, make sure the excel file is working correctly afterwards, as I have noticed that it can get broken if the macro is stopped while it is saving, which takes longer the more data is in it.  

You should be able to download it from here as a OneDrive attachment; I tested that, and it worked ok.  Since we may not be able to use flash drives much longer as per an email from IT last week.  Just click the hyperlink, it should take you to a OneDrive URL, and click download.  I will also include a link below to a GitHub Repository below in case your computer doesn't like it for whatever reason.
