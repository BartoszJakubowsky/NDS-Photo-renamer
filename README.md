# NDS-First-deploy

It's my first simple app for automatically creating directories named on the basis of data from excel file and copying photos to this directories.

The application is programmed to find a excel's file with a specific name. It has exception if this file is already in use. Next it looks for specific 
string of characters and numbers. If app find one, it's going to get all the strings from current column and one before (alternatif strings for renaming photo). 
At the end it checks if all the array's length with strings from both column are equals.

Next step is creating directories named from excel's data. App looks if main folder already exist, creating one and then all the directories named from excel's data.

Next step is to find photos for copy and rename. In this step app looks for normal folder with photos and then for certain directories inside this folder. Same with 
zip file only that it unzipping this at first in new directory. 

Last step is copying photos with names of final directory.

Application has all the safety exceptions like "if the dir exist" and displays it on the console but i didn't know that relase version of app doesn't show
automatically console. Due to the correct operation of the application I decided about not changing anything.
