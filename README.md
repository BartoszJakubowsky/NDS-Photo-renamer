# NDS-First-deploy

It's my first simple C# app that's actually was helpfull and used in practice => it's for automatically creating directories named on the basis of data from excel file and copying photos to directories named same like photos.

The application is programmed to find a excel's file with a specific name. Next it looks for specific 
string of characters and numbers. If it find one, it's going to get all the strings from current column and one column before (with alternatif strings for renaming photo). Next step is creating directories named from excel data. 

Of course code has all the "check if " file exist, directory exist, if strings arrays lengths are the same, if excel is in use etc.

Next step is to find photos for copy and rename. In this step app looks for normal folder with photos and then for certain directories inside this folder. Source file may be in zip.

Last step is copying photos with names of final directory.
