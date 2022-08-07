
using System;
//dla ścieżki
using System.IO;
//dla czytania i pobierania danych z excel
using Microsoft.Office.Interop.Excel;
//dla wyrażeń regularnych
using System.Text.RegularExpressions;



class ProgramNDS
{
    static void Main(string[] args)
    {

        string currentDir = Directory.GetCurrentDirectory();

        //check if excel file exist
        string excelFile = doesExcelFileExist(currentDir);

        if (excelFile == "")
        {
            Console.WriteLine("File not found");
            Console.ReadLine();

        }

        //read excel file
        string[,] mufaArray = readExcelFile(excelFile);


        //create directories for files
        string finalDirectory = createDirectories(mufaArray, currentDir);

        //move and rename photos
        string[] finalPhotoPath = validPhotos(mufaArray, finalDirectory, currentDir);
        if (finalPhotoPath[0] == "null")
        {
            Console.WriteLine("Photos not found");
            Console.ReadLine();
        }

        MoveRenameMufy(mufaArray, finalPhotoPath, Directory.GetDirectories(finalDirectory));
        Console.ReadLine();


    }

    public static string doesExcelFileExist(string directory)
    {

        string[] dirNames = Directory.GetFiles(directory);

        //"Zał\\. 3.*\\.xlsx" //Zał\..*3.*\.xlsx
        Regex regexWithSpace = new Regex("Zał\\. 3.*\\.xlsx", RegexOptions.IgnoreCase);
        Regex regexWithoutSpace = new Regex("Zał\\..*3.*\\.xlsx", RegexOptions.IgnoreCase);

        //flag
        bool doesFileExist = false;

        //ścieżka z plikiem
        string filePathName = "";

        foreach (string dirName in dirNames)
        {

            //check with space
            regexWithSpace.IsMatch(dirName);
            if (regexWithSpace.IsMatch(dirName) == true)
            {
                doesFileExist = true;
                filePathName = dirName;
                break;
            }

            //check without space
            regexWithoutSpace.IsMatch(dirName);
            if (regexWithoutSpace.IsMatch(dirName) == true)
            {
                doesFileExist = true;
                filePathName = dirName;
                break;
            }
        }

        if (doesFileExist == true)
        {
            Console.WriteLine(filePathName);

            return filePathName;
        }
        else
        {
            return filePathName = "";
        }

    }


    //dont forget -- using Microsoft.Office.Interop.Excel;
    public static string[,] readExcelFile(string filePath)
    {

        //Create COM Objects.
        Application excelApp = new Application();

        string[,] mufaArray = new string[1, 1];


        if (excelApp == null)
        {
            Console.WriteLine("Excel is not installed!!");

            //return null
            return mufaArray;
        }
        try
        {
            excelApp.Workbooks.Open(filePath);
        }
        catch (Exception)
        {
            Console.WriteLine("Excel file is being used!\n");
            throw;
            
        }
        Workbook excelBook = excelApp.Workbooks.Open(filePath);
        _Worksheet excelSheet = (_Worksheet)excelBook.Sheets[1];
        Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;

        Regex mufaName = new Regex("ms[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+", RegexOptions.IgnoreCase);

        List<string> mufaList = new List<string>();
        List<string> nazwaMufaList = new List<string>();


        for (int i = 3; i <= rows; i++)
        {
            //create new line
            Console.Write("\r\n");
            for (int j = 1; j <= cols; j++)
            {

                //write the console
                if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                {
                    if (mufaName.IsMatch(excelRange.Cells[i, j].Value2.ToString()) && excelRange.Cells[i, j - 1] != null && excelRange.Cells[i, j - 1].Value2 != null)
                    {
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + " " + excelRange.Cells[i, j - 1].Value2.ToString() + "\t");

                        mufaList.Add(excelRange.Cells[i, j].Value2.ToString());
                        nazwaMufaList.Add(excelRange.Cells[i, j - 1].Value2.ToString());


                    }
                }
            }
        }
        //after reading, relaase the excel project
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

        int mufaListLength = mufaList.ToArray().Length;
        int nazwaMufaListLength = nazwaMufaList.ToArray().Length;

        if (mufaListLength != nazwaMufaListLength)
        {
            throw new Exception("Ilość muf się nie zgadza ze sobą\n\n");
        }

        mufaArray = new string[2, nazwaMufaListLength];

        //poprawić, ogarnąć podwójne tablice

        for (int k = 0; k < mufaListLength; k++)
        {
            mufaArray[1, k] = mufaList[k];
            mufaArray[0, k] = nazwaMufaList[k];
        }

        return mufaArray;
    }

    public static string createDirectories(string[,] mufaArray, string excelFilePath)
    {
        string[] dirs = Directory.GetDirectories(excelFilePath, "*", SearchOption.TopDirectoryOnly);
        Regex mainDirForMufy = new Regex("5\\.2\\.[a-zA-Z]+_[a-zA-Z]+_[a-zA-Z]+_[a-zA-Z]+_[a-zA-Z]+_[a-zA-Z]+", RegexOptions.IgnoreCase);

        string finalDirectory = Path.Combine(excelFilePath, "5.2.Zdjecia_Przelacznic_Glownych_Punktow_Dostepowych_muf");

        for (int i = 0; i <= dirs.Length; i++)
        {
            if (i < dirs.Length && mainDirForMufy.IsMatch(dirs[i]))
            {
                Console.WriteLine("\nfolder \"Folder 5.2\" już istnieje\n");
                break;
            }
            if (i == dirs.Length)
            {
                //do zmiany na docelowy folder
                Directory.CreateDirectory(finalDirectory);
            }
        }

        //finalna ścieżka do zmiany



        for (int i = 0; i < mufaArray.GetLength(1); i++)
        {
            if (Directory.Exists(Path.Combine(finalDirectory, mufaArray[1, i])) == true)
            {
                Console.WriteLine($"\nFolder {mufaArray[1, i].ToUpper()} już istnieje\n");
                continue;
            }
            else
            {
                Directory.CreateDirectory(Path.Combine(finalDirectory, mufaArray[1, i].ToUpper()));
            }
        }

        return finalDirectory;
    }

    public static string[] validPhotos(string[,] mufaArray, string finalPath, string mainPath)
    {

        //get array of all files and folders in directory

        string[] fileArray = Directory.GetFileSystemEntries(mainPath);
        string[] lookForPhotoDirectory;
        string[] photosDirNames;

        List<string> tempList = new List<string>();


        Regex regex = new Regex("foto-gis+", RegexOptions.IgnoreCase);
        Regex mufy = new Regex("mufy", RegexOptions.IgnoreCase);
        Regex po = new Regex("po", RegexOptions.IgnoreCase);

        //unzip file
        string unzipDirectory = @$"{mainPath}\UNZIPPED GIS PHOTOS";



        for (int fileInMainDir = 0; fileInMainDir < fileArray.Length; fileInMainDir++)
        {
            //use unzipped folder first


            if (regex.IsMatch(fileArray[fileInMainDir]) == true && Path.GetExtension(fileArray[fileInMainDir]).Equals(".zip") == false)
            {


                lookForPhotoDirectory = Directory.GetDirectories(fileArray[fileInMainDir], "*", SearchOption.AllDirectories);
                


                for (int dir = 0; dir < lookForPhotoDirectory.Length; dir++)
                {


                    photosDirNames = photoDirectoriesNames(dir, lookForPhotoDirectory);
                    if (photosDirNames[0] != "null")
                    {
                        return photosDirNames;
                    }


                }

            }


            else if (Directory.Exists(unzipDirectory) == true)
            {
                lookForPhotoDirectory = Directory.GetDirectories(unzipDirectory, "*", SearchOption.AllDirectories);

                for (int dir = 0; dir < lookForPhotoDirectory.Length; dir++)
                {

                        photosDirNames = photoDirectoriesNames(dir, lookForPhotoDirectory);
                        if (photosDirNames[0] != "null")
                        {
                            return photosDirNames;
                        }

                }
            }

            else if (Path.GetExtension(fileArray[fileInMainDir]).Equals(".zip"))
            {

                try
                {
                    if (Directory.Exists(unzipDirectory) == false)
                    {
                        Directory.CreateDirectory(unzipDirectory);

                    }
                    System.IO.Compression.ZipFile.ExtractToDirectory(fileArray[fileInMainDir], unzipDirectory);

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error happend during unzipping file: ");
                    throw new Exception(ex.Message);
                }

                lookForPhotoDirectory = Directory.GetDirectories(unzipDirectory, "*", SearchOption.AllDirectories);

                    for (int dir = 0; dir < lookForPhotoDirectory.Length; dir++)
                    {
                        if (regex.IsMatch(lookForPhotoDirectory[dir]))
                        {
                            photosDirNames = photoDirectoriesNames(dir, lookForPhotoDirectory);
                            if (photosDirNames[0] != "null")
                            {
                                return photosDirNames;
                            }
                        }

                        //if regex didnt find anything
                        photosDirNames = photoDirectoriesNames(dir, lookForPhotoDirectory);
                        if (photosDirNames[0] != "null")
                        {
                            return photosDirNames;
                        }
                    }
            }


        }
        Console.WriteLine("Files not found");
        photosDirNames = new string[1];
        photosDirNames[0] = "null";
        return photosDirNames;
    }




    public static void MoveRenameMufy(string[,] mufaArray, string[] photoPathArray, string[] finalDirPath)
    {
        //new Regex("ms.*[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+[0-9]+", RegexOptions.IgnoreCase
        Regex regexMufaNumber;
        Regex regexPhotoName;
        
        for (int i = 0; i < mufaArray.GetLength(1); i++)
        {
          foreach (string photo in photoPathArray)
          {
                string ext = Path.GetExtension(photo);
                regexMufaNumber = new Regex(mufaArray[1, i], RegexOptions.IgnoreCase);
                string mufaLitera = mufaArray[0, i];

                regexPhotoName = new Regex($"mufa_{mufaLitera}_", RegexOptions.IgnoreCase);


                if (regexPhotoName.IsMatch(photo) ^ regexMufaNumber.IsMatch(photo))
                {

                    for (int j = 0; j < finalDirPath.Length; j++)
                {


                        if (regexMufaNumber.IsMatch(finalDirPath[j]))
                        {

                            if (File.Exists($@"{finalDirPath[j]}\{mufaArray[1, i]} {ext}"))
                            {
                                for (int k = 0; k < Directory.GetFiles($@"{finalDirPath[j]}").Length; k++)
                                {
                                    if (File.Exists($@"{finalDirPath[j]}\{mufaArray[1, i]} ({k+1}){ext}") == false)
                                    {
                                        try
                                        {
                                            File.Copy(photo, $@"{finalDirPath[j]}\{mufaArray[1, i]} ({k+1}){ext}");
                                            break;
                                        }
                                        catch (Exception)
                                        {
                                            Console.WriteLine("Wystąpił błąd z" + photo + "\n");
                                            Console.WriteLine($@"{finalDirPath[j]}\{mufaArray[1, i]} ({k + 1}){ext}");
                                            throw;

                                        }
                                    } 
                                    
                                }
                            }
                            else
                            {
                                try
                                {
                                    File.Copy(photo, $@"{finalDirPath[j]}\{mufaArray[1, i]} {ext}");
                                    break;
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Wystąpił błąd z" + photo);


                                    throw;
                                }

                            }


                            //Console.WriteLine($@"File: {finalMufaPath}\{mufaNumer}({i}) already exists");

                        }
                        

                    }
                }
            }
        }
        if (photoPathArray.Length == 0)
        {
            Console.WriteLine($"W folerze {photoPathArray} nic nie zostało");
        }
        else
        {
            Console.WriteLine($"W folerze {photoPathArray} zostało {photoPathArray.Length} elementów");
        }

    }


    //look lookForPhotoDirectory is array of main folder where program is located
    public static string[] photoDirectoriesNames(int dir, string[] lookForPhotoDirectory)
    {
        string[] photosDirNames;
        List<string> tempList = new List<string>();

        Regex mufy = new Regex("mufy", RegexOptions.IgnoreCase);
        Regex po = new Regex("po", RegexOptions.IgnoreCase);


            //look for mufy
            if (mufy.IsMatch(lookForPhotoDirectory[dir]))
            {
               if (po.IsMatch(lookForPhotoDirectory[dir]))
               {

                //look for photos
                if (Directory.GetFiles(lookForPhotoDirectory[dir]).Length > 0)

                    {
                        foreach (string photo in Directory.GetFiles(lookForPhotoDirectory[dir]))
                        {
                            if (Path.GetExtension(photo).Equals(".jpg") ^ Path.GetExtension(photo).Equals(".png") ^ Path.GetExtension(photo).Equals(".gif") == true)
                            {

                                tempList.Add(photo);


                            }
                        }
                        photosDirNames = tempList.ToArray();
                        return photosDirNames;
                    }
               }
            }
            else if (po.IsMatch(lookForPhotoDirectory[dir]))
            {

            //look for photos

                if (Directory.GetFiles(lookForPhotoDirectory[dir]).Length > 0)
                {
                    foreach (string photo in Directory.GetFiles(lookForPhotoDirectory[dir]))
                    {
                        if (Path.GetExtension(photo).Equals(".jpg") ^ Path.GetExtension(photo).Equals(".png") ^ Path.GetExtension(photo).Equals(".png") == true)
                        {

                            tempList.Add(photo);

                        }
                    }
                    photosDirNames = tempList.ToArray();
                    return photosDirNames;
                }

            }
            else if (mufy.IsMatch(lookForPhotoDirectory[dir]))
            {

                Console.WriteLine("WARNING\n");
                Console.WriteLine("Folder \"Po\" not found");

            //look for photos                              

                if (Directory.GetFiles(lookForPhotoDirectory[dir]).Length > 0)
                {
                    foreach (string photo in Directory.GetFiles(lookForPhotoDirectory[dir]))
                    {

                        if (Path.GetExtension(photo).Equals(".jpg") ^ Path.GetExtension(photo).Equals(".png") ^ Path.GetExtension(photo).Equals(".png") == true)
                        {

                            tempList.Add(photo);

                        }

                    }
                    photosDirNames = tempList.ToArray();
                    return photosDirNames;
                }
            }

        photosDirNames = new string[1];
        photosDirNames[0] = "null";
        return photosDirNames;
    }
   
}



    









