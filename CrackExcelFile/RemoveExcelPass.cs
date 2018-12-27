using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;

namespace CrackExcelFile
{
    public enum CrackOption
    {
        /// <summary>
        /// remove password complite and don't save password
        /// </summary>
        RemovePassComplite = 0,
        /// <summary>
        /// remove password and save it
        /// </summary>
        RemovePassAndKeep = 1,



    }

    public static class RemoveExcelPass
    {
        private static string ExtractPath { get; set; }
        private static string FileName { get; set; }
        private static string FileExtension { get; set; }
        private static string FileLocation { get; set; }
        private static string TempPath { get; set; }
        private static string WorkbookAddress { get; set; }
        private static string WorksheetAddress { get; set; }
      

        private static CrackOption CrackOptionInfo { get; set; }

        private static readonly List<Passwords> Password = new List<Passwords>();


        /// <summary>
        /// start to remove workbook password and all sheets pasword from target
        /// </summary>
        /// <param name="path">target file address</param>
        /// <param name="option"></param>
        public static async void OpenPass(string path, CrackOption option = CrackOption.RemovePassAndKeep)
        {
            //check path
            CheckAndSetPath(path);

            try
            {
                CrackOptionInfo = option;

                //open work book and remove pass
                var b1 = await SearchWorkbook();

                //open work sheets and remove pass
                var b2 = await SearchWorksheets();

                //make excel file and clean path
                while (true)
                {
                    if (b1 && b2)
                    {
                        CleanPath();
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }


        /// <summary>
        /// check some file in target path and create nececery path
        /// </summary>
        /// <param name="path"></param>
        private static void CheckAndSetPath(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine("File not exist");
                return;
            }
            FindFileName(path);
            TempPath = $"{FileLocation}temp";

            if (Directory.Exists(TempPath))
                Directory.Delete(TempPath, true);
            ExtractPath = $"{TempPath}";
            var zipPath = $"{TempPath}\\~{FileName}.zip";
            var directoryInfo = Directory.CreateDirectory(TempPath);
            WorkbookAddress = $"{TempPath}\\xl\\workbook.xml";
            WorksheetAddress = $"{TempPath}\\xl\\worksheets";
            directoryInfo.Attributes = FileAttributes.Hidden;

            //Todo:Hidden all file in directory
            //Todo:Read only all files and directory

            File.Copy(path, zipPath);

            //Extract zip file to temp path
            ZipFile.ExtractToDirectory(zipPath, ExtractPath);

        }


        /// <summary>
        /// Open workbook and remove password
        /// </summary>
        /// <returns></returns>
        private static Task<bool> SearchWorkbook()
        {
            var bol = Task.Run(() =>
                {
                    var doc = new XmlDocument();
                    doc.Load(WorkbookAddress);

                    var tag = doc.GetElementsByTagName("workbookProtection").Item(0);
                    if (tag?.Attributes != null)
                    {
                        if (CrackOptionInfo == CrackOption.RemovePassAndKeep)
                        {
                            var passValue = tag.Attributes.GetNamedItem("workbookHashValue").Value;
                            var passType = tag.Attributes.GetNamedItem("workbookAlgorithmName").Value;

                            var passwords = new Passwords
                            {
                                FileName = "",
                                PassType = passType,
                                PassValue = passValue,
                                Target = new TargetInfo
                                {
                                    FileAddress = $"{FileName}_New.{FileExtension}",
                                    TargetName = $"{FileName}",
                                    TargetType = "Workbook",
                                }
                            };
                            Password.Add(passwords);
                            tag.Attributes.GetNamedItem("lockStructure").Value = "";
                        }
                        else if (CrackOptionInfo == CrackOption.RemovePassComplite)
                        {
                            tag.RemoveAll();
                        }
                    }


                    doc.Save(WorkbookAddress);
                    return true;
                }

            );
            return bol;
        }


        /// <summary>
        /// search all worksheets and remove password
        /// </summary>
        /// <returns></returns>
        private static Task<bool> SearchWorksheets()
        {
            var workSheets = Directory.GetFiles(WorksheetAddress);
            var bol = Task.Run(() =>
            {
                foreach (var file in workSheets)
                {
                    var sheet = new XmlDocument();
                    sheet.Load(file);


                    var tag = sheet.GetElementsByTagName("sheetProtection").Item(0);
                    if (tag?.Attributes != null)
                    {
                        if (CrackOptionInfo == CrackOption.RemovePassAndKeep)
                        {
                            var passValue = tag.Attributes.GetNamedItem("hashValue").Value;
                            var passType = tag.Attributes.GetNamedItem("algorithmName").Value;
                            var fileName = string.Empty;
                            FindFileName(file, ref fileName);

                            var passwords = new Passwords
                            {
                                FileName = fileName,
                                PassType = passType,
                                PassValue = passValue,
                                Target = new TargetInfo
                                {
                                    FileAddress = $"{FileName}_New.{FileExtension}",
                                    TargetName = $"{FileName}",
                                    TargetType = "Worksheet",
                                }
                            };
                            Password.Add(passwords);

                            tag.Attributes.GetNamedItem("sheet").Value = "";
                        }
                        else if (CrackOptionInfo == CrackOption.RemovePassComplite)
                        {
                            tag.RemoveAll();
                        }
                    }


                    sheet.Save(file);
                }

                return true;
            }
            );
            return bol;
        }

        private static void WritePassFile()
        {
            Task.Run(() =>
            {
                try
                {
                    var newFilePath = $"{FileLocation}{FileName}_New.xlp";
                    File.WriteAllText(newFilePath, JsonConvert.SerializeObject(Password));
                    Console.WriteLine($"Password removed and new file is {FileName}_New.{FileExtension}");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

            });

        }
        /// <summary>
        /// make new excel file and clean path
        /// </summary>
        private static void CleanPath()
        {
            var zipPath2 = $"{FileLocation}xls.zip";
            var zipPath3 = $"{FileLocation}{FileName}_New.{FileExtension}";

            if (CrackOptionInfo == CrackOption.RemovePassAndKeep)
            {
                WritePassFile();
            }


            if (File.Exists(zipPath3))
            {
                File.Delete(zipPath3);
            }

            File.Delete(zipPath2);

            ZipFile.CreateFromDirectory(ExtractPath, zipPath2);

            File.Move(zipPath2, zipPath3);

            Directory.Delete(TempPath, true);

            File.Delete(zipPath2);
            //Console.WriteLine($"Password removed and new file is {FileName}_New.{FileExtension}");
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public static void ReadSavedPasswords(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine("File not exist");
                return;
            }

            var password = JsonConvert.DeserializeObject<List<Passwords>>(File.ReadAllText(path)).ToList();
            foreach (var item in password)
            {
                var ta= item.Target.FileAddress !=string.Empty ? $"address={item.Target.FileAddress}":string.Empty;
                var t1 = $"<{item.Target.TargetType} name={item.Target.TargetName} {ta}>\n";
                var t2 = $"     <file name> {item.FileName} </file name>\n";
                var t3 = $"     <password type> {item.PassType} </password type>\n";
                var t4 = $"     <password value> {item.PassValue} </password value>\n";
                var t5 = $"</{item.Target}>\n";
                Console.WriteLine(t1 + t2 + t3 + t4 + t5);
            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public static void ReturnPassword(string path)
        {
            //Todo:return password method
        }


        public static void SetNewPassword(string path, string newPass, string check)
        {
            //todo:set new pasword method
        }

        /// <summary>
        /// find file name , file location and file extension
        /// </summary>
        /// <param name="path">is target path</param>
        /// <param name="fileName"></param>
        private static void FindFileName(string path, ref string fileName)
        {
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));

            var chr = path.LastIndexOf("\\", StringComparison.Ordinal);
            var name = path.Substring(chr + 1, path.Length - chr - 1);
            var dot = name.LastIndexOf(".", StringComparison.Ordinal);
            fileName = name.Substring(0, dot);

        }

        /// <summary>
        /// find file name , file location and file extension
        /// </summary>
        /// <param name="path">is target path</param>
        private static void FindFileName(string path)
        {

            var chr = path.LastIndexOf("\\", StringComparison.Ordinal);
            FileLocation = path.Substring(0, chr + 1);
            var name = path.Substring(chr + 1, path.Length - chr - 1);

            var dot = name.LastIndexOf(".", StringComparison.Ordinal);
            FileName = name.Substring(0, dot);
            FileExtension = name.Substring(dot + 1, name.Length - dot - 1);


        }
    }
}
