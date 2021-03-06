﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
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

    /// <summary>
    /// Class for crack excel files
    /// </summary>
    public static class RemoveExcelPass
    {
        #region properties

        private static string ExtractPath { get; set; }

        private static string FileName { get; set; }

        private static string FileExtension { get; set; }

        private static string FileLocation { get; set; }

        private static string TempPath { get; set; }

        private static string WorkbookAddress { get; set; }

        private static string WorksheetAddress { get; set; }

        private static CrackOption CrackOptionInfo { get; set; }

        private static bool Buisy { get; set; }

        public static string FileText { get; set; }

        private static readonly List<Passwords> Password = new List<Passwords>();

        #endregion

        #region Message events

        /// <summary>
        /// Change messages
        /// </summary>
        /// <returns></returns>
        public delegate void Status(string result, string message);

        /// <summary>
        /// Change messages event
        /// </summary>
        public static event Status OnChange;

        public static string Message { get; set; }

        private static Action _messageChaned;

        private static void Change(string result, string message)
        {
            if (result != string.Empty && message != string.Empty && OnChange != null)
            {
                Message = $"{result}:{message}";

                if (_messageChaned != null) ShowMessage();
            }
            else
            {
                Message = "Nothing happend";
            }

        }

        private static void ShowMessage()
        {
            _messageChaned.Invoke();
        }

        /// <summary>
        /// Add action method for show message 
        /// for example:Console.WriteLine(RemoveExcelPass.Message);
        /// </summary>
        /// <param name="action"></param>
        public static void ShowMessagesAsync(Action action)
        {
            _messageChaned = action;
        }

        #endregion

        #region Main methods

        /// <summary>
        /// start to remove workbook password and all sheets pasword from target
        /// </summary>
        /// <param name="path">target file address</param>
        /// <param name="option"></param>
        public static async void OpenPass(string path, CrackOption option = CrackOption.RemovePassAndKeep)
        {
            Buisy = true;

            OnChange += Change;

            OnChange?.Invoke("Start", "Serching Password started");

            //check path
            OnChange?.Invoke("Cheking paths", "Cheking paths started");

            CheckAndSetPath(path);

            OnChange?.Invoke("Cheking paths", "Paths checked");

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
                        OnChange?.Invoke("Clean paths", "Clean Paths started");

                        CleanPath();

                        OnChange?.Invoke("Clean paths", "Paths cleaned");

                        break;

                    }
                }

                OnChange?.Invoke("Finished", "Passwords removed");

            }
            catch (Exception e)
            {
                OnChange?.Invoke("Error", e.Message);
            }
            finally
            {
                Buisy = false;
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public static async void ReturnPassword(string path)
        {
            Buisy = true;

            try
            {
                OnChange += Change;

                OnChange?.Invoke("Start", "Serching Password started");

                //check path
                OnChange?.Invoke("Cheking paths", "Cheking paths started");

                CheckAndSetPath(path);

                OnChange?.Invoke("Cheking paths", "Paths checked");

                //open work book and remove pass
                var b1 = await SearchWorkbook(returnPass: true);

                //open work sheets and remove pass
                var b2 = await SearchWorksheets(returnPass: true);

                //make excel file and clean path
                while (true)
                {
                    if (b1 && b2)
                    {
                        OnChange?.Invoke("Clean paths", "Clean Paths started");

                        CleanPath();

                        OnChange?.Invoke("Clean paths", "Paths cleaned");

                        break;
                    }
                }

                OnChange?.Invoke("Finished", "Passwords removed");
            }
            catch (Exception e)
            {
                OnChange?.Invoke("Error", e.Message);
            }
            finally
            {
                Buisy = false;
            }

        }

        /// <summary>
        /// Set new password to workbook and worksheets
        /// </summary>
        /// <param name="path"></param>
        /// <param name="workbookPass"></param>
        /// <param name="workbookConfirmPassword"></param>
        /// <param name="workSheetPass"></param>
        /// <param name="worksheetConfirmPassword"></param>
        public static async void SetNewPassword(string path, string workbookPass,
            string workbookConfirmPassword, string workSheetPass, string worksheetConfirmPassword)
        {
            Buisy = true;

            try
            {

                OnChange += Change;

                OnChange?.Invoke("Start", "Serching Password started");

                if (!workbookPass.Equals(workbookConfirmPassword))
                {
                    OnChange?.Invoke("Error", "Workbook password is difrent whith confirm password");
                    return;
                }

                if (!workSheetPass.Equals(worksheetConfirmPassword))
                {
                    OnChange?.Invoke("Error", "Worksheet password is difrent whith confirm password");
                    return;
                }

                if (workbookPass == string.Empty)
                {
                    OnChange?.Invoke("Error", "Workbook password can't be empty");
                    return;
                }

                if (workSheetPass == string.Empty)
                {
                    OnChange?.Invoke("Error", "Worksheet password can't be empty");
                    return;
                }

                OnChange?.Invoke("Cheking paths", "Cheking paths started");

                CheckAndSetPath(path);

                OnChange?.Invoke("Cheking paths", "Paths checked");

                //todo:make hash password

                var shaWorksheetPass = SHA512.Create(workSheetPass);

                var shaWorkbookPass = SHA512.Create(workbookPass);

                //open work book and remove pass
                var b1 = await SearchWorkbook(pass: shaWorkbookPass.Hash.ToString());

                //open work sheets and remove pass
                var b2 = await SearchWorksheets(pass: shaWorksheetPass.Hash.ToString());

                //make excel file and clean path
                while (true)
                {
                    if (b1 && b2)
                    {
                        OnChange?.Invoke("Clean paths", "Clean Paths started");

                        CleanPath();

                        OnChange?.Invoke("Clean paths", "Paths cleaned");

                        break;
                    }
                }
            }
            catch (Exception e)
            {
                OnChange?.Invoke("Error", e.Message);
            }
            finally
            {
                Buisy = false;
            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public static string ReadSavedPasswords(string path)
        {
            var txt = Task.Run(() =>
              {
                  while (Buisy)
                  {

                  }

                  OnChange += Change;

                  var str = string.Empty;

                  try
                  {

                      if (!File.Exists(path))
                      {
                          OnChange?.Invoke("Error", "File not exist");
                          return str;
                      }

                      var password = JsonConvert.DeserializeObject<List<Passwords>>(File.ReadAllText(path)).ToList();

                      foreach (var item in password)
                      {
                          var ta = item.Target.FileAddress != string.Empty ? $"address={item.Target.FileAddress}" : string.Empty;

                          var ct = item.Target.CreateTime != null ? $"createTime={item.Target.CreateTime}" : string.Empty;

                          var co = item.Target.TargetType == "workbook" ? item.Target.CrackOption.ToString() : string.Empty;

                          var t1 = $"<{item.Target.TargetType} name={item.Target.TargetName} {ta} {ct} {co}>\n";
                          var t2 = $"     <file name> {item.FileName} </file name>\n";
                          var t3 = $"     <password type> {item.PassType} </password type>\n";
                          var t4 = $"     <password value> {item.PassValue} </password value>\n";
                          var t5 = $"</{item.Target.TargetType}>\n";

                          str += t1 + t2 + t3 + t4 + t5;
                      }


                  }
                  catch (Exception e)
                  {
                      OnChange?.Invoke("Error", e.Message);

                  }
                  return str;

              });
            FileText = txt.Result;
            return txt.Result;
        }

        #endregion

        #region Actions method

        /// <summary>
        /// check some file in target path and create nececery path
        /// </summary>
        /// <param name="path"></param>
        private static void CheckAndSetPath(string path)
        {
            if (!File.Exists(path))
            {
                OnChange?.Invoke("Error", "File not exist");
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

            File.Copy(path, zipPath, true);

            //Extract zip file to temp path
            ZipFile.ExtractToDirectory(zipPath, ExtractPath);

        }

        /// <summary>
        /// make new excel file and clean path
        /// </summary>
        private static void CleanPath()
        {
            var zipPath = $"{TempPath}\\~{FileName}.zip";

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

            File.Delete(zipPath);

            File.Delete(zipPath2);

            ZipFile.CreateFromDirectory(ExtractPath, zipPath2);

            File.Move(zipPath2, zipPath3);

            Directory.Delete(TempPath, true);

            File.Delete(zipPath2);

        }

        /// <summary>
        /// Open workbook and remove password
        /// </summary>
        /// <returns></returns>
        private static Task<bool> SearchWorkbook(string pass = "", bool returnPass = false)
        {
            var bol = Task.Run(() =>
                {
                    var doc = new XmlDocument();

                    doc.Load(WorkbookAddress);

                    var tag = doc.GetElementsByTagName("workbookProtection").Item(0);

                    if (tag?.Attributes != null)
                    {
                        if (pass == string.Empty)
                        {
                            if (returnPass)
                            {
                                tag.Attributes.GetNamedItem("lockStructure").Value = "1";
                            }
                            else if (CrackOptionInfo == CrackOption.RemovePassAndKeep)
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
                                        CreateTime = DateTime.Now,
                                        CrackOption = CrackOptionInfo,
                                    }
                                };

                                Password.Add(passwords);

                                tag.Attributes.GetNamedItem("lockStructure").Value = "";
                            }
                            else if (CrackOptionInfo == CrackOption.RemovePassComplite)
                            {
                                //tag.RemoveAll();
                                tag.Attributes.GetNamedItem("lockStructure").Value = "";

                                tag.Attributes.GetNamedItem("workbookHashValue").Value = "";

                                tag.Attributes.GetNamedItem("workbookAlgorithmName").Value = "";
                            }

                        }
                        else
                        {

                            tag.Attributes.GetNamedItem("lockStructure").Value = "1";

                            tag.Attributes.GetNamedItem("workbookHashValue").Value = pass;

                            tag.Attributes.GetNamedItem("workbookAlgorithmName").Value = "SHA-512";
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
        private static Task<bool> SearchWorksheets(string pass = "", bool returnPass = false)
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
                        if (pass == string.Empty)
                        {
                            if (returnPass)
                            {
                                tag.Attributes.GetNamedItem("sheet").Value = "1";
                            }
                            else if (CrackOptionInfo == CrackOption.RemovePassAndKeep)
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
                                tag.Attributes.GetNamedItem("sheet").Value = "";

                                tag.Attributes.GetNamedItem("hashValue").Value = "";

                                tag.Attributes.GetNamedItem("algorithmName").Value = "";
                                //tag.RemoveAll();
                            }

                        }
                        else
                        {
                            tag.Attributes.GetNamedItem("sheet").Value = "1";

                            tag.Attributes.GetNamedItem("hashValue").Value = pass;

                            tag.Attributes.GetNamedItem("algorithmName").Value = "SHA-512";
                        }

                    }


                    sheet.Save(file);
                }

                return true;
            }
            );
            return bol;
        }

        /// <summary>
        /// 
        /// </summary>
        private static void WritePassFile()
        {
           try
            {
                var newFilePath = $"{FileLocation}{FileName}_New.xlp";

                var newFilePath2 = $"{TempPath}\\{FileName}.xlp";

                File.WriteAllText(newFilePath, JsonConvert.SerializeObject(Password));

                File.WriteAllText(newFilePath2, JsonConvert.SerializeObject(Password));

                OnChange?.Invoke("Password file:", $"New file is {FileName}_New.{FileExtension}");

            }
            catch (Exception e)
            {
                OnChange?.Invoke("Error", e.Message);
            }

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

        #endregion

    }
}
