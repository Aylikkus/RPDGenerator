using System;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;
using RPDGenerator.Data;
using System.Threading;
using System.Linq;
using System.Diagnostics;

namespace RPDGenerator.WordGenerator
{
    public static class WordGenerator
    {
        public class WordHelper
        {
            private FileInfo _fileInfo;
            public WordHelper(string fileName)
            {
                if (File.Exists(fileName))
                {
                    _fileInfo = new FileInfo(fileName);
                }
                else
                {
                    throw new ArgumentException("File not found");
                }
            }

            public bool Process(DocAttributes da)
            {

                Dictionary<string,string> items = new Dictionary<string, string>() { 
                        { "<FACULTY>",da.Faculty},
                        { "<DEPARTMENT>", da.Departament},                        
                        { "<SPECIALIZATION>",da.Specialization},
                        { "<PROFILE>",da.Profile},
                        { "<EDUCATIONLEVEL>",da.EducationLevel},
                        { "<FORM>",da.EducationType},
                        { "<YEAROF>",da.YearOfEntrance.ToString()},
                        { "<GRADUATIONLEVEL>",da.GraduationLevel},
                        {"<YEAROFENTRANCE>",da.YearOfEntrance.ToString() } };
                int count = da.Disciplines.Count();
                Word.Application app = null; 
                try
                {
                    for (int i = 0; i < count; i++)
                    {
                        items["<DISCIPLINE>"] = da.Disciplines[i].Name;
                        app = new Word.Application();
                        Object file = _fileInfo.FullName;
                        Object missing = Type.Missing;
                        app.Documents.Open(file);
                        Word.Find find = app.Selection.Find;

                        foreach (var item in items)
                        {
                            find.Text = item.Key;
                            find.Replacement.Text = item.Value;
                            Object wrap = Word.WdFindWrap.wdFindContinue;
                            Object replace = Word.WdReplace.wdReplaceAll;
                            find.Execute(FindText: Type.Missing,
                                MatchCase: false,
                                MatchWholeWord: false,
                                MatchWildcards: false,
                                MatchSoundsLike: missing,
                                MatchAllWordForms: false,
                                Forward: true,
                                Wrap: wrap,
                                Format: false,
                                ReplaceWith: missing, Replace: replace);

                        }
                        string newFileName = Path.Combine(_fileInfo.DirectoryName, "РПД_" + da.YearOfEntrance.ToString() + "_" + da.Specialization.Substring(0,8) + "_" + da.EducationType[0] + "_" + da.Disciplines[i].Code) + ".docx";
                        app.ActiveDocument.SaveAs2(newFileName);
                        app.ActiveDocument.Close();
                        app.Quit();
                        Thread.Sleep(100);
                    }                   
                    return true;
                }
                catch (Exception ex) { Console.WriteLine(ex); }              
                return false;

            }
        }
    }
    
}