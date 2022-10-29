using System;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using RPDGenerator.Data;

namespace RPDGenerator.Interops
{
    public class WordGenerator : IDisposable
    {
        Application _app;
        Documents _documents;
        Selection _selection;
        Find _find;

        public void Process(DocAttributes attrs, string pathToFile)
        {
            FileInfo fi = new FileInfo(pathToFile);
            Dictionary<string, string> tags = new Dictionary<string, string>() {
                        { "<FACULTY>",          attrs.Faculty                   },
                        { "<DEPARTMENT>",       attrs.Departament               },
                        { "<SPECIALIZATION>",   attrs.Specialization            },
                        { "<PROFILE>",          attrs.Profile                   },
                        { "<EDUCATIONLEVEL>",   attrs.EducationLevel            },
                        { "<FORM>",             attrs.EducationType             },
                        { "<YEAROF>",           attrs.YearOfEntrance.ToString() },
                        { "<GRADUATIONLEVEL>",  attrs.GraduationLevel           },
                        { "<YEAROFENTRANCE>",   attrs.YearOfEntrance.ToString() },
            };
            
            int count = attrs.Disciplines.Count();
            for (int i = 0; i < count; i++)
            {
                tags["<DISCIPLINE>"] = attrs.Disciplines[i].Name;

                _app = new Application();
                _documents = _app.Documents;
                _documents.Open(fi.FullName);
                _selection = _app.Selection;
                _find = _selection.Find;

                foreach (var tag in tags)
                {
                    var wrap = WdFindWrap.wdFindContinue;
                    var replace = WdReplace.wdReplaceAll;

                    _find.Text = tag.Key;
                    _find.Replacement.Text = tag.Value;

                    _find.Execute(
                        FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: Type.Missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: Type.Missing, 
                        Replace: replace);

                }

                StringBuilder fileName = new StringBuilder();
                fileName.Append("РПД_" + attrs.YearOfEntrance.ToString() + "_" 
                    + attrs.Specialization.Substring(0, 8) + "_" + attrs.EducationType[0] 
                    + "_" + attrs.Disciplines[i].Code + ".docx");

                string newFileName = Path.Combine(fi.DirectoryName, fileName.ToString());
                _app.ActiveDocument.SaveAs2(newFileName);

                // Очистка
                _app.ActiveDocument.Close();
                _app.Quit();

                while (Marshal.ReleaseComObject(_app) > 0) { };
                while (Marshal.ReleaseComObject(_documents) > 0) { };
            }
        }

        public void Dispose()
        {
            while (Marshal.ReleaseComObject(_app) > 0) { };
            while (Marshal.ReleaseComObject(_documents) > 0) { };
            while (Marshal.ReleaseComObject(_selection) > 0) { };
            while (Marshal.ReleaseComObject(_find) > 0) { };

            _app = null;
            _documents = null;
            _selection = null;
            _find = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

}