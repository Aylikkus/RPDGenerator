using System;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using RPDGenerator.Data;

namespace RPDGenerator.Interops
{
    public class WordGenerator : IDisposable
    {
        Application _app;
        Documents _documents;

        public void GenerateDocs(DocAttributes attrs, string pathToTemplate)
        {
            FileInfo templ = new FileInfo(pathToTemplate);
            Dictionary<string, string> tags = new Dictionary<string, string>() {
                        { "<FACULTY>",          attrs.Faculty                   },
                        { "<DEPARTMENT>",       attrs.Departament               },
                        { "<SPECIALIZATION>",   attrs.Specialization            },
                        { "<PROFILE>",          attrs.Profile                   },
                        { "<PROFILEABBR>",      attrs.ProfileAbbrevation        },
                        { "<EDUCATIONLEVEL>",   attrs.EducationLevel            },
                        { "<FORM>",             attrs.EducationType             },
                        { "<YEAROF>",           attrs.YearOfEntrance.ToString() },
                        { "<GRADUATIONLEVEL>",  attrs.GraduationLevel           },
                        { "<YEAROFENTRANCE>",   attrs.YearOfEntrance.ToString() },
            };
            
            _app = new Application();
            _app.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            _documents = _app.Documents;
            _documents.Open(templ.FullName, Type.Missing, true);

            bool processed = false;
            int count = attrs.Disciplines.Count();
            for (int i = 0; i < count; i++)
            {
                tags["<DISCIPLINE>"] = attrs.Disciplines[i].Name;
                WordProcess wp = new WordProcess(
                    tags, attrs.Disciplines[i], templ);

                // Сначала форматируем основные теги,
                // а потом изменяем только теги дисцплины
                if (!processed)
                {
                    wp.Process(_app);
                    processed = true;
                }
                else
                {
                    wp.ProcessDiscipline(_app);
                }
            };
        }

        public void Dispose()
        {
            _app.ActiveDocument.Close();
            _app.Quit();

            while (Marshal.ReleaseComObject(_app) > 0) { };
            while (Marshal.ReleaseComObject(_documents) > 0) { };

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}