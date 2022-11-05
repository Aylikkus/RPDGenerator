using System;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using RPDGenerator.Data;
using System.CodeDom;
using System.Text;

namespace RPDGenerator.Interops
{
    public class WordGenerator : IDisposable
    {
        Application _app;
        Documents _documents;

        string formatSemArray(int[]a)
        {
            StringBuilder sb = new StringBuilder(); 
            if (a.Length > 1)
            {
                for (int i = 0; i < a.Length; i++)
                {
                    sb.Append(a[i] + ((i == a.Length) ? "" : ", "));
                }
            }
            else
            {
                sb.Append(a[0]);
            }

            return sb.ToString();
        }

        string formatAttestation(WorkInfo exam, WorkInfo credits, WorkInfo ratedCredits)
        {
            StringBuilder sb = new StringBuilder();
            if (credits != null) sb.AppendLine("зачёт ");
            if (ratedCredits != null) sb.AppendLine("зачёт с оценкой ");
            if (exam != null) sb.AppendLine("экзамен ");
            string str = sb.ToString().Replace('\n', ',');
            return char.ToUpper(str[0]) + str.Substring(1);
        }

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
                int totalh = attrs.Disciplines[i].Lectures.Total + attrs.Disciplines[i].Practice.Total +
                    attrs.Disciplines[i].Laboratory.Total + attrs.Disciplines[i].Independent.Total;
                tags["<DISCIPLINE>"] = attrs.Disciplines[i].Name;
                tags["<TOTALH>"] =  totalh.ToString();
                tags["<LECTRURESH>"] = attrs.Disciplines[i].Lectures.Total.ToString();
                tags["<PRACTICEH>"] = attrs.Disciplines[i].Practice.Total.ToString();
                tags["<LABORATORYH>"] = attrs.Disciplines[i].Laboratory.Total.ToString();
                tags["<INDEPENDENTH>"] = attrs.Disciplines[i].Independent.Total.ToString();
                tags["<COURSES>"] = formatSemArray(attrs.Disciplines[i].Semester.Courses);
                tags["<SEMESTERS>"] = formatSemArray(attrs.Disciplines[i].Semester.Semesters);
                tags["<TOTALCU>"] = (totalh / 36).ToString();
                tags["ACCREDITATION"] = formatAttestation(attrs.Disciplines[i].Exam, attrs.Disciplines[i].Credits,
                    attrs.Disciplines[i].RatedCredits);
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