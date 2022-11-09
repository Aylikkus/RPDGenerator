using System;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using RPDGenerator.Data;
using System.Text;

namespace RPDGenerator.Interops
{
    public class WordGenerator : IDisposable
    {
        Application _app;
        Documents _documents;
        Document _template;

        string formatSemArray(int[] arr)
        {
            StringBuilder semBld = new StringBuilder(); 
            if (arr.Length > 1)
            {
                for (int i = 0; i < arr.Length; i++)
                {
                    semBld.Append(arr[i] + ((i == arr.Length - 1) ? "" : ", "));
                }
            }
            else
            {
                semBld.Append(arr[0]);
            }

            return semBld.ToString();
        }

        string formatAttestation(WorkInfo exam, WorkInfo credits, WorkInfo ratedCredits)
        {
            List<string> strs = new List<string>(3);

            if (credits != null) strs.Add("зачёт");
            if (ratedCredits != null) strs.Add("зачёт с оценкой");
            if (exam != null) strs.Add("экзамен");

            string att = string.Join(", ", strs.ToArray());
            return char.ToUpper(att[0]) + att.Substring(1);
        }

        int getTotalDisc(in Discipline disc)
        {
            WorkInfo[] works = new WorkInfo[]
            {
                disc.Lectures,
                disc.Practice,
                disc.Laboratory,
                disc.Independent,
                disc.Control,
            };

            int count = 0;

            foreach (var w in works)
            {
                if (w != null)
                    count += w.Total;
            }

            return count;
        }

        public void GenerateDocs(DocAttributes attrs, string pathToTemplate)
        {
            FileInfo templ = new FileInfo(pathToTemplate);
            Dictionary<string, string> tagsCommon = new Dictionary<string, string>() {
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

            int count = attrs.Disciplines.Count();
            for (int i = 0; i < count; i++)
            {
                _template = _documents.Open(templ.FullName);

                int totalh = getTotalDisc(attrs.Disciplines[i]);
                string totalle = attrs.Disciplines[i].Lectures == null ? "-" : attrs.Disciplines[i].Lectures.Total.ToString();
                string totalpr = attrs.Disciplines[i].Practice == null ? "-" : attrs.Disciplines[i].Practice.Total.ToString();
                string totalla = attrs.Disciplines[i].Laboratory == null ? "-" : attrs.Disciplines[i].Laboratory.Total.ToString();
                string totalin = attrs.Disciplines[i].Independent == null ? "-" : attrs.Disciplines[i].Independent.Total.ToString();

                Dictionary<string, string> tagsDiscipline = new Dictionary<string, string>() {
                    { "<DISCIPLINE>", attrs.Disciplines[i].Name },
                    { "<TOTALH>",  totalh.ToString() },
                    { "<LECTURESH>", totalle },
                    { "<PRACTICEH>", totalpr },
                    { "<LABORATORYH>", totalla },
                    { "<INDEPENDENTH>", totalin },
                    { "<COURSES>", formatSemArray(attrs.Disciplines[i].Semester.Courses) },
                    { "<SEMESTERS>", formatSemArray(attrs.Disciplines[i].Semester.Semesters) },
                    { "<TOTALCU>", (totalh / 36).ToString() },
                    { "<ACCREDITATION>", formatAttestation(attrs.Disciplines[i].Exam, attrs.Disciplines[i].Credits,
                        attrs.Disciplines[i].RatedCredits) },
                };

                
                WordProcess wp = new WordProcess(tagsCommon, tagsDiscipline, 
                    attrs.Disciplines[i], templ.FullName, _template);

                wp.Process(_app);
            };
        }

        public void Dispose()
        {
            _app.Quit();

            while (Marshal.ReleaseComObject(_app) > 0) { };
            while (Marshal.ReleaseComObject(_documents) > 0) { };
            while (Marshal.ReleaseComObject(_template) > 0) { };

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        
    }
}