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
            private DocAttributes da = new DocAttributes();
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

            public bool Process(Dictionary<string, string> items)
            {
                Word.Application app = null;
                try
                {
                    app = new Word.Application();
                    Object file = _fileInfo.FullName;
                    Object missing = Type.Missing;
                    app.Documents.Open(file);
                    foreach (var item in items)
                    {
                        Word.Find find = app.Selection.Find;
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

                    

                    Object newFileName = Path.Combine(_fileInfo.DirectoryName, "РПД_" + da.YearOfEntrance.ToString() + "_" + da.Specialization.Substring(0, 8) + "_" + da.EducationType.Substring(0) + "_+");
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();
                    app.Quit();
                    return true;
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
                finally
                {
                    if (app != null)
                    {
                        app.Quit();
                    }
                }
                return false;

            }
        }
    }
    
}



//int count = da.Disciplines.Count();
//for (int i = 0; i < count; i++)
//{
//    string disciplinecode = da.Disciplines[i].Code;


//    var items = new Dictionary<string, string>()
//                        { "<FACULTY>",da.Faculty},
//                        { "<DEPARTMENT>",da.Departament},
//                        { "<DISCIPLINE>", da.Disciplines[i].Name},
//                        { "<SPECIALIZATION>",da.Specialization},
//                        { "<PROFILE>",da.Profile}

//            var helper = new WordHelper("report_pattern.docx");
//            var items = new Dictionary<string, string>()
//            {
//                {"<NAME_OF_COMPANY>", tbnameofcompany.Text},
//                {"<NAME_OF_OBJECT>", tbnameofobject.Text},
//                {"<COUNT_OF_PERSONNEL>", Storage.CountOfPersonal.ToString()},
//                {"<COUNT_OF_PCS>", Storage.CountOfPCs.ToString()},
//                {"<COUNT_OF_SERVERS>", Storage.CountOfServers.ToString()},
//                {"<TYPE_OF_OBJECT>", lbltypeofobject_variable.Text},
//                {"<TYPEOFOBJECT_ACTIVITY>", lbltypeofobject_activity_variable.Text},
//                {"<TYPEOFOBJECT_PERSONNEL>", lbltypeofobject_personnel_variable.Text},
//                {"<TYPEOFOBJECT_PRODUCTION>", lbltypeofobject_production_variable.Text},
//                {"<TYPEOFOBJECT_SOFTWARE>", lbltypeofobject_software_variable.Text},
//                {"<TYPEOFOBJECT_LICENSED>", lbltypeofobject_licensed_variable.Text},
//                {"<PERSONNEL_SAFETY_INDEX>", Math.Round(Storage.personnel_safety / 10, 2).ToString()},
//                {"<INFRASTRUCTURE_SAFETY_INDEX>", Math.Round(Storage.infrastructure_safety / 8, 2).ToString()},
//                {"<APPLICATION_SAFETY_INDEX>", Math.Round(Storage.application_safety / 10, 2).ToString()},
//                {"<SECURITY_OF_OPERATIONS_INDEX>", Math.Round(Storage.security_of_operations / 7, 2).ToString()},
//                {"<RBQUEST3_1>", Answer((bool)Settings.Default["rbquest3_1_1"]) },
//                {"<RBQUEST3_2>", Answer((bool)Settings.Default["rbquest3_2_2"]) },
//                {"<RBQUEST3_3>", Answer((bool)Settings.Default["rbquest3_3_1"]) },
//                {"<RBQUEST3_4>", Answer((bool)Settings.Default["rbquest3_4_2"]) },
//                {"<RBQUEST3_5>", Answer((bool)Settings.Default["rbquest3_5_1"]) },
//                {"<RBQUEST3_6>", Answer((bool)Settings.Default["rbquest3_6_1"]) },
//                {"<RBQUEST3_7>", Answer((bool)Settings.Default["rbquest3_7_1"]) },
//                {"<RBQUEST3_8>", Answer((bool)Settings.Default["rbquest3_8_2"]) },
//                {"<RBQUEST3_9>", Answer((bool)Settings.Default["rbquest3_9_3"]) },
//                {"<RBQUEST4_1>", Answer((bool)Settings.Default["rbquest4_1_1"]) },
//                {"<RBQUEST4_2>", Answer((bool)Settings.Default["rbquest4_2_2"]) },
//                {"<RBQUEST4_3>", Answer((bool)Settings.Default["rbquest4_3_2"]) },
//                {"<RBQUEST4_4>", Answer((bool)Settings.Default["rbquest4_4_2"]) },
//                {"<RBQUEST4_5>", Answer((bool)Settings.Default["rbquest4_5_2"]) },
//                {"<RBQUEST4_6>", Answer((bool)Settings.Default["rbquest4_6_2"]) },
//                {"<RBQUEST4_7>", Answer((bool)Settings.Default["rbquest4_7_2"]) },
//                {"<RBQUEST4_8>", Answer((bool)Settings.Default["rbquest4_8_2"]) },
//                {"<RBQUEST5_1>", Answer((bool)Settings.Default["rbquest5_1_2"]) },
//                {"<RBQUEST5_2>", Answer((bool)Settings.Default["rbquest5_2_2"]) },
//                {"<RBQUEST5_3>", Answer((bool)Settings.Default["rbquest5_3_2"]) },
//                {"<RBQUEST5_4>", Answer((bool)Settings.Default["rbquest5_4_2"]) },
//                {"<RBQUEST5_5>", Answer((bool)Settings.Default["rbquest5_5_2"]) },
//                {"<RBQUEST5_6>", Answer((bool)Settings.Default["rbquest5_6_2"]) },
//                {"<RBQUEST5_7>", Answer((bool)Settings.Default["rbquest5_7_2"]) },
//                {"<RBQUEST5_8>", Answer((bool)Settings.Default["rbquest5_8_2"]) },
//                {"<RBQUEST5_9>", Answer((bool)Settings.Default["rbquest5_9_1"]) },
//                {"<RBQUEST5_10>", Answer((bool)Settings.Default["rbquest5_10_1"]) },
//                {"<RBQUEST6_1>", Answer((bool)Settings.Default["rbquest6_1_2"]) },
//                {"<RBQUEST6_2>", Answer((bool)Settings.Default["rbquest6_2_2"]) },
//                {"<RBQUEST6_3>", Answer((bool)Settings.Default["rbquest6_3_2"]) },
//                {"<RBQUEST6_4>", Answer((bool)Settings.Default["rbquest6_4_2"]) },
//                {"<RBQUEST6_5>", Answer((bool)Settings.Default["rbquest6_5_2"]) },
//                {"<RBQUEST6_6>", Answer((bool)Settings.Default["rbquest6_6_2"]) },
//                {"<RBQUEST6_7>", Answer((bool)Settings.Default["rbquest6_7_2"]) }
//            };
//{
//    if (MessageBox.Show("Создать отчёт в формате .docx? ", "Создание отчёта", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
//    {
//        helper.Process(items);
//        string filename = DateTime.Now.ToString("yyyy_MM_dd HH_mm_ss ") + "report_pattern.docx";
//        try
//        {
//            Process.Start(@"C:\Users\Honor\source\repos\Kursovik\Kursovik\bin\Debug\" + filename);
//        }
//        catch
//        {
//            MessageBox.Show("Файл " + filename + " успешно сохранён, открыть его не удалось в связи с неизвестной ошибкой.", "Дополнительная информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
//        }
//    }
//}