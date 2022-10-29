using System;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;

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

                    Object newFileName = Path.Combine(_fileInfo.DirectoryName, DateTime.Now.ToString("yyyy_MM_dd HH_mm_ss ") + _fileInfo.Name);
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
