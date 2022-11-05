using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using RPDGenerator.Data;
using System.Linq;

namespace RPDGenerator.Interops
{
    class WordProcess
    {
        Dictionary<string, string> _tagsComm;
        Dictionary<string, string> _tagsDisc;
        Discipline _disc;
        FileInfo _template;

        Selection _selection;
        Find _find;

        bool findExecute(string text, string repl)
        {
            _find.Text = text;
            _find.Replacement.Text = repl;

            return _find.Execute(
                FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: Type.Missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: WdFindWrap.wdFindContinue,
                Format: false,
                ReplaceWith: Type.Missing,
                Replace: WdReplace.wdReplaceAll);
        }

        string getPath()
        {
            string fileName = string.Join("_", "РПД", _tagsComm["<YEAROFENTRANCE>"], 
                _tagsComm["<SPECIALIZATION>"].Substring(0, 8), _tagsComm["<PROFILEABBR>"].ToLowerInvariant(),
                _tagsComm["<FORM>"][0], _disc.Code, _disc.Abbrevation);

            return Path.Combine(_template.DirectoryName, fileName + ".docx");
        }

        void formatTrudTable(Application app, ref int counter)
        {
            Table trudTable = app.ActiveDocument.Bookmarks["Трудоёмкость"].Range.Tables[1];
            Cell semCell = trudTable.Cell(3, 3);

            int semCount = _disc.Semester.Semesters.Count();
            for (int i = 3; i <= trudTable.Rows.Count; i++)
            {
                trudTable.Cell(i, 3).Split(1, semCount);
                counter++;
            }

            // Колонки с семестрами
            for (int j = 3; j <= trudTable.Columns.Count; j++)
            {
                int currSem = _disc.Semester.Semesters[j - 3];

                trudTable.Cell(3, j).Range.Text = currSem.ToString();
                counter++;

                // Лекции
                if (_disc.Lectures != null)
                {
                    trudTable.Cell(5, j).Range.Text = _disc.Lectures.HoursOnSemester(currSem).ToString();
                    counter++;
                }
                
                // Лабы
                if (_disc.Laboratory != null)
                {
                    trudTable.Cell(6, j).Range.Text = _disc.Laboratory.HoursOnSemester(currSem).ToString();
                    trudTable.Cell(7, j).Range.Text = _disc.Laboratory.HoursOnSemester(currSem).ToString();
                    counter += 2;
                }

                // Практики
                if (_disc.Practice != null)
                {
                    trudTable.Cell(8, j).Range.Text = _disc.Practice.HoursOnSemester(currSem).ToString();
                    trudTable.Cell(9, j).Range.Text = _disc.Practice.HoursOnSemester(currSem).ToString();
                    counter += 2;
                }

                // Сам. Работы
                if (_disc.Independent != null)
                {
                    trudTable.Cell(10, j).Range.Text = _disc.Independent.HoursOnSemester(currSem).ToString();
                    counter++;
                }
            }
        }

        public WordProcess(Dictionary<string, string> tagsComm, Dictionary<string, string> tagsDisc, 
            Discipline disc, FileInfo template)
        {
            _tagsComm = tagsComm;
            _tagsDisc = tagsDisc;
            _disc = disc;
            _template = template;
        }

        public void Process(Application app)
        {
            _selection = app.Selection;
            _find = _selection.Find;

            foreach (var tag in _tagsComm)
                findExecute(tag.Key, tag.Value);

            ProcessDiscipline(app);
        }

        public void ProcessDiscipline(Application app)
        {
            int countUndo = 0;
            _selection = app.Selection;
            _find = _selection.Find;

            foreach(var tag in _tagsDisc)
            {
                countUndo += findExecute(tag.Key, tag.Value) ? 1 : 0;
            }

            formatTrudTable(app, ref countUndo);

            app.ActiveDocument.SaveAs2(getPath());
            app.ActiveDocument.Undo(countUndo);
        }
    }
}
