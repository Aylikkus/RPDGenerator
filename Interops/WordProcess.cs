using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using RPDGenerator.Data;
using System.Linq;
using System.Text;

namespace RPDGenerator.Interops
{
    class WordProcess
    {
        Dictionary<string, string> _tagsComm;
        Dictionary<string, string> _tagsDisc;
        Discipline _disc;

        string _tempPath;

        Document _template;
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

        void replaceCompetentions()
        {
            Range compsRange = _template.Bookmarks["ВсеКомпетенции"].Range;

            StringBuilder comps = new StringBuilder();
            foreach (var kv in _disc.Competentions)
            {
                comps.Append($"{kv.Key} – {kv.Value}\n");
            }

            compsRange.Text = comps.ToString();
        }

        string getPath()
        {
            string fileName = string.Join("_", "РПД", _tagsComm["<YEAROFENTRANCE>"], 
                _tagsComm["<SPECIALIZATION>"].Substring(0, 8), _tagsComm["<PROFILEABBR>"].ToLowerInvariant(),
                _tagsComm["<FORM>"][0], _disc.Code, _disc.Abbrevation);

            return Path.Combine(_template.Path, fileName + ".docx");
        }

        void formatCompTable()
        {
            Table compTable = _template.Bookmarks["Компетенции"].Range.Tables[1];

            int i = 3;
            foreach (var kv in _disc.Competentions)
            {
                compTable.Rows.Add();
                compTable.Cell(i, 1).Range.Text = (i - 2).ToString();
                compTable.Cell(i, 2).Range.Text = kv.Key;
                compTable.Cell(i, 3).Range.Text = kv.Value;

                i++;
            }
        }

        void pasteInCellWorkInfo(int row, int column, int sem, Table tb, WorkInfo wi)
        {
            tb.Cell(row, column).Range.Text = wi == null ? "-" : wi.HoursOnSemester(sem).ToString();
        }

        void formatTrudTable()
        {
            Table trudTable = _template.Bookmarks["Трудоёмкость"].Range.Tables[1];

            int semCount = _disc.Semester.Semesters.Count();
            for (int i = 3; i <= trudTable.Rows.Count; i++)
            {
                trudTable.Cell(i, 3).Split(1, semCount);
            }

            // Колонки с семестрами
            for (int j = 3; j <= trudTable.Columns.Count; j++)
            {
                int currSem = _disc.Semester.Semesters[j - 3];

                trudTable.Cell(3, j).Range.Text = currSem.ToString();

                // Лекции
                pasteInCellWorkInfo(5, j, currSem, trudTable, _disc.Lectures);

                // Лабы
                pasteInCellWorkInfo(6, j, currSem, trudTable, _disc.Laboratory);
                pasteInCellWorkInfo(7, j, currSem, trudTable, _disc.Laboratory);

                // Практики
                pasteInCellWorkInfo(8, j, currSem, trudTable, _disc.Practice);
                pasteInCellWorkInfo(9, j, currSem, trudTable, _disc.Practice);

                // Сам. Работы
                pasteInCellWorkInfo(10, j, currSem, trudTable, _disc.Independent);
            }
        }

        void formatDiscThemes()
        {
            Table themesTable = _template.Bookmarks["ТемыДисциплины"].Range.Tables[1];

            for (int i = 2; i <= themesTable.Rows.Count; i++)
            {
                themesTable.Cell(i, 2).Split(1, _disc.Competentions.Count);
            }

            int columnCount = 2;
            foreach(var comp in _disc.Competentions.Keys)
            {
                themesTable.Cell(2, columnCount).Range.Text = comp;

                columnCount++;
            }
        }

        public WordProcess(Dictionary<string, string> tagsComm, Dictionary<string, string> tagsDisc, 
            Discipline disc, string tempPath, Document template)
        {
            _tagsComm = tagsComm;
            _tagsDisc = tagsDisc;
            _disc = disc;
            _template = template;
            _tempPath = tempPath;
        }

        public void Process(Application app)
        {
            _selection = app.Selection;
            _find = _selection.Find;

            foreach (var tag in _tagsComm)
                findExecute(tag.Key, tag.Value);

            foreach (var tag in _tagsDisc)
                findExecute(tag.Key, tag.Value);

            replaceCompetentions();
            formatCompTable();
            formatTrudTable();
            formatDiscThemes();

            _template.SaveAs2(getPath());
            _template.Close();
        }
    }
}
