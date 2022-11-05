using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using RPDGenerator.Data;

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

            app.ActiveDocument.SaveAs2(getPath());
            app.ActiveDocument.Undo(countUndo);
        }
    }
}
