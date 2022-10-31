using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using System.Text;
using RPDGenerator.Data;

namespace RPDGenerator.Interops
{
    class WordProcess
    {
        Dictionary<string, string> _tags;
        Discipline _disc;
        FileInfo _template;

        Selection _selection;
        Find _find;

        void findExecute(string text, string repl)
        {
            _find.Text = text;
            _find.Replacement.Text = repl;

            _find.Execute(
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
            string fileName = string.Join("_", "РПД", _tags["<YEAROFENTRANCE>"], 
                _tags["<SPECIALIZATION>"].Substring(0, 8), _tags["<PROFILEABBR>"].ToLowerInvariant(),
                _tags["<FORM>"][0], _disc.Code, _disc.Abbrevation);

            return Path.Combine(_template.DirectoryName, fileName + ".docx");
        }

        public WordProcess(Dictionary<string, string> tags, Discipline disc, FileInfo template)
        {
            _tags = tags;
            _disc = disc;
            _template = template;
        }

        public void Process(Application app)
        {
            _selection = app.Selection;
            _find = _selection.Find;

            foreach (var tag in _tags)
                findExecute(tag.Key, tag.Value);

            app.ActiveDocument.SaveAs2(getPath());
            app.ActiveDocument.Undo();
        }

        public void ProcessDiscipline(Application app)
        {
            _selection = app.Selection;
            _find = _selection.Find;

            findExecute("<DISCIPLINE>", _tags["<DISCIPLINE>"]);

            app.ActiveDocument.SaveAs2(getPath());
            app.ActiveDocument.Undo();
        }
    }
}
