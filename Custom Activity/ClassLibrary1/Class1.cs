using System;
using System.Collections.Generic;
using System.Activities;
using System.ComponentModel;
using Mword  = Microsoft.Office.Interop.Word;
namespace Emphasis.Activities
{
    [DisplayName("Read Bold Texts")]
    public class ReadBoldTexts:CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public OutArgument<List<string>> Output { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var path = FilePath.Get(context);
            List<string> BoldCharacters = new List<string>();
            var wapp = new Mword.Application();
            wapp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            var oDoc = wapp.Documents.Open(path,false, false, false);
             //var Senetence =new Mword.Range();
            //Range Senetence;
            foreach (Mword.Range Sentence in oDoc.Sentences)
            {
                if (Sentence.Font.Bold == -1 || Sentence.Font.BoldBi == -1)
                {
                    BoldCharacters.Add(Sentence.Text);
                    //Console.WriteLine(Sentence.Text);
                }
            }

            Output.Set(context, BoldCharacters);
            oDoc.Close();

        }
    }
    [DisplayName("Read Italics Texts")]
    public class ReadItalicsTexts : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> FilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public OutArgument<List<string>> Output { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var path = FilePath.Get(context);
            List<string> BoldCharacters = new List<string>();
            var wapp = new Mword.Application();
            wapp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            var oDoc = wapp.Documents.Open(path, false, false, false);
            //var Senetence =new Mword.Range();
            //Range Senetence;
            foreach (Mword.Range Sentence in oDoc.Sentences)
            {
                if (Sentence.Font.Italic == -1 || Sentence.Font.ItalicBi == -1)
                {
                    BoldCharacters.Add(Sentence.Text);
                    //Console.WriteLine(Sentence.Text);
                }
            }

            Output.Set(context, BoldCharacters);
            oDoc.Close();

        }
    }
}
