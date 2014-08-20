using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocMigration
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Word.Application application;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            for (int i = 0; i < 100; i++)
            {
                OpenDocument();

                ElementRange elementRange = null;
                try
                {
                    var startTagMatch = FindFirstMatch(@"\<FC\>", out elementRange);

                    while (startTagMatch)
                    {
                        var startTag = new ElementRange(elementRange.Start, elementRange.End);

                        var endTagMatch = FindFirstMatch(@"\</FC\>", out elementRange);

                        if (!endTagMatch)
                        {
                            throw new InvalidOperationException("No matching end tag found");
                        }

                        ReplaceWithContentControl(new ElementRange(startTag.Start, elementRange.End));

                        startTagMatch = FindFirstMatch(@"\<FC\>", out elementRange);
                    }
                }
                finally
                {
                }
                
            }
            stopWatch.Stop();
            MessageBox.Show(stopWatch.Elapsed.TotalSeconds.ToString());
        }

        private void ReplaceWithContentControl(ElementRange elementRange)
        {
            var range = application.ActiveDocument.Range(elementRange.Start, elementRange.End);
            range.Text = range.Text.Replace("<FC>", string.Empty).Replace("</FC>", string.Empty).Replace("\r", string.Empty);
            range.ContentControls.Add(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlText, range);
        }

        private void OpenDocument()
        {
            if (application != null)
            {
                application.ActiveDocument.Close();
            }
            else
            {
                application = new Microsoft.Office.Interop.Word.Application();
                application.Visible = false;
            }
            File.Copy(@"C:\Tempy\FC2.docx", @"C:\Tempy\Output.docx", true);
            application.Documents.Open(@"C:\tempy\output.docx", Visible: true);

        }

        private bool FindFirstMatch(string wordToMatch, out ElementRange elementRange)
        {
            elementRange = null;

            var range = application.ActiveDocument.Content;
            var fnd = range.Find;

            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.MatchWildcards = true;
            fnd.Forward = true;
            fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

            fnd.Text = wordToMatch;

            fnd.Execute();

            if (fnd.Found)
            {
                elementRange = new ElementRange(range.Start, range.End);

                return true;
            }

            return false;
        }
    }


    public class ElementRange
    {
        public int Start { get; set; }
        public int End { get; set; }

        public ElementRange(int start, int end)
        {
            this.Start = start;
            this.End = end;
        }
    }
}
