using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Novacode;
using System.IO;
using System.Diagnostics;
using System.Drawing;


namespace PSWord
{
    using System.Runtime.CompilerServices;

    [Cmdlet(VerbsCommon.Add, "WordText")]
    public class AddWordText : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory =true)]
        public string FilePath { get; set; }
        
        [Parameter(Position = 1, ValueFromPipeline =true)]
        public String[] Text { get; set; }
        
        [Parameter(HelpMessage = "Please choose a font size between 4 and 72")]
        [ValidateRange(4, 72)]
        public Int32 Size { get; set; }

        [Parameter]
        public SwitchParameter Bold { get; set; }

        [Parameter]
        public SwitchParameter Italic { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

        [Parameter]
        public FontFamily FontFamily { get; set; }

        [Parameter]
        public KnownColor FontColor { get; set; }

        private DocX wordDocument { get; set; }
        private string resolvedPath { get; set; }
        protected override void BeginProcessing()
        {
            this.resolvedPath = this.GetUnresolvedProviderPathFromPSPath(this.FilePath);

            if (!File.Exists(this.resolvedPath))
            {
                this.wordDocument = DocX.Create(this.resolvedPath);
            }
            else
            {
                this.wordDocument = DocX.Load(this.resolvedPath);
            }
        }

        protected override void ProcessRecord()
        {
            var formatting = new Formatting
            {
                FontFamily = this.FontFamily,
                Size = this.Size
            };

            

            if (this.Bold.IsPresent)
            {
                formatting.Bold = true;
            }
            if (this.Italic.IsPresent)
            {
                formatting.Italic = true;
            }

            formatting.FontColor = Color.FromKnownColor(this.FontColor);


            if (this.Text.Length > 0)
            {
                for (int i = 0; i < this.Text.Length; i++)
                {
                    Paragraph p = this.wordDocument.InsertParagraph(this.Text[i], false, formatting);
                }
            }
            else
            {
                try
                {
                    Paragraph p = this.wordDocument.InsertParagraph(this.Text[0], false, formatting);
                }
                catch (Exception exception)
                {
                    this.WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
                }
            }
        }
        protected override void EndProcessing()
        {
            try
            {
                using (this.wordDocument)
                {
                    this.wordDocument.Save();
                }
            }
            catch (Exception exception)
            {
                this.WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
            }
            if (this.Show.IsPresent)
            {
                Process.Start(this.resolvedPath);
            }
        }
    }
}

