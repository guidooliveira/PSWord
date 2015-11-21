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
using System.Text.RegularExpressions;

namespace PSWord
{
    

    [Cmdlet(VerbsData.Update, "WordText")]
    public class UpdateWordText : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public string FilePath { get; set; }

        [Parameter(Position = 1)]
        public String ReplacingText { get; set; }

        [Parameter(Position = 2, ValueFromPipeline = true)]
        public String NewText { get; set; }
        
        [Parameter]
        public SwitchParameter TrackChanges { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

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
            try
            {
                if (this.TrackChanges.IsPresent)
                {
                    this.wordDocument.ReplaceText(this.ReplacingText, this.NewText, true, RegexOptions.IgnoreCase);
                }
                else
                {
                    this.wordDocument.ReplaceText(this.ReplacingText, this.NewText, false, RegexOptions.IgnoreCase);
                }

                this.wordDocument.Save();
            }
            catch (Exception exception)
            {
                this.WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
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

