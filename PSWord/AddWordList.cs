using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.IO;
using Novacode;
using System.Diagnostics;

namespace PSWord
{
    [Cmdlet(VerbsCommon.Add, "WordList")]
    public class AddWordList : PSCmdlet
    {
        [Parameter]
        public string FilePath { get; set; }

        [Parameter]
        public Novacode.List List { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }
        private string resolvedPath { get; set; }
        private DocX wordDocument { get; set; }
        protected override void BeginProcessing()
        {
            this.resolvedPath = this.GetUnresolvedProviderPathFromPSPath(this.FilePath);
            if (File.Exists(this.resolvedPath))
            {
                this.wordDocument = DocX.Load(this.resolvedPath);
            }
        }
        protected override void ProcessRecord()
        {
            this.wordDocument.InsertList(this.List);
        }

        protected override void EndProcessing()
        {
            try
            {
                using (this.wordDocument)
                {
                    this.wordDocument.SaveAs(this.resolvedPath);
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