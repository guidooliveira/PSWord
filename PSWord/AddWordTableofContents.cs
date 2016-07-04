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

    [Cmdlet(VerbsCommon.Add, "WordTableofContents", DefaultParameterSetName = "Default")]
    public class AddWordTableofContents : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public string FilePath { get; set; }

        [Parameter]
        public string Title { get; set; }

        [Parameter]
        public TableOfContentsSwitches TableSwitch { get; set; } = TableOfContentsSwitches.None;

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
            this.wordDocument.InsertTableOfContents(this.Title, this.TableSwitch, "Title");
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

