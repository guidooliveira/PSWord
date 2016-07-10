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
    [Cmdlet(VerbsCommon.Set, "WordListItem")]
    public class SetWordListItem : PSCmdlet
    {
        [Parameter]
        public string FilePath { get; set; }

        [Parameter]
        public Novacode.List List { get; set; }

        [Parameter]
        public string Text { get; set; }

        [Parameter]
        public int? ParentIndex { get; set; }

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
            if (this.ParentIndex == null)
            {
                this.wordDocument.AddListItem(this.List,this.Text);
            }
            else
            {
                this.wordDocument.AddListItem(this.List, this.Text, (int)this.ParentIndex);
            }

            this.WriteObject(this.List);
        }
      
    }
}