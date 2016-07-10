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


    [Cmdlet(VerbsCommon.New, "WordList")]
    public class NewWordList : PSCmdlet
    {
        [Parameter]
        public string FilePath { get; set; }

        [Parameter]
        public string[] Items { get; set; }

        [Parameter]
        public ListItemType ListItemType { get; set; }


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
            var numberedList = this.wordDocument.AddList(this.Items[0], 0, this.ListItemType);

            for (var i = 1; i < this.Items.Length; i++)
            {
                this.wordDocument.AddListItem(numberedList, this.Items[i]);
            }
            
            this.WriteObject(numberedList);
        }
    }
}