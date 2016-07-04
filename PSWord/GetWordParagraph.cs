using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.IO;
using Novacode;

namespace PSWord
{


    [Cmdlet(VerbsCommon.Get, "WordParagraph")]
    public class GetWordParagraph : PSCmdlet
    {
        [Parameter]
        public string FilePath { get; set; }

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
            for (var i = 0; i < this.wordDocument.Paragraphs.Count; i++)
            {
                this.WriteObject(new
                                {
                                    Index = i,
                                    Text = this.wordDocument.Paragraphs[i].Text,
                                    StyleName = this.wordDocument.Paragraphs[i].StyleName
                });
            }
        }
    }
}
