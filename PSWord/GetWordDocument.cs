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
    [Cmdlet(VerbsCommon.Get, "WordDocument")]
    public class GetWordDocument : PSCmdlet
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
            this.WriteObject(
                new
                    {
                        DocumentName = Path.GetFileName(this.resolvedPath),
                        Paragraphs = this.wordDocument.Paragraphs.Count,
                        Tables = this.wordDocument.Tables.Count,
                        Images =  this.wordDocument.Images.Count,
                        IsProtected = this.wordDocument.isProtected
                    });
        }
    }
}