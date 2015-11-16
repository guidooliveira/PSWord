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
        protected override void BeginProcessing()
        {
            var fullPath = Path.GetFullPath(this.FilePath);
            if (!File.Exists(fullPath))
            {
                var createDoc = DocX.Create(fullPath);
                createDoc.Save();
                createDoc.Dispose();
            }
        }

        protected override void ProcessRecord()
        {
            ProviderInfo providerInfo = null;
            var resolvedFile = this.GetResolvedProviderPathFromPSPath(this.FilePath, out providerInfo);
            WriteVerbose(String.Format("Loading file {0}", resolvedFile[0]));

            using (DocX document = DocX.Load(resolvedFile[0]))
            {
                

                try
                {
                    if (this.TrackChanges.IsPresent)
                    {
                        document.ReplaceText(this.ReplacingText, this.NewText, true, RegexOptions.IgnoreCase);
                    }
                    else
                    {
                        document.ReplaceText(this.ReplacingText, this.NewText, false, RegexOptions.IgnoreCase);
                    }

                    document.Save();
                }
                catch (Exception exception)
                {
                    WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
                }
                finally
                {
                    if (this.Show.IsPresent)
                    {
                        Process.Start(resolvedFile[0]);
                    }
                }
            }
        }
    }
}

