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

        protected override void BeginProcessing()
        {
            var resolvedFile = this.GetUnresolvedProviderPathFromPSPath(this.FilePath);
            if (!File.Exists(resolvedFile))
            {
                var createDoc = DocX.Create(resolvedFile);
                createDoc.Save();
                createDoc.Dispose();
            }
        }

        protected override void ProcessRecord()
        {
            ProviderInfo providerInfo = null;
            var resolvedFile = this.GetUnresolvedProviderPathFromPSPath(this.FilePath);
            WriteVerbose(String.Format("Loading file {0}",resolvedFile[0]));

            using (DocX document = DocX.Load(resolvedFile))
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
                
                try
                {
                    foreach (var word in Text)
                    {
                        Paragraph p = document.InsertParagraph(word, false, formatting);
                    }
                        
                    document.Save();
                }
                catch
                {
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

