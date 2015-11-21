using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using Novacode;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace PSWord
{
    using System.Collections;
    using System.Collections.ObjectModel;
    using System.Data.Odbc;

    [Cmdlet(VerbsCommon.Add, "WordPicture")]
    public class AddWordPicture : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string FilePath { get; set; }

        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true)]
        [ValidateNotNullOrEmpty]
        public string PicturePath { get; set; }

        //[Parameter]
        //public string PostContent { get; set; }

        //[Parameter]
        //public string PreContent { get; set; }

        //[Parameter]
        //public int PictureHeight { get; set; }

        //[Parameter]
        //public int PictureWidtht { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

        //private Image documentPicture { get; set; }
        private DocX wordDocument { get; set; }
        private int indexCount { get; set; }
        private Paragraph paragraph { get; set; }
        protected override void BeginProcessing()
        {
            this.indexCount = 0;
            var resolvedPath = this.GetUnresolvedProviderPathFromPSPath(this.FilePath);

            if (!File.Exists(resolvedPath))
            {
                this.wordDocument = DocX.Create(resolvedPath);
            }
            else
            {
                this.wordDocument = DocX.Load(resolvedPath);
            }
        }

        protected override void ProcessRecord()
        {
            try
            {
                
                
                var PictureFilePath = this.GetUnresolvedProviderPathFromPSPath(this.PicturePath);


                WriteVerbose(String.Format(@"Appending {0} to the Word Document...", PictureFilePath));
                Image documentPicture = this.wordDocument.AddImage(PictureFilePath);

                Picture picture = documentPicture.CreatePicture();

                // Insert an emptyParagraph into this document.
                //if (String.IsNullOrEmpty(this.PostContent))
                //{
                //    this.paragraph = this.wordDocument.InsertParagraph("", false);
                //    this.paragraph.InsertPicture(picture, 0);
                //}
                //else
                //{
                //    this.paragraph = this.wordDocument.InsertParagraph(this.PostContent, false);
                //    this.paragraph.InsertPicture(picture, 0);
                //}

                //if (String.IsNullOrEmpty(this.PreContent))
                //{
                //    this.paragraph = this.wordDocument.InsertParagraph("", false);

                //}
                //else
                //{

                //   
                //}
                this.paragraph.InsertPicture(picture, 0);


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
                ProviderInfo providerInfo = null;
                var resolvedFile = this.GetResolvedProviderPathFromPSPath(this.FilePath, out providerInfo);
                Process.Start(resolvedFile[0]);
            }
        }
    }
}