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
using System.Drawing;

namespace PSWord
{
    using System.Collections;
    using System.Collections.ObjectModel;
    using System.Data.Odbc;
    using System.Runtime.CompilerServices;

    using Image = System.Drawing.Image;

    [Cmdlet(VerbsCommon.Add, "WordPicture")]
    public class AddWordPicture : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string FilePath { get; set; }

        [Parameter]
        [ValidateNotNullOrEmpty]
        public BasicShapes PictureShape { get; set; }

        [Parameter]
        [ValidateNotNullOrEmpty]
        public string PicturePath { get; set; }

        //[Parameter]
        //public string PostContent { get; set; }

        //[Parameter]
        //public string PreContent { get; set; }

        [Parameter]
        public int PictureHeight { get; set; }

        [Parameter]
        public int PictureWidth { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

        //private Image documentPicture { get; set; }
        private DocX wordDocument { get; set; }
        private Paragraph paragraph { get; set; }
        private string PictureFilePath { get; set; }
        protected override void BeginProcessing()
        {
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

                this.PictureFilePath = this.GetUnresolvedProviderPathFromPSPath(this.PicturePath);

                WriteVerbose(String.Format(@"Appending {0} to the Word Document...", PictureFilePath));

                using (MemoryStream ms = new MemoryStream())
                {
                    System.Drawing.Image myImg = System.Drawing.Image.FromFile(PictureFilePath);

                    myImg.Save(ms, myImg.RawFormat);  // Save your picture in a memory stream.
                    ms.Seek(0, SeekOrigin.Begin);

                    Novacode.Image image = this.wordDocument.AddImage(ms); // Create image.
                    
                    this.paragraph = this.wordDocument.InsertParagraph("", false);

                    Picture picture = image.CreatePicture();     // Create picture.
                    if (!String.IsNullOrEmpty(this.PictureHeight.ToString()))
                    {
                        picture.Height = this.PictureHeight;
                    }
                    else
                    {
                        picture.Height = myImg.Height;
                    }
                    if (String.IsNullOrEmpty(this.PictureWidth.ToString()))
                    {
                        picture.Width = this.PictureWidth;
                    }
                    else
                    {
                        picture.Width = myImg.Width;
                    }
                    if(!String.IsNullOrEmpty(this.PictureShape.ToString()))
                    {
                        picture.SetPictureShape(this.PictureShape); // Set picture shape (if needed)
                    }

                    this.paragraph.InsertPicture(picture, 0); // Insert picture into paragraph.
                }


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