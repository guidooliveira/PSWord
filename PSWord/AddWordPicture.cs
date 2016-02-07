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
        [ValidateRange(0,360)]
        public uint Rotation { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

        //private Image documentPicture { get; set; }
        private DocX wordDocument { get; set; }
        private Paragraph paragraph { get; set; }
        
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
            
            try
            {
                ProviderInfo propertyInfo = null;
                var pictureFilePath = this.GetResolvedProviderPathFromPSPath(this.PicturePath, out propertyInfo);

                this.WriteVerbose(String.Format(@"Appending {0} to the Word Document...", pictureFilePath[0]));

                using (MemoryStream ms = new MemoryStream())
                {
                    System.Drawing.Image myImg = System.Drawing.Image.FromFile(pictureFilePath[0]);

                    myImg.Save(ms, myImg.RawFormat);  // Save your picture in a memory stream.
                    ms.Seek(0, SeekOrigin.Begin);

                    Novacode.Image image = this.wordDocument.AddImage(ms); // Create image.
                    
                    this.paragraph = this.wordDocument.InsertParagraph("", false);

                    Picture picture = image.CreatePicture();     // Create picture.
                    if (this.PictureHeight > 0)
                    {
                        picture.Height = this.PictureHeight;
                    }
                    else
                    {
                        picture.Height = myImg.Height;
                    }
                    if(this.PictureWidth > 0)
                    {
                        picture.Width = this.PictureWidth;
                    }
                    else
                    {
                        picture.Width = myImg.Width;
                    }
                    //if (!String.IsNullOrEmpty(this.PictureShape.ToString()))
                    //{
                    //    picture.SetPictureShape(this.PictureShape); // Set picture shape (if needed)
                    //}

                    if (this.Rotation > 0)
                    {
                        picture.Rotation = this.Rotation;
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
                Process.Start(this.resolvedPath);
            }
        }
    }
}