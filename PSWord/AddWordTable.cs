using System;
using System.Linq;
using System.Management.Automation;
using Novacode;
using System.IO;
using System.Diagnostics;


namespace PSWord
{

    [Cmdlet(VerbsCommon.Add, "WordTable")]
    public class AddWordTable : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public string FilePath { get; set; }

        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public PSObject[] InputObject { get; set; }

        [Parameter]
        public TableDesign Design { get; set; } = TableDesign.TableNormal;

        [Parameter]
        public string PostContent { get; set; }

        [Parameter]
        public string PreContent { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

        private int Count { get; set; }
        private Table DocumentTable { get; set; }
        private DocX WordDocument { get; set; }
        private Paragraph PreContentparagraph { get; set; }
        private Paragraph PostContentparagraph { get; set; }
        private string ResolvedPath { get; set; }
        protected override void BeginProcessing()
        {
            this.ResolvedPath = this.GetUnresolvedProviderPathFromPSPath(this.FilePath);
           
            if (!File.Exists(this.ResolvedPath))
            {
                this.WordDocument = DocX.Create(this.ResolvedPath);
            }
            else
            {
                this.WordDocument = DocX.Load(this.ResolvedPath);
            }
            this.Count = 0;

            if (!string.IsNullOrEmpty(this.PreContent))
            {
                this.PreContentparagraph = this.WordDocument.InsertParagraph(this.PreContent, false);
            }
        }

        protected override void ProcessRecord()
        {
           
            var header = this.InputObject[0].Properties;
            
            try
            {
                if (this.Count == 0)
                {
                    this.DocumentTable = this.WordDocument.InsertTable(1, header.Count());
                    this.DocumentTable.Design = this.Design;
                    
                    var column = 0;
                    var row = 0;
                    foreach (var name in header)
                    {
                        this.DocumentTable.Rows[row].Cells[column].Paragraphs[0].Append(name.Name);
                        column++;
                    }
                    this.Count++;
                }
                    
                var columnIndex = 0;
                Row newRow = this.DocumentTable.InsertRow();

              
                foreach (var name in header)
                {
                    try
                    {
                        int i = 0;
                        string appendData;
                        if (this.InputObject.Length < 0)
                        {
                            appendData = this.InputObject[i++].Properties[name.Name].Value.ToString();
                        }
                        else
                        {
                            appendData = this.InputObject[0].Properties[name.Name].Value.ToString();
                        }
                        newRow.Cells[columnIndex++].Paragraphs[0].Append(appendData);
                        this.DocumentTable.Rows.Add(newRow);
                    }
                    catch (Exception exception)
                    {
                        this.WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
                    }
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
                if (!string.IsNullOrEmpty(this.PostContent))
                {
                    this.PostContentparagraph = this.WordDocument.InsertParagraph(this.PostContent, false);
                }
                using (this.WordDocument)
                {
                    this.WordDocument.Save();
                }
            }
            catch (Exception exception)
            {
                this.WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
            }
            if (this.Show.IsPresent)
            {

                Process.Start(this.ResolvedPath);
            }
        }
    }
}