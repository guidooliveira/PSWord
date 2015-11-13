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

    [Cmdlet(VerbsCommon.Add, "WordTable")]
    public class AddWordTable : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public string FilePath { get; set; }

        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public PSObject[] InputObject { get; set; }

        [Parameter]
        public TableDesign Design { get; set; }

        [Parameter]
        public SwitchParameter Show { get; set; }

        private int Contagem { get; set; }
        private Table documentTable { get; set; }
        private DocX wordDocument { get; set; }
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
            this.Contagem = 0;
        }

        protected override void ProcessRecord()
        {
           
            var header = this.InputObject[0].Properties;
            
            try
            {
                if (this.Contagem == 0)
                {
                    this.documentTable = this.wordDocument.InsertTable(1, header.Count());
                    this.documentTable.Design = this.Design;

                    this.Contagem++;
                    var Column = 0;
                    var row = 0;
                    foreach (var name in header)
                    {
                        documentTable.Rows[row].Cells[Column].Paragraphs[0].Append(name.Name);
                        Column++;
                    }
                }
                    
                var ColumnIndex = 0;
                Row newRow = documentTable.InsertRow();

                foreach (var name in header)
                {
                    string appendData = this.InputObject[0].Properties[name.Name].Value.ToString();
                        
                    newRow.Cells[ColumnIndex++].Paragraphs[0].Append(appendData);
                    documentTable.Rows.Add(newRow);
                }
            }
            catch (Exception exception)
            {
                //this.WriteError(new ErrorRecord(exception, exception.HResult.ToString(), ErrorCategory.WriteError, exception));
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