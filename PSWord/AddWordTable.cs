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
    

    [Cmdlet(VerbsCommon.Add, "WordTable")]
    class AddWordTable : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public string FilePath { get; set; }

        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public Object InputObject { get; set; }
        
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
                Table docTable = document.InsertTable(1, this.InputObject.GetType().GetProperties().Count());
                try
                {
                    foreach (PropertyInfo prp in this.InputObject.GetType().GetProperties())
                    {
                        if (prp.CanRead)
                        {
                            object value = prp.GetValue(InputObject, null);
                            string s;
                            if (value == null)
                            {
                                s = "";
                            }
                            else
                            {
                                s = value.ToString();
                            }
                            string name = prp.Name;
                        }
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
}
