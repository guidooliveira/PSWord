using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;

namespace PSWord
{
    using System.Drawing;
    using System.Management.Automation;

    [Cmdlet(VerbsCommon.New, "WordFormatting")]
    class NewWordFormatting : PSCmdlet
    {
        [Parameter]
        public CapsStyle CapsStyle { get; set; }

        [Parameter]
        public SwitchParameter Bold { get; set; }

        [Parameter]
        public SwitchParameter Italic { get; set; }

        [Parameter]
        public FontFamily FontFamily { get; set; }

        [Parameter]
        public KnownColor FontColor { get; set; }

        [Parameter]
        public SwitchParameter Hidden { get; set; }

        [Parameter]
        public Highlight Highlight { get; set; }

        [Parameter]
        public Misc Misc { get; set; }

        [Parameter(Mandatory = true)]
        public int? Size { get; set; }

        [Parameter]
        public double Spacing { get; set; }

        [Parameter]
        public StrikeThrough StrikeThrough { get; set; }

        [Parameter]
        public UnderlineStyle UnderlineStyle { get; set; }

        [Parameter]
        public KnownColor UnderlineColor { get; set; }
        private Formatting formatting { get; set; }

        protected override void BeginProcessing()
        {
            this.formatting = new Formatting
                                {
                                   Size = this.Size
                                };
           
        }

        protected override void ProcessRecord()
        {
            this.formatting.CapsStyle = this.CapsStyle;
            if (this.Bold.IsPresent)
            {
                this.formatting.Bold = true;
            }
            if (this.Italic.IsPresent)
            {
                this.formatting.Italic = true;
            }
            this.formatting.FontFamily = this.FontFamily;
            this.formatting.FontColor = Color.FromKnownColor(this.FontColor);
            if (this.Hidden.IsPresent)
            {
                this.formatting.Hidden = true;
            }
            if (!string.IsNullOrEmpty(this.Highlight.ToString()))
            {
                this.formatting.Highlight = this.Highlight;
            }
            if (!String.IsNullOrEmpty(this.Misc.ToString()))
            {
                this.formatting.Misc = this.Misc;
            }
            if (!String.IsNullOrEmpty(this.Spacing.ToString()))
            {
                this.formatting.Spacing = this.Spacing;
            }
            
            if (!String.IsNullOrEmpty(this.StrikeThrough.ToString()))
            {
                this.formatting.StrikeThrough = this.StrikeThrough;
            }
           
            if (!String.IsNullOrEmpty(this.UnderlineStyle.ToString()))
            {
                this.formatting.UnderlineStyle = this.UnderlineStyle;
            }
            
            if (!String.IsNullOrEmpty(this.UnderlineColor.ToString()))
            {
                this.formatting.UnderlineColor = Color.FromKnownColor(this.UnderlineColor);
            }
            
            this.WriteObject(this.formatting);
        }
    }
}