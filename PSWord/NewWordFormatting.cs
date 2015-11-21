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
    public class NewWordFormatting : PSCmdlet
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

        [Parameter]
        [ValidateLength(1,72)]
        public int Size { get; set; }

        [Parameter]
        public double Spacing { get; set; }

        [Parameter]
        public StrikeThrough StrikeThrough { get; set; }

        [Parameter]
        public UnderlineStyle UnderlineStyle { get; set; }

        [Parameter]
        public KnownColor UnderlineColor { get; set; }
        protected override void ProcessRecord()
        {
            Formatting formatting = new Formatting();

            formatting.CapsStyle = this.CapsStyle;
            if (this.Bold.IsPresent)
            {
                formatting.Bold = true;
            }
            if (this.Italic.IsPresent)
            {
                formatting.Italic = true;
            }
            formatting.FontFamily = this.FontFamily;
            formatting.FontColor = Color.FromKnownColor(this.FontColor);
            if (this.Hidden.IsPresent)
            {
                formatting.Hidden = true;
            }
            if (!string.IsNullOrEmpty(this.Highlight.ToString()))
            {
                formatting.Highlight = this.Highlight;
            }
            if (!String.IsNullOrEmpty(this.Misc.ToString()))
            {
                formatting.Misc = this.Misc;
            }
           
            if (!String.IsNullOrEmpty(this.Size.ToString()))
            {
                formatting.Size = this.Size;
            }
            
            if (!String.IsNullOrEmpty(this.Spacing.ToString()))
            {
                formatting.Spacing = this.Spacing;
            }
            
            if (!String.IsNullOrEmpty(this.StrikeThrough.ToString()))
            {
                formatting.StrikeThrough = this.StrikeThrough;
            }
           
            if (!String.IsNullOrEmpty(this.UnderlineStyle.ToString()))
            {
                formatting.UnderlineStyle = this.UnderlineStyle;
            }
            
            if (!String.IsNullOrEmpty(this.UnderlineColor.ToString()))
            {
                formatting.UnderlineColor = Color.FromKnownColor(this.UnderlineColor);
            }
            
            this.WriteObject(formatting);
        }
    }
}