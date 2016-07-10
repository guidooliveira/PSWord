using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PSWord
{
    using System.ComponentModel;

    public enum HeadingType
    {
        [Description("Heading1")]
        Heading1,

        [Description("Heading2")]
        Heading2,

        [Description("Heading3")]
        Heading3,

        [Description("Heading4")]
        Heading4,

        [Description("Heading5")]
        Heading5,

        [Description("Heading6")]
        Heading6,

        [Description("Heading7")]
        Heading7,

        [Description("Heading8")]
        Heading8,

        [Description("Heading9")]
        Heading9,
        
        [Description("NoSpacing")]
        NoSpacing,

        [Description("Title")]
        Title,

        [Description("Subtitle")]
        Subtitle,

        [Description("Quote")]
        Quote,

        [Description("IntenseQuote")]
        IntenseQuote,
        
        [Description("Emphasis")]
        Emphasis,

        [Description("IntenseEmphasis")]
        IntenseEmphasis,

        [Description("Strong")]
        Strong,

        [Description("ListParagraph")]
        ListParagraph,

        [Description("SubtleReference")]
        SubtleReference,

        [Description("IntenseReference")]
        IntenseReference,

        [Description("BookTitle")]
        BookTitle
    }
}
