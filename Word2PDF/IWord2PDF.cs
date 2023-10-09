using OutSystems.ExternalLibraries.SDK;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word2PDF
{
    [OSInterface(Description = "Convert File Word to PDF.")]
    public interface IWord2PDF
    {
        [OSAction(Description = "Convert Doc to PDF")]
        void Doc2PDF(
            [OSParameter(DataType = OSDataType.BinaryData, Description = "input doc binary file")]
            byte[] Doc,
            [OSParameter(Description = "File PDF")]
            out byte[] PDF,
            [OSParameter(Description = "Message")]
            out string Message,
            [OSParameter(Description = "Code")]
            out int Code
            );

        [OSAction(Description = "Convert Doc to PDF")]
        void Doc2PDF_New(
            [OSParameter(DataType = OSDataType.BinaryData, Description = "input doc binary file")]
            byte[] Doc,
            [OSParameter(Description = "File PDF")]
            out byte[] PDF,
            [OSParameter(Description = "Message")]
            out string Message,
            [OSParameter(Description = "Code")]
            out int Code
            );
    }
}
