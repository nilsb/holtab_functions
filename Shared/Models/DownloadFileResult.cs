using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Shared.Models
{
    public class DownloadFileResult
    {
        public bool Success { get; set; }
        public Stream? Contents { get; set; }
    }
}
