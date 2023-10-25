using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Shared.Models
{
    public class DownloadFileResult
    {
        public bool Success { get; set; }
        public MemoryStream Contents { get; set; } = new MemoryStream();
    }
}
