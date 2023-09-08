using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class CopyItem
    {
        public string GroupId { get; set; } = "";
        public string FolderId { get; set; } = "";
        public string Path { get; set; } = "";
        public string FileId { get; set; } = "";

        public CopyItem()
        {

        }

        public CopyItem(string groupId, string folderId, string path, string fileId)
        {
            GroupId = groupId;
            FolderId = folderId;
            Path = path;
            FileId = fileId;
        }
    }
}
