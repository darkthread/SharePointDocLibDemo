using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointTools
{ 
    public class SPItemInfo
    {
        public FileSystemObjectType FsoType { get; set; }

        public int Id { get; set; }
        public string Path { get; set; }

        public string ParentPath =>
            string.Join("/", Path.Split('/').Take(Path.Split('/').Length - 1));
        public string Name { get; set; }
        public string Url { get; set; }
        [JsonIgnore]
        public List<SPItemInfo> Children { get; set; } = new List<SPItemInfo>();

        public SPItemInfo(string siteUrl, FileSystemObjectType fsoType, int id, string path, string name)
        {
            Id = id;
            Path = path;
            FsoType = fsoType;
            Name = name;
            Url = fsoType == FileSystemObjectType.Folder ? 
                  new Uri(siteUrl).GetLeftPart(UriPartial.Authority) + path :
                  $"{siteUrl}_layouts/15/download.aspx?SourceUrl=" + // 需依SharePoint版本調整
                  Microsoft.SharePoint.Client.Utilities.HttpUtility.UrlPathEncode(path, true, true);
        }
    }

}
