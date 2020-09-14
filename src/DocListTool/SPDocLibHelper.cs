using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointTools
{
    public class SPDocLibHelper
    {
        public static bool SharePointServerMode = false;
        static AuthenticationManager authMan = new AuthenticationManager();
        public static ClientContext CreateClientContext(string siteUrl)
        {
            //若為 SharePoint Server, 確認使用者可自動登入 siteUrl 並 
            //傳回 new ClientContext(siteUrl) 即可
            if (SharePointServerMode) 
                return new ClientContext(siteUrl);
            return authMan.GetWebLoginClientContext(siteUrl);
        }

        public static List GetListByTitle(Web web, string docLibName)
        {
            List docLibList;
            var ctx = web.Context;
            if (SharePointServerMode)
            {
                docLibList = web.Lists.GetByTitle(docLibName);
            }
            else //實測 SharePoint Online GetByTitle 找中文名稱有問題，取回清單自己查
            {
                var lists = web.Lists;
                ctx.Load(lists);
                ctx.ExecuteQuery();
                docLibList = lists.Single(o => o.Title == docLibName);
            }
            ctx.Load(docLibList);
            return docLibList;
        }


        public static void ExploreDisplay(SPItemInfo item, int level)
        {
            var padding = new string(' ', level * 2);
            if (item.FsoType == FileSystemObjectType.Folder)
            {
                Console.WriteLine($"{padding}[{item.Name}]");
                item.Children.ToList().ForEach(o => ExploreDisplay(o, level + 1));
            }
            else
                Console.WriteLine($"{padding}* {Path.GetFileName(item.Path)}");
        }

        //一次取回整個文件庫所有資料夾及文件資訊，若檔案數量過多建議改用DirDocLibrary
        public static SPItemInfo GetDocLibStructure(string siteUrl, string docLibName)
        {
            using (var ctx = CreateClientContext(siteUrl))
            {
                var web = ctx.Web;
                var docLibList = GetListByTitle(web, docLibName);
                ctx.Load(docLibList);
                ctx.ExecuteQuery();
                var items = docLibList.GetItems(CamlQuery.CreateAllItemsQuery());
                ctx.Load(items, colList => colList.Include(
                    item => item.Id,
                    item => item.FileSystemObjectType,
                    item => item.DisplayName,
                    item => item["FileRef"]
                    ));
                ctx.ExecuteQuery();
                var dict = new Dictionary<string, SPItemInfo>();

                var urlPrefix = siteUrl;
                SPItemInfo root = null;
                foreach (var item in items)
                {
                    var path = (string)item["FileRef"];
                    var spItem = new SPItemInfo(siteUrl, item.FileSystemObjectType, item.Id, path, item.DisplayName);
                    var parentPath = spItem.ParentPath;
                    if (!dict.ContainsKey(parentPath))
                    {
                        //For root children
                        root = new SPItemInfo(siteUrl, FileSystemObjectType.Folder, 0, parentPath, docLibName);
                        dict.Add(parentPath, root);
                    }
                    dict[parentPath].Children.Add(spItem);
                    dict.Add(path, spItem);
                }
                return root;
            }
        }

        //取回指定資料夾下的子資料夾及檔案清單
        public static List<SPItemInfo> DirDocLibrary(string siteUrl, string docLibName, string folderPath)
        {
            using (var ctx = CreateClientContext(siteUrl))
            {
                var web = ctx.Web;

                List docLibList = GetListByTitle(web, docLibName);
                ctx.Load(docLibList.RootFolder);
                ctx.ExecuteQuery();

                //取得文件庫URL
                var docLibUrl = docLibList.RootFolder.ServerRelativeUrl;
                if (!string.IsNullOrEmpty(folderPath) && !folderPath.StartsWith("/"))
                    folderPath = "/" + folderPath;
                if (folderPath == "/") folderPath = "";
                var query = new CamlQuery();
                //Scope 
                // DefaultValue - 清單資料夾加檔案 
                // Recursive - 檔案+展開清單資料夾內項目 
                // RecursiveAll - 一路查進子資料夾不斷展開
                // FilesOnly - 只查檔案
                query.ViewXml = $@"<View Scope='RecursiveAll'>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='FileDirRef' />
                <Value Type='Text'>{docLibUrl}{folderPath}</Value>
            </Eq>
        </Where>
    </Query>
</View>";
                var listItems = docLibList.GetItems(query);
                ctx.Load(listItems, itemCol => itemCol.Include(
                    item => item.Id,
                    item => item.FileSystemObjectType,
                    item => item.DisplayName,
                    item => item["FileRef"]
                    ));
                ctx.ExecuteQuery();
                var urlPrefix = siteUrl;
                var list = new List<SPItemInfo>();
                foreach (var item in listItems)
                {
                    var path = (string)item["FileRef"];
                    var spItem = new SPItemInfo(siteUrl, item.FileSystemObjectType, item.Id, path, item.DisplayName);
                    list.Add(spItem);
                }
                return list;
            }
        }

        //新增或更新指定路徑的文件庫檔案
        public static void InsertOrUpdateFile(string siteUrl, string docLibName, string filePath, byte[] fileContent)
        {
            using (var ctx = CreateClientContext(siteUrl))
            {
                var web = ctx.Web;
                var docLibList = GetListByTitle(web, docLibName);
                ctx.Load(docLibList);
                ctx.Load(docLibList.RootFolder);
                ctx.ExecuteQuery();
                var fileRef = docLibList.RootFolder.ServerRelativeUrl + "/" + filePath;
                using (var ms = new MemoryStream(fileContent))
                {
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, fileRef, ms, true);
                }
            }
        }


    }
}
