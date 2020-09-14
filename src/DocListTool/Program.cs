using SharePointTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocListTool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("** SharePoint 文件庫程式存取展示 **");
                var testNo = Input("請選擇：1) 完整檔案清單  2) 查詢子資料夾 3) 更新文件庫檔案 (1-3): ");
                switch (testNo)
                {
                    case "1":
                        FullList();
                        break;
                    case "2":
                        DirFolder();
                        break;
                    case "3":
                        TestUpload();
                        break;
                    default:
                        Console.WriteLine("無效選擇 - " + testNo);
                        return;
                }
        }

        static string Input(string prompt, bool newLine = false)
        {
            Console.Write(prompt);
            if (newLine) Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Yellow;
            var res = Console.ReadLine();
            Console.ResetColor();
            return res;
        }

        static void FullList()
        {
            var siteUrl = Input("請輸入站台網址 (例如：https://xxxx.sharepoint.com)：", true);
            var docLibName = Input("請輸入文件庫名稱(例如：文件)：");
            Console.ForegroundColor = ConsoleColor.Cyan;
            var root = SPDocLibHelper.GetDocLibStructure(siteUrl, docLibName);
            Console.WriteLine($"Directory [{root.Path}]");
            RecursiveDisplay(root.Children, 0);
        }


        static void RecursiveDisplay(IEnumerable<SPItemInfo> items, int level)
        {
            var padding = new string(' ', level * 4);
            foreach (var item in items)
            {
                if (item.FsoType == Microsoft.SharePoint.Client.FileSystemObjectType.Folder)
                {
                    Console.WriteLine($"{padding}[{item.Name}]");
                    if (item.Children.Any()) RecursiveDisplay(item.Children, level + 1);
                }
                else
                    Console.WriteLine($"{padding}{Path.GetFileName(item.Path)}");
            }
        }

        static void DirFolder()
        {
            var siteUrl = Input("請輸入站台網址 (例如：https://xxxx.sharepoint.com)：", true);
            var docLibName = Input("請輸入文件庫名稱(例如：文件)：");
            var folderPath = Input("請輸入查詢路徑(例如：/資料夾名稱/子資料夾名稱)：");
            var items = SPDocLibHelper.DirDocLibrary(siteUrl, docLibName, folderPath);
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"資料夾[{folderPath}]下的項目：");
            foreach (var item in items)
            {
                if (item.FsoType == Microsoft.SharePoint.Client.FileSystemObjectType.Folder)
                    Console.WriteLine($"資料夾 <{item.Name}>");
                else
                {
                    Console.WriteLine($"檔案 {Path.GetFileName(item.Path)}");
                    Console.WriteLine(item.Url);
                }
            }
        }

        static void TestUpload()
        {
            var siteUrl = Input("請輸入站台網址 (例如：https://xxxx.sharepoint.com)：", true);
            var docLibName = Input("請輸入文件庫名稱(例如：文件)：");
            var uploadPath = Input("請輸入查詢路徑(例如：/資料夾名稱/子資料夾名稱)：");
            SPDocLibHelper.InsertOrUpdateFile(siteUrl, docLibName, uploadPath, Encoding.UTF8.GetBytes(DateTime.Now.ToString("HH:mm:ss.fff")));
        }
    }
}
