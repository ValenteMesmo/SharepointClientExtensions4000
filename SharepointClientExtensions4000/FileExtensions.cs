using System;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class FileExtensions
    {
        //TODO: fix bug with nested folders 
        public static async Task<bool> FolderExists(this List list, string folderUrl)
        {
            var folders = list.GetItems(CamlQuery.CreateAllFoldersQuery());
            list.Context.Load(list.RootFolder);
            list.Context.Load(folders);
            await list.Context.ExecuteQueryAsync();

            var folderRelativeUrl = string.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, folderUrl);
            folderRelativeUrl = folderRelativeUrl
              .Replace(@"\", @"/")
              .Replace(@"//", @"/");

            return Enumerable.Any(
                folders
                , folderItem => (string)folderItem["FileRef"] == folderRelativeUrl
            );
        }

        public static async Task<bool> FileExists(this List list, string fileUrl)
        {
            list.Context.Load(list.RootFolder);
            await list.Context.ExecuteQueryAsync();

            var completeRelativePath = string.Format(
                "{0}/{1}"
                , list.RootFolder.ServerRelativeUrl
                , fileUrl
            );
            completeRelativePath = completeRelativePath
              .Replace(@"\", @"/")
              .Replace(@"//", @"/");

            var context = list.Context;
            var query = new CamlQuery();
            query.ViewXml = string.Format(
                "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FileRef\"/><Value Type=\"Url\">{0}</Value></Eq></Where></Query></View>"
                , completeRelativePath
            );
            var items = list.GetItems(query);
            context.Load(items);
            await context.ExecuteQueryAsync();
            return items.Count > 0;
        }

        public static async Task<File> GetFile(this List list, string fileUrl)
        {
            list.Context.Load(list.RootFolder);
            await list.Context.ExecuteQueryAsync();

            var completeRelativePath = string.Format(
                "{0}/{1}"
                , list.RootFolder.ServerRelativeUrl
                , fileUrl
            );
            completeRelativePath = completeRelativePath
              .Replace(@"\", @"/")
              .Replace(@"//", @"/");
            var context = list.Context;
            var file = list.ParentWeb.GetFileByServerRelativeUrl(completeRelativePath);

            context.Load(file);
            await context.ExecuteQueryAsync();
            return file;
        }

        public static async Task<File> UploadFile(this List list, byte[] content, string fileUrl)
        {
            list.Context.Load(list.RootFolder);
            await list.Context.ExecuteQueryAsync();

            var completeRelativePath = string.Format(
                "{0}/{1}"
                , list.RootFolder.ServerRelativeUrl
                , fileUrl
            );
            completeRelativePath = completeRelativePath
              .Replace(@"\", @"/")
              .Replace(@"//", @"/");

            var fileCreationInfo = new FileCreationInformation
            {
                Content = content,
                Overwrite = true,
                Url = completeRelativePath
            };

            var uploadFile = list.RootFolder.Files.Add(fileCreationInfo);
            list.Context.Load(uploadFile);
            await list.Context.ExecuteQueryAsync();

            return uploadFile;
        }

        public static async Task CreateFolder(this List list, string folderName)
        {
            list.Context.Load(list.RootFolder);
            await list.Context.ExecuteQueryAsync();

            folderName = folderName
              .Replace(@"\", @"/")
              .Replace(@"//", @"/");

            await AddFolter(list.RootFolder, folderName);

            await list.Context.ExecuteQueryAsync();
        }

        private static async Task AddFolter(Folder folder, string path)
        {
            var segments = path.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (!segments.Any())
                return;

            var newFolder = folder.Folders.Add(segments.First());
            folder.Update();
            await folder.Context.ExecuteQueryAsync();
            await AddFolter(newFolder, string.Join("/", segments.Skip(1)));
        }

    }
}
