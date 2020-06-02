using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class FileExtensions
    {
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

        public static async Task UploadFile(this List list, byte[] content, string fileUrl)
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
        }
        
        public static async Task CreateFolder(this List list, string folderName)
        {
            var completeRelativePath = string.Format(
                "{0}/{1}"
                , list.RootFolder.ServerRelativeUrl
                , folderName
            );

            completeRelativePath = completeRelativePath
              .Replace(@"\", @"/")
              .Replace(@"//", @"/");

            list.RootFolder.Folders.Add(completeRelativePath);
            list.RootFolder.Update();

            await list.Context.ExecuteQueryAsync();
        }       
    }
}
