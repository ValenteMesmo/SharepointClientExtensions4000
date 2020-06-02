using System.IO;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class FileExtensions
    {
        public static async Task UploadFile(this List list, byte[] content, string fileUrl)
        {
            var fileCreationInfo = new FileCreationInformation
            {
                Content = content,
                Overwrite = true,
                Url = Path.Combine(fileUrl)
            };

            var uploadFile = list.RootFolder.Files.Add(fileCreationInfo);
            list.Context.Load(uploadFile);
            await list.Context.ExecuteQueryAsync();
        }

        public static async Task<bool> FileExists(this List list, string fileUrl)
        {
            var context = list.Context;
            var query = new CamlQuery();
            query.ViewXml = string.Format(
                "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FileRef\"/><Value Type=\"Url\">{0}</Value></Eq></Where></Query></View>"
                , fileUrl
            );
            var items = list.GetItems(query);
            context.Load(items);
            await context.ExecuteQueryAsync();
            return items.Count > 0;
        }
    }
}
