using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class ListItemExtensions
    {
        /// <summary>
        /// Helper method that simplifies the creation of a new item.
        /// </summary>
        /// <param name="itemProperties">
        /// <para>
        /// anonymous type object containing all values to the item's columns
        /// </para>
        /// <para>
        /// Example: <code>new { Title= "Example", Test = true }</code>
        /// </para>
        /// </param>
        public static async Task<ListItem> AddItem(
            this List list
            , dynamic itemProperties)
        {
            var props = itemProperties?.GetType().GetProperties();

            ListItem newItem = list.AddItem(new ListItemCreationInformation());
            foreach (var pair in props)
                newItem[pair.Name] = pair.GetValue(itemProperties);

            newItem.Update();
            await list.Context.AsClientContext().ExecuteQueryAsync();

            return newItem;
        }

        private static int OneIfZero(this int value)
        {
            if (value == 0)
                return 1;
            return value;
        }

        public static async Task SetFieldDefaultValue(this List list, string fieldName, string defaultValue)
        {
            var clientContext = list.Context.AsClientContext();
            Field field = list.Fields.GetByTitle(fieldName);
            field.DefaultValue = defaultValue;
            field.Update();
            clientContext.Load(field);
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task<IList<ListItem>> GetAllItems(this List list, string query, IProgress<int> progress = null, params Expression<Func<ListItemCollection, object>>[] retrievals)
        {
            if (progress == null)
                progress = new Progress<int>();

            var context = list.Context.AsClientContext();
            context.Load(list, f => f.ItemCount);
            await context.ExecuteQueryAsync();

            ListItemCollectionPosition itemPosition = null;
            var result = new List<ListItem>();

            while (true)
            {
                var camlQuery = new CamlQuery
                {
                    ListItemCollectionPosition = itemPosition,
                    ViewXml = query
                };

                var itemCollection = list.GetItems(camlQuery);

                {
                    var retrievalsList = (retrievals ?? new Expression<Func<ListItemCollection, object>>[] { }).ToList();
                    retrievalsList.Add(f => f.ListItemCollectionPosition);

                    context.Load(itemCollection, retrievalsList.ToArray());
                }

                try
                {
                    await context.ExecuteQueryAsync();
                }
                catch (ServerException ex)

                {
                    if (ex.Message.Contains("threshold"))
                        throw new Exception(@"List view threshold problem!
If you are using a custom query, make sure to set rowLimit.
If rowLimit is not enough to solve, create a index on your list and use that column on your query.");
                    else
                        throw;
                }

                itemPosition = itemCollection.ListItemCollectionPosition;

                foreach (ListItem item in itemCollection)
                    result.Add(item);

                progress.Report((result.Count / list.ItemCount.OneIfZero()) * 100);

                if (itemPosition == null)
                    break;
            }

            progress.Report(100);

            return result;
        }

        public static async Task<IList<ListItem>> GetAllItems(this List list, IProgress<int> progress = null, params Expression<Func<ListItemCollection, object>>[] retrievals) =>
            await list.GetAllItems(new[] { "ID", "Title" }, progress, retrievals);

        public static async Task<IList<ListItem>> GetAllItems(this List list, string[] viewFields, IProgress<int> progress = null, params Expression<Func<ListItemCollection, object>>[] retrievals)
        {
            var fields = "";

            foreach (var field in viewFields)
                fields += $"<FieldRef Name='{field}' />";

            var view = $@"
                <View>
                    <ViewFields>
                        {fields}
                    </ViewFields>
                    <RowLimit>3000</RowLimit>
                </View>";

            return await list.GetAllItems(view, progress, retrievals);
        }

        public static async Task DeleteAllItems(this List list, IProgress<int> progress = null)
        {
            if (progress == null)
                progress = new Progress<int>();

            var clientContext = list.Context.AsClientContext();

            clientContext.Load(list, f => f.ItemCount);
            await clientContext.ExecuteQueryAsync();

            var batchLimit = 100;

            var deletedItems = 0;
            var listItems = await list.GetAllItems(new Progress<int>(f => progress.Report(f / 2)));

            if (listItems.Count > 0)
            {
                for (var i = listItems.Count - 1; i > -1; i--)
                {
                    listItems[i].DeleteObject();
                    if (i % batchLimit == 0)
                        await clientContext.ExecuteQueryAsync();
                    deletedItems++;

                    progress.Report((100 + ((deletedItems / list.ItemCount.OneIfZero()) * 100)) / 2);
                }
                await clientContext.ExecuteQueryAsync();
            }

            progress.Report(100);
        }

        //https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/upload-large-files-sample-app-for-sharepoint
        public static async Task<File> UploadFile(
            this List library,
            string relativeUrl,
            byte[] content,
            int fileChunkSizeInMB = 3,
            IProgress<int> progress = null
        )
        {

            if (progress == null)
                progress = new Progress<int>();

            string filename = Path.GetFileName(relativeUrl);

            var uploadId = Guid.NewGuid();
            var fileName = Path.GetFileName(filename);

            if (relativeUrl[0] != '/')
                relativeUrl = "/" + relativeUrl;

            library.Context.Load(library.RootFolder, f => f.ServerRelativeUrl);
            await library.Context.ExecuteQueryAsync();

            Folder folder;
            var folderPath = Path.GetDirectoryName(relativeUrl);
            if (folderPath == "\\")
            {
                folder = library.RootFolder;
            }
            else
            {
                folder = library.RootFolder.Folders.GetByUrl(library.RootFolder.ServerRelativeUrl + folderPath.Replace("\\", "/"));
                library.Context.Load(folder);
                await library.Context.ExecuteQueryAsync();
            }

            File uploadFile = null;

            var blockSize = fileChunkSizeInMB * 1024 * 1024;

            var fileSize = content.Length;

            progress.Report(0);
            if (fileSize <= blockSize)
            {
                var fileInfo = new FileCreationInformation
                {
                    ContentStream = new MemoryStream(content),
                    Url = fileName,
                    Overwrite = true
                };

                uploadFile = folder.Files.Add(fileInfo);
                library.Context.Load(uploadFile);
                progress.Report(100);
                await library.Context.ExecuteQueryAsync();
                return uploadFile;
            }
            else
            {
                ClientResult<long> bytesUploaded = null;

                using (Stream stream = new MemoryStream(content))
                {
                    stream.Position = 0;
                    using (var binaryReader = new BinaryReader(stream))
                    {
                        var buffer = new byte[blockSize];
                        byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        var first = true;
                        var last = false;

                        while ((bytesRead = binaryReader.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;
                            progress.Report((int)((totalBytesRead / fileSize.OneIfZero()) * 100));

                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (var contentStream = new MemoryStream())
                                {
                                    var fileInfo = new FileCreationInformation
                                    {
                                        ContentStream = contentStream,
                                        Url = fileName,
                                        Overwrite = true
                                    };

                                    uploadFile = folder.Files.Add(fileInfo);

                                    using (var memoryStream = new MemoryStream(buffer))
                                    {
                                        memoryStream.Position = 0;
                                        bytesUploaded = uploadFile.StartUpload(uploadId, memoryStream);
                                        await library.Context.ExecuteQueryAsync();
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    first = false;
                                }
                            }
                            else
                            {
                                uploadFile = library.ParentWeb.GetFileByServerRelativeUrl(folder.ServerRelativeUrl + Path.AltDirectorySeparatorChar + fileName);

                                if (last)
                                {
                                    using (var memoryStream = new MemoryStream(lastBuffer))
                                    {
                                        memoryStream.Position = 0;
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, memoryStream);
                                        await library.Context.ExecuteQueryAsync();
                                        return uploadFile;
                                    }
                                }
                                else
                                {
                                    using (var memoryStream = new MemoryStream(buffer))
                                    {
                                        memoryStream.Position = 0;
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, memoryStream);
                                        await library.Context.ExecuteQueryAsync();
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }
                        }

                        progress.Report(100);
                    }
                }
            }

            return uploadFile;
        }
    }
}
