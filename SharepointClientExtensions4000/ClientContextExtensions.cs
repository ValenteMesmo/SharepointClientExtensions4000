using System;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class ClientContextExtensions
    {
        public static async Task<List> CreateList(
            this ClientContext context, string internalName, string displayName) =>
            await CreateList(context, internalName, displayName, documentLibrary: false, hidden: false);

        public static async Task<List> CreateList(
            this ClientContext context, string displayName) =>
            await CreateList(context, displayName, displayName, documentLibrary: false, hidden: false);

        public static async Task<List> CreateLibrary(this ClientContext context, string internalName, string displayName) =>
            await CreateList(context, internalName, displayName, documentLibrary: true, hidden: false);

        public static async Task<List> CreateLibrary(this ClientContext context, string displayName) =>
            await CreateList(context, displayName, displayName, documentLibrary: true, hidden: false);

        private static async Task<List> CreateList(this ClientContext clientContext, string internalName, string displayName, bool documentLibrary, bool hidden)
        {
            if (await clientContext.ListExists(displayName))
                throw new Exception($@"""{displayName}"" list already exists!");

            ListCreationInformation listCreationInfo = new ListCreationInformation();
            listCreationInfo.Title = displayName;
            listCreationInfo.TemplateType = (int)(documentLibrary ?
                ListTemplateType.DocumentLibrary
                : ListTemplateType.GenericList
            );

            if (documentLibrary)
                listCreationInfo.Url = internalName;
            else
                listCreationInfo.Url = "Lists/" + internalName;

            List list = clientContext.Web.Lists.Add(listCreationInfo);

            if (documentLibrary)
                list.ImageUrl = "/_layouts/15/images/itdl.gif?rev=45";
            else
                list.ImageUrl = "/_layouts/15/images/itgen.gif?rev=45";

            list.Hidden = hidden;
            list.EnableAttachments = false;
            list.EnableFolderCreation = false;
            list.EnableMinorVersions = false;
            list.EnableVersioning = false;
            list.AllowDeletion = false;
            list.Update();
            await clientContext.ExecuteQueryAsync();

            return list;
        }

        public static async Task RenameList(this ClientContext clientContext, string currentDisplayName, string newDisplayName)
        {
            var list = await clientContext.GetList(currentDisplayName);
            clientContext.Load(list);
            list.Title = newDisplayName;
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task DeleteList(this ClientContext clientContext, string listDisplayName)
        {
            var list = await clientContext.GetList(listDisplayName);
            list.AllowDeletion = true;
            list.Update();
            list.DeleteObject();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task<bool> ListExists(this ClientContext clientContext, string listDisplayName)
        {
            ListCollection listCollection = clientContext.Web.Lists;
            clientContext.Load(
                listCollection
                , f => f
                    .Include(g => g.Title)
                    .Where(g => g.Title == listDisplayName)
            );

            await clientContext.ExecuteQueryAsync();

            return listCollection.Count > 0;
        }

        public static async Task<List> GetList(this ClientContext clientContext, string displayName)
        {
            ListCollection listCollection = clientContext.Web.Lists;
            clientContext.Load(
                listCollection
                , f => f
                    .Include(g => g.Title)
                    .Where(g => g.Title == displayName)
            );
            await clientContext.ExecuteQueryAsync();

            if (listCollection.Count == 0)
                throw new Exception($@"""{displayName}"" list not found!");

            return listCollection.First();
        }
    }
}
