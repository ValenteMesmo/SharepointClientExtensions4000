using System;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class GroupExtensions
    {
        public static async Task<bool> GroupExists(this ClientContext clientContext, string name)
        {
            var groups = clientContext.LoadQuery(
                clientContext
                .Web
                .SiteGroups
                .Where(f => f.LoginName == name)
            );

            await clientContext.ExecuteQueryAsync();
            return groups.Count() > 0;
        }

        public static async Task CreateGroup(this ClientContext clientContext, string name)
        {
            if (await clientContext.GroupExists(name))
                throw new Exception($@"""{name}"" group already exists!");

            var groupCreationInformation = new GroupCreationInformation { Title = name };
            clientContext.Web.SiteGroups.Add(groupCreationInformation);
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task<Group> GetGroup(this ClientContext clientContext, string name)
        {
            var groups = clientContext.LoadQuery(
                clientContext
                .Web
                .SiteGroups
                .Where(f => f.LoginName == name)
            );

            await clientContext.ExecuteQueryAsync();
            if (groups.Count() == 0)
                throw new Exception($@"""{name}"" group not found!");

            return groups.First();
        }
    }
}
