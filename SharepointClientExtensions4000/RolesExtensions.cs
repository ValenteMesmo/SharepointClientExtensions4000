using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class RolesExtensions
    {
        public static async Task AddRole(this Group group, RoleType role)
        {
            var clientContext = group.Context.AsClientContext();

            var roles = clientContext.Web.RoleAssignments;
            clientContext.Load(roles);
            await clientContext.ExecuteQueryAsync();

            var roletype = clientContext.Web.RoleDefinitions.GetByType(role);
            clientContext.Load(roletype);
            await clientContext.ExecuteQueryAsync();

            roles.Add(
                group
                , new RoleDefinitionBindingCollection(clientContext) {
                    roletype
                }
            );

            await clientContext.ExecuteQueryAsync();
        }

        public static async Task ClearRoles(this Group group)
        {
            var clientContext = group.Context.AsClientContext();

            var roles = clientContext.LoadQuery(
                clientContext
                .Web
                .RoleAssignments
                .Where(f => f.PrincipalId == group.Id)
            );
            await clientContext.ExecuteQueryAsync();

            foreach (var item in roles)
            {
                item.DeleteObject();
                await clientContext.ExecuteQueryAsync();
            }
        }

        public static async Task SetRoles(this List list, params Role[] Roles)
        {
            var clientContext = list.Context.AsClientContext();

            list.BreakRoleInheritance(false, true);
            list.Update();
            await clientContext.ExecuteQueryAsync();

            var listRoles = list.RoleAssignments;
            clientContext.Load(listRoles);
            await clientContext.ExecuteQueryAsync();

            if (listRoles.Count > 0)
            {
                for (var counter = listRoles.Count - 1; counter > -1; counter--)
                {
                    listRoles[counter].DeleteObject();
                    await clientContext.ExecuteQueryAsync();
                }
            }

            foreach (var role in Roles)
            {
                var group = await clientContext.GetGroup(role.GroupName);

                var roletype = clientContext.Web.RoleDefinitions.GetByType(role.RoleType);
                clientContext.Load(roletype);
                await clientContext.ExecuteQueryAsync();

                var collRoleDefinitionBinding = new RoleDefinitionBindingCollection(clientContext) {
                    roletype
                };
                listRoles.Add(group, collRoleDefinitionBinding);

                await clientContext.ExecuteQueryAsync();
            }
        }
    }
}
