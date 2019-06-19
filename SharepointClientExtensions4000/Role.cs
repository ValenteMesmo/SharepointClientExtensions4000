namespace Microsoft.SharePoint.Client
{
    public struct Role
    {
        public string GroupName { get; }
        public RoleType RoleType { get; }

        public Role(string GroupName, RoleType RoleType)
        {
            this.GroupName = GroupName;
            this.RoleType = RoleType;
        }
    }
}
