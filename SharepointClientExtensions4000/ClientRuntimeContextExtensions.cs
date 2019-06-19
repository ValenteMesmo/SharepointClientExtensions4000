namespace Microsoft.SharePoint.Client
{
    public static class ClientRuntimeContextExtensions
    {
        internal static ClientContext AsClientContext(this ClientRuntimeContext context) => 
            (ClientContext)context;
    }
}
