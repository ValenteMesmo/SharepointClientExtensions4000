using System.Security;

namespace Microsoft.SharePoint.Client
{
    public static class SecureStringExtensions
    {
        public static SecureString ToSecureString(this string value)
        {
            var secure = new SecureString();

            foreach (char c in value)
                secure.AppendChar(c);

            return secure;
        }
    }
}
