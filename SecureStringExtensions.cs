using System.Security;

namespace SharepointMigrations
{
    public static class SecureStringExtensions
    {
        public static SecureString ToSecureString(this string value)
        {
            var secure = new SecureString();
            foreach (char c in value)
            {
                secure.AppendChar(c);
            }
            return secure;
        }

        public static string RemoveAccents(this string value)
        {

            byte[] tempBytes;
            tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(value);
            string asciiStr = System.Text.Encoding.UTF8.GetString(tempBytes);

            return asciiStr;
        }

        public static string RemoveWhiteSpaces(this string value)
        {
            return value.Replace(" ", "");
        }
    }
}
