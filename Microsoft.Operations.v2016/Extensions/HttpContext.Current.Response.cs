using System.Linq;
using System.Web;

namespace Microsoft.Operations
{
    public static partial class ExtensionMethods
    {
        public static string CookieRead(string cookieName)
        {
            if (HttpContext.Current.Response.Cookies.AllKeys.Contains(cookieName))
            {
                return HttpContext.Current.Response.Cookies[cookieName].Value;
            }
            else
            {
                return string.Empty;
            }
        }

        public static void CookieWrite(this HttpContext context, string cookieName, string cookieValue)
        {
            var hcCookie = new HttpCookie(cookieName, cookieValue);
            // context.Current.Response.Cookies.Set(hcCookie);
        }
    }
}