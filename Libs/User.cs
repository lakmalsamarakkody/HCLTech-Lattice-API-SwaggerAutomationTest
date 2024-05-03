using System;

namespace SwaggerWebAPI.Libs
{
    static class User
    {
        static string domain = Environment.UserDomainName;
        static string corpId = Environment.UserName;

        public static string GetDomain()
        {
            return domain;
        }

        public static string GetCorpId()
        {
            return corpId;
        }
        public static string GetUsername()
        {
            return string.Format("{0}\\{1}", domain, corpId);
        }

        public static string getEncodedUserName()
        {
            byte[] mybyte = System.Text.Encoding.UTF8.GetBytes(GetUsername());
            string returntext = System.Convert.ToBase64String(mybyte);
            return returntext;
        }
    }
}
