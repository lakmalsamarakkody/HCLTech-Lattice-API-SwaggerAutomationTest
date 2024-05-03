using System.Data.SqlClient;

namespace SwaggerWebAPI.Libs
{
    class DBConnection
    {
        static string username, password, schema, hostname, metadata, provider, catalogue;

        public static void SetConnection(string host,string ctlg,string un,string pw)
        {
            hostname = host;
            catalogue = ctlg;
            username = un;
            password = pw;
        }

        public static SqlConnection GetConnection() {

            SqlConnection con = null;
            string constrng = string.Format("data source={0};initial catalog={1};user id={2};password={3};",hostname,catalogue,username,password);
            con = new SqlConnection(constrng);

            return con;
        }
    }
}
