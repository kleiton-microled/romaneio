using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;

namespace Romaneio.Helpers
{
    public static class Config
    {
        public static string StringConexao()
            => ConfigurationManager.ConnectionStrings["StringConexao"].ConnectionString;

        public static int QueryTimeoutInSeconds()
            => 1200;
    }
}