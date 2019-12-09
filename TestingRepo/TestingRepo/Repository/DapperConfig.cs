using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace TestingRepo.Repository
{
    public class DapperConfig
    {
        private readonly IConfiguration _config;
        public DapperConfig(IConfiguration config)
        {
            _config = config;
        }

        public IDbConnection TBossDbConnection => new SqlConnection(_config.GetConnectionString("TBossDB"));
    }
   
}
