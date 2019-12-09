using Dapper;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingRepo.Repository
{
    public interface ICountryRepository
    {
        Task<ReturnObject> AddAsync(CountryModel model);
        Task<List<CountryModel>> GetAllAsync();
        Task<CountryModel> GetByIdAsync();

    }
    public class CountryRepository : ICountryRepository
    {
        private readonly IConfiguration _config;
        public CountryRepository(IConfiguration config)
        {
            _config = config;
        }
        public Task<ReturnObject> AddAsync(CountryModel model)
        {
            throw new NotImplementedException();
        }

        public async Task<List<CountryModel>> GetAllAsync()
        {
            try
            {
                using (IDbConnection cn = new DapperConfig(_config).TBossDbConnection)
                {
                    string storedProcName = "spGetAllCountryForExcel";

                    var retAsync = await cn.QueryAsync<CountryModel>(storedProcName, commandType: CommandType.StoredProcedure);

                    return retAsync.ToList();
                }
            }
            catch (Exception)
            {
                return new List<CountryModel>();
            }
        }

        public Task<CountryModel> GetByIdAsync()
        {
            throw new NotImplementedException();
        }
    }
}
