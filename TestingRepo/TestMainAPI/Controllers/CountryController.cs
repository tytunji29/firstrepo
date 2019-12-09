using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using TestingRepo.Repository;

namespace TestAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CountryController : ControllerBase
    {
        public static ReturnObject _retObj;
        public const string _errMsg = "An error occured while processing your request.";

        private readonly ICountryRepository _repo;
        static CountryController()
        {
            _retObj = new ReturnObject { Id = 0, Status = false, StatusMessage = "", Data = { } };
        }
        public CountryController(ICountryRepository repo)
        {
            _repo = repo;
        }
        [HttpGet]
        [Route("All")]
        public async Task<ReturnObject> GetAll()
        {
            try
            {
                var obj = await _repo.GetAllAsync();
                _retObj.Data = obj;
                _retObj.Status = true;
                _retObj.StatusMessage = "Successful";
                return _retObj;
            }
            catch (Exception ex)
            {
                _retObj.Data = null;
                _retObj.Status = false;
                _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
                return _retObj;
            }
        }
    }
}