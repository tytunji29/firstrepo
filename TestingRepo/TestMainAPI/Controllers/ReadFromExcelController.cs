//using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Linq;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using TestAPI.Helpers;

//namespace TestAPI.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class ReadFromExcelController : ControllerBase
//    {
//          private static ReturnObject _retObj;
//            private const string _errMsg = "An error occured while processing your request.";

//            private readonly IConfiguration _config;
//            private readonly IDmfInvestorComplaintRepository _repo;
//            private readonly IDmfInvestorComplaintDocRepository _docRepo;
//            private readonly IDmfComplaintRegisterRepository _dmfCompRegRepo;
//            private readonly IDmfFraudulentTransactionRepository _dmfFraudTrxnRepo;

//            static BascInvestorComplaintController()
//            {
//                _retObj = new ReturnObject { Id = 0, Status = false, StatusMessage = "", Data = { } };

//            }
//            public BascInvestorComplaintController(IConfiguration config, IDmfInvestorComplaintRepository repo, IDmfInvestorComplaintDocRepository docRepo,
//                IDmfComplaintRegisterRepository dmfCompRegRepo, IDmfFraudulentTransactionRepository dmfFraudTrxnRepo)
//            {
//                _docRepo = docRepo;
//                _dmfCompRegRepo = dmfCompRegRepo;
//                _dmfFraudTrxnRepo = dmfFraudTrxnRepo;
//                _config = config;
//                _repo = repo;
//            }

//            [HttpGet]
//            [Route("All")]
//            public async Task<ReturnObject> GetAllAsync()
//            {
//                try
//                {
//                    var dmf = User.Identity?.DmfId();
//                    if (string.IsNullOrWhiteSpace(dmf))
//                    {
//                        // Abort current operation
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = "Invalid Token, kindly login again.";
//                        return _retObj;
//                    }
//                    int.TryParse(dmf, out int dmfId);
//                    var data = await _repo.GetAll(dmfId);
//                    _retObj.Status = true;
//                    _retObj.StatusMessage = "Successful!";
//                    _retObj.Data = data;

//                    return _retObj;
//                }
//                catch (Exception ex)
//                {
//                    ActivityLogger.ErrorFileLog($"{ex.Message}, More: {ex.InnerException?.Message}, {Environment.NewLine} Stack Trace: {ex.StackTrace}");
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
//                    return _retObj;
//                }
//            }

//            [HttpGet]
//            [Route("ById/{id}")]
//            public async Task<ReturnObject> GetByIdAsync([FromRoute] int id)
//            {
//                if (!ModelState.IsValid)
//                {
//                    string errorMessages = string.Join(" ", ModelState.Values.SelectMany(x => x.Errors).Select(x => x.ErrorMessage));

//                    _retObj.Status = false;
//                    _retObj.StatusMessage = errorMessages;
//                    return _retObj;
//                }
//                try
//                {

//                    var data = await _repo.GetById(id);
//                    _retObj.Data = data;
//                    _retObj.Status = true;
//                    _retObj.StatusMessage = "Success";
//                    return _retObj;
//                }
//                catch (Exception ex)
//                {
//                    ActivityLogger.ErrorFileLog($"{ex.Message}, More: {ex.InnerException?.Message}, {Environment.NewLine} Stack Trace: {ex.StackTrace}");
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
//                    _retObj.Data = null;

//                    return _retObj;
//                }
//            }

//            [HttpPost]
//            [Route("Add")]
//            public async Task<ReturnObject> Post([FromForm]DmfInvestorComplaintFormModel model)
//            {
//                if (!ModelState.IsValid)
//                {
//                    string errorMessages = string.Join(" ", ModelState.Values.SelectMany(x => x.Errors).Select(x => x.ErrorMessage));

//                    _retObj.Status = false;
//                    _retObj.StatusMessage = errorMessages;
//                    return _retObj;
//                }
//                try
//                {
//                    var user = User.Identity.Name;
//                    var dmf = User.Identity?.DmfId();
//                    if (string.IsNullOrWhiteSpace(dmf))
//                    {
//                        // Abort current operation
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = "Invalid Token, kindly login again.";
//                        return _retObj;
//                    }
//                    int.TryParse(dmf, out int dmfId);

//                    var obj = new DmfInvestorComplaintSubmitModel
//                    {
//                        CreatedBy = user,
//                        DmfId = dmfId,
//                        Description = model.Description,
//                        StartDate = model.StartDate,
//                        EndDate = model.EndDate
//                    };

//                    if (model.File != null)
//                    {
//                        using (var memoryStream = new MemoryStream())
//                        {
//                            await model.File.CopyToAsync(memoryStream);
//                            var nuArr = memoryStream.ToArray();
//                            obj.FileData = nuArr;
//                        }

//                        var ext = Path.GetExtension(model.File.FileName).ToLowerInvariant();
//                        if (ext.ToLower() != ".xlsx")
//                        {
//                            _retObj.Status = false;
//                            _retObj.StatusMessage = $"An error occured while processing your request. Invalid file type!";
//                            return _retObj;
//                        }

//                        obj.Tag = ext;
//                        obj.IsFileUploaded = true;
//                    }
//                    var retVal = await _repo.AddAsync(obj);
//                    if (retVal.Status)
//                    {
//                        BackgroundJob.Enqueue(() => ProcessUploadedDocAsync((int)retVal.Id, user));
//                    }
//                    return retVal;
//                }
//                catch (Exception ex)
//                {
//                    ActivityLogger.ErrorFileLog($"{ex.Message}, More: {ex.InnerException?.Message}, {Environment.NewLine} Stack Trace: {ex.StackTrace}");
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"An error occured while processing your request. {ex.Message}";
//                    return _retObj;
//                }
//            }

//            [HttpGet]
//            [Route("ByDocId/{id}")]
//            public async Task<ReturnObject> GetDocId([FromRoute] int id)
//            {
//                if (!ModelState.IsValid)
//                {
//                    string errorMessages = string.Join(" ", ModelState.Values.SelectMany(x => x.Errors).Select(x => x.ErrorMessage));

//                    _retObj.Status = false;
//                    _retObj.StatusMessage = errorMessages;
//                    return _retObj;
//                }

//                try
//                {
//                    var obj = await _docRepo.GetByIdAsync(id);
//                    _retObj.Data = obj;
//                    if (obj == null)
//                    {
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = "Does not exist!";
//                        return _retObj;
//                    }

//                    _retObj.Status = true;
//                    _retObj.StatusMessage = "Success";
//                    _retObj.Data = obj;

//                    return _retObj;
//                }
//                catch (Exception ex)
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
//                    return _retObj;
//                }
//            }

//            [HttpGet]
//            [Route("downloadById/{Id}")]
//            public async Task<IActionResult> DownloadDoc([FromRoute] int Id)
//            {
//                if (!ModelState.IsValid)
//                {
//                    return BadRequest(ModelState);
//                }
//                try
//                {
//                    var docRec = await _docRepo.GetByDmfInvestorComplaintIdAsync(Id);

//                    if (docRec == null)
//                    {
//                        return BadRequest(new ReturnObject { Id = 0, Status = false, StatusMessage = "No doc for this record!" });
//                    }

//                    var docStream = new MemoryStream(docRec.FileData);

//                    return File(docStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Doc.xlsx"); ;
//                }
//                catch (Exception ex)
//                {
//                    ActivityLogger.ErrorFileLog($"{ex.Message}, More: {ex.InnerException?.Message}, {Environment.NewLine} Stack Trace: {ex.StackTrace}");
//                    return BadRequest(new ReturnObject { Id = 0, Status = false, StatusMessage = $"An error occured while processing your request. {ex.Message}" });
//                }
//            }


//            [NonAction]
//            public async Task ProcessUploadedDocAsync(int investorComplaintId, string user)
//            {
//                try
//                {
//                    var getDocRec = await _docRepo.GetByDmfInvestorComplaintIdAsync(investorComplaintId);
//                    if (getDocRec == null)
//                    {
//                        return;
//                    }
//                    var lsFraudulEntTrxn = new List<DmfFraudulentTransactionModel>();
//                    var lsComplaintReg = new List<DmfComplaintRegisterModel>();

//                    using (MemoryStream ms = new MemoryStream(getDocRec.FileData))
//                    {
//                        using (ExcelPackage p = new ExcelPackage(ms))
//                        {
//                            // get the first worksheet in the workbook which is Fraudulent Transactions
//                            ExcelWorksheet ws1 = p.Workbook.Worksheets[0];
//                            for (int row = 3; row < int.MaxValue; row++)
//                            {
//                                var firstRwCell = ws1.Cells[row, 1].Value?.ToString();
//                                if (string.IsNullOrWhiteSpace(firstRwCell))
//                                {
//                                    break;
//                                }
//                                var nuFraud = new DmfFraudulentTransactionModel
//                                {
//                                    ComplaintDate = ws1.Cells[row, 2].Value?.ToString(),
//                                    AccountInvolved = ws1.Cells[row, 3].Value?.ToString(),
//                                    TransactionDate = ws1.Cells[row, 4].Value?.ToString(),
//                                    ComplaintDetails = ws1.Cells[row, 5].Value?.ToString(),
//                                    FrequencyOfOccurence = ws1.Cells[row, 6].Value?.ToString(),
//                                    ValueOfTransaction = ws1.Cells[row, 7].Value?.ToString(),
//                                    SharesInvolved = ws1.Cells[row, 8].Value?.ToString(),
//                                    Comment = ws1.Cells[row, 9].Value?.ToString(),
//                                    CreatedBy = user,
//                                    DateCreated = DateTime.Now,
//                                    InvestorComplaintId = investorComplaintId,
//                                };

//                                lsFraudulEntTrxn.Add(nuFraud);
//                            }

//                            // get the first worksheet in the workbook which is Complaint Register
//                            ExcelWorksheet ws2 = p.Workbook.Worksheets[1];
//                            for (int row = 2; row < int.MaxValue; row++)
//                            {
//                                var firstRwCell = ws2.Cells[row, 1].Value?.ToString();
//                                if (string.IsNullOrWhiteSpace(firstRwCell))
//                                {
//                                    break;
//                                }
//                                var nuComplaint = new DmfComplaintRegisterModel
//                                {
//                                    NameOfComplainant = ws2.Cells[row, 2].Value?.ToString(),
//                                    ComplaintDate = ws2.Cells[row, 3].Value?.ToString(),
//                                    NatureOfComplaint = ws2.Cells[row, 4].Value?.ToString(),
//                                    ComplaintDetailsInBrief = ws2.Cells[row, 5].Value?.ToString(),
//                                    Remarks = ws2.Cells[row, 6].Value?.ToString(),
//                                    StatusOfComplaint = ws2.Cells[row, 7].Value?.ToString(),
//                                    CreatedBy = user,
//                                    DateCreated = DateTime.Now,
//                                    InvestorComplaintId = investorComplaintId,
//                                };

//                                lsComplaintReg.Add(nuComplaint);
//                            }
//                        }
//                    }

//                    using (IDbConnection cn = new DapperConfig(_config).XBossDbConnection)
//                    {
//                        var retFraud = await _dmfFraudTrxnRepo.AddAsync(lsFraudulEntTrxn, cn);
//                        if (!retFraud.Status)
//                        {
//                            ActivityLogger.ErrorFileLog($"Error Occured while trying to process Fraudulent Transaction Data from BASC; {retFraud.StatusMessage}");
//                        }

//                        var retCompReg = await _dmfCompRegRepo.AddAsync(lsComplaintReg, cn);
//                        if (!retCompReg.Status)
//                        {
//                            ActivityLogger.ErrorFileLog($"Error Occured while trying to process Complaint Register Data from BASC {retCompReg.StatusMessage}");
//                        }
//                    }
//                }
//                catch (Exception ex)
//                {
//                    ActivityLogger.ErrorFileLog($"{ex.Message}, More: {ex.InnerException?.Message}, {Environment.NewLine} Stack Trace: {ex.StackTrace}");
//                }
//            }
//        }
//    }