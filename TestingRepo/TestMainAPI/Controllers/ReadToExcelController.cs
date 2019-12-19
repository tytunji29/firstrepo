//using System;
//using System.Collections.Generic;
//using System.Drawing;
//using System.IO;
//using System.Linq;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using OfficeOpenXml;
//using OfficeOpenXml.Style;

//namespace TestAPI.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class ReadToExcelController : ControllerBase
//    {
//        public static ReturnObject _retObj;

//        public const string _errMsg = "An error occured while processing your request.";
//        private readonly ISubmissionRepository _repo;
//        private readonly IOperationNotificationRepository _opNotifRepo;
//        private readonly UserManager<XBossUser> _userManager;
//        static ComplaintReportController()
//        {
//            _retObj = new ReturnObject { Id = 0, Status = false, StatusMessage = "", Data = { } };
//        }
//        public ComplaintReportController(UserManager<XBossUser> userManager, IOperationNotificationRepository opNotifRepo, ISubmissionRepository repo)
//        {
//            _repo = repo;
//            _userManager = userManager;
//            _opNotifRepo = opNotifRepo;
//        }

//        [HttpPost]
//        [Route("FilterReport")]
//        public async Task<ReturnObject> FilterReport([FromBody] ComplaintReportFormModel model)
//        {
//            if (!ModelState.IsValid)
//            {
//                string errorMessages = string.Join(" ", ModelState.Values.SelectMany(x => x.Errors).Select(x => x.ErrorMessage));

//                _retObj.Status = false;
//                _retObj.StatusMessage = errorMessages;
//                return _retObj;
//            }

//            try
//            {
//                var user = User.Identity.Name;
//                var obj = new SubmissionReportModel
//                {
//                    SubmissionActionStatus = model.SubmissionActionStatus,
//                    Channel = model.Channel,
//                    EndDate = model.EndDate,
//                    EntityId = model.EntityId,
//                    OngoingConductType = model.OngoingConductType,
//                    StartDate = model.StartDate,
//                    Status = model.Status,
//                    SubmissionKind = 1,
//                    SubmissionOwner = (byte)SubmissionOwnerEnum.ComplaintUnit,
//                    SubmissionType = model.SubmissionType
//                };

//                var retLsData = await _repo.GetComplaintDetailReportDataAsync(obj);
//                var retSummary = await _repo.GetComplaintSummaryReportDataAsync(obj);

//                _retObj.Status = true;
//                _retObj.StatusMessage = "Success!";

//                _retObj.Data = new { Payload = retLsData, Summary = retSummary };

//                return _retObj;
//            }


//            catch (Exception ex)
//            {
//                _retObj.Status = false;
//                _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
//                return _retObj;
//            }
//        }
//        [HttpPost]
//        [Route("Report")]
//        public async Task<IActionResult> Report([FromBody] ReportForComplaintFormModel model)
//        {
//            if (!ModelState.IsValid)
//            {
//                string errorMessages = string.Join(" ", ModelState.Values.SelectMany(x => x.Errors).Select(x => x.ErrorMessage));

//                _retObj.Status = false;
//                _retObj.StatusMessage = errorMessages;
//                return StatusCode(202, _retObj);
//            }


//            try
//            {
//                var userEmail = User.Identity.Name;

//                if (string.IsNullOrWhiteSpace(userEmail))
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} Unknown User";

//                    return Unauthorized(_retObj);
//                }

//                var user = await _userManager.FindByEmailAsync(userEmail);
//                Stream strm = null;
//                var lsSubThisWeek = await _repo.GetReportByDateCreatedDateRangeByOwnerByChannelByEntityForComplaint(new SubmissionByOwnerByChannelByEntityModel
//                {
//                    Channel = model.Channel,
//                    EndDate = model.EndDate,
//                    Entity = model.Entity,
//                    EntityId = model.EntityId,
//                    EntityType = model.EntityType,
//                    OngoingConductType = model.OngoingConductType,
//                    OwnerEnum = SubmissionOwnerEnum.ComplaintUnit,
//                    StartDate = model.StartDate,
//                    Status = model.Status,
//                    SubmissionActionStatus = model.SubmissionActionStatus,
//                    SubmissionType = model.SubmissionType
//                });
//                if (lsSubThisWeek.Count < 1)
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                    return StatusCode(202, _retObj);
//                }

//                if (lsSubThisWeek.Count < 1)
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                    return StatusCode(202, _retObj);
//                }

//                if (model.ReportType == 1)
//                {
//                    strm = ProcessDataForExcel(model.StartDate, model.EndDate, lsSubThisWeek);
//                    if (strm == null)
//                    {
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                        return StatusCode(202, _retObj);
//                    }

//                    strm.Seek(0, SeekOrigin.Begin);
//                }
//                else
//                {
//                    strm = ProcessDataViewForExcel(lsSubThisWeek);

//                    if (strm == null)
//                    {
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                        return StatusCode(202, _retObj);
//                    }

//                    strm.Seek(0, SeekOrigin.Begin);
//                }
//                if (model.DeliveryMethod == 1)
//                {
//                    //
//                    return File(strm, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
//                }
//                else if (model.DeliveryMethod == 2)
//                {
//                    using (MemoryStream ms = new MemoryStream())
//                    {
//                        strm.Position = 0;
//                        ms.Position = 0;

//                        await strm.CopyToAsync(ms);
//                        var msArr = ms.ToArray();

//                        //// Report Notification to User using Hangfire.
//                        BackgroundJob.Enqueue(() => RunNotif(msArr, $"{user.LastName} {user.FirstName}", userEmail));

//                        _retObj.Status = true;
//                        _retObj.StatusMessage = $"Report has been queued for sending to your Email Address.";

//                        return StatusCode(202, _retObj);
//                    }
//                }
//                else
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} Wrong Delivery Type";

//                    return StatusCode(202, _retObj);
//                }
//            }
//            catch (Exception ex)
//            {
//                _retObj.Status = false;
//                _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
//                return StatusCode(202, _retObj);
//            }
//        }

//        [HttpPost]
//        [Route("SubmissionsReport")]
//        public async Task<IActionResult> ActionsOnSubmissionsReport([FromBody] SubmissionsReportFormModel model)
//        {
//            if (!ModelState.IsValid)
//            {
//                string errorMessages = string.Join(" ", ModelState.Values.SelectMany(x => x.Errors).Select(x => x.ErrorMessage));

//                _retObj.Status = false;
//                _retObj.StatusMessage = errorMessages;
//                return StatusCode(202, _retObj);
//            }

//            try
//            {
//                var userEmail = User.Identity.Name;

//                if (string.IsNullOrWhiteSpace(userEmail))
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} Unknown User";

//                    return Unauthorized(_retObj);
//                }

//                var user = await _userManager.FindByEmailAsync(userEmail);
//                Stream strm = null;

//                //var lsSubThisWeek = await _repo.GetReportByDateCreatedDateRangeByOwnerByChannel(model.StartDate.Date, model.EndDate.Date, model.SubmissionActionStatus, model.SubmissionNature, SubmissionOwnerEnum.MSI,model.OngoingConductType, model.Channel);
//                var lsSubThisWeek = await _repo.GetReportByDateCreatedDateRangeByOwnerByChannelByEntity(new SubmissionByOwnerByChannelByEntityModel
//                {
//                    Channel = model.Channel,
//                    EndDate = model.EndDate,
//                    Entity = model.Entity,
//                    EntityId = model.EntityId,
//                    EntityType = model.EntityType,
//                    OngoingConductType = model.OngoingConductType,
//                    OwnerEnum = 0,
//                    StartDate = model.StartDate,
//                    Status = model.Status,
//                    SubmissionActionStatus = model.SubmissionActionStatus,
//                    SubmissionKind = model.SubmissionKind,
//                    SubmissionNature = model.SubmissionNature,
//                    SubmissionType = model.SubmissionType
//                });

//                if (lsSubThisWeek.Count < 1)
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                    return StatusCode(202, _retObj);
//                }

//                if (model.ReportType == 1)
//                {
//                    strm = ProcessDataForExcel(model.StartDate, model.EndDate, lsSubThisWeek);
//                    if (strm == null)
//                    {
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                        return StatusCode(202, _retObj);
//                    }

//                    strm.Seek(0, SeekOrigin.Begin);
//                }
//                else
//                {
//                    strm = ProcessDataViewForExcel(lsSubThisWeek);

//                    if (strm == null)
//                    {
//                        _retObj.Status = false;
//                        _retObj.StatusMessage = $"{_errMsg} No Record to generate!";

//                        return StatusCode(202, _retObj);
//                    }

//                    strm.Seek(0, SeekOrigin.Begin);
//                }

//                if (model.DeliveryMethod == 1)
//                {
//                    //
//                    return File(strm, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
//                }
//                else if (model.DeliveryMethod == 2)
//                {
//                    using (MemoryStream ms = new MemoryStream())
//                    {
//                        strm.Position = 0;
//                        ms.Position = 0;

//                        await strm.CopyToAsync(ms);
//                        var msArr = ms.ToArray();

//                        //// Report Notification to User using Hangfire.
//                        BackgroundJob.Enqueue(() => RunNotif(msArr, $"{user.LastName} {user.FirstName}", userEmail));

//                        _retObj.Status = true;
//                        _retObj.StatusMessage = $"Report has been queued for sending to your Email Address.";

//                        return StatusCode(202, _retObj);
//                    }
//                }
//                else
//                {
//                    _retObj.Status = false;
//                    _retObj.StatusMessage = $"{_errMsg} Wrong Delivery Type";

//                    return StatusCode(202, _retObj);
//                }
//            }
//            catch (Exception ex)
//            {
//                _retObj.Status = false;
//                _retObj.StatusMessage = $"{_errMsg} {ex.Message}";
//                return StatusCode(202, _retObj);
//            }
//        }

//        [NonAction]
//        private Stream ProcessDataViewForExcel(List<SubmissionModel> lsReportData)
//        {
//            if (lsReportData.Count < 1)
//            {
//                return null;
//            }
//            try
//            {
//                var memoryStream = new MemoryStream();

//                using (var p = new ExcelPackage())
//                {
//                    var user = User.Identity.Name;
//                    var today = DateTime.Now.ToString("dddd, dd MMMM, yyyy");
//                    var time = DateTime.Now.ToString("hh:mm tt");
//                    var ws = p.Workbook.Worksheets.Add($"Complaint REPORT");
//                    ws.Cells["A1"].Value = $"DETAILED COMPLAINT REPORT Generated On {today} At {time} By {user}";
//                    ws.Cells["A1:M1"].Merge = true;

//                    ws.Cells[2, 1].Value = "S/N";
//                    ws.Cells[2, 2].Value = "REF";
//                    ws.Cells[2, 3].Value = "DATE CREATED";
//                    ws.Cells[2, 4].Value = "COMPLAINANT";
//                    ws.Cells[2, 5].Value = "OPERATOR";
//                    ws.Cells[2, 6].Value = "STATUS OF FIRM";
//                    ws.Cells[2, 7].Value = "COMPLAINT TYPE";
//                    ws.Cells[2, 8].Value = "STOCK";
//                    ws.Cells[2, 9].Value = "VOLUME";
//                    ws.Cells[2, 10].Value = "STATUS";
//                    ws.Cells[2, 11].Value = "DATE OF RESOLUTION";
//                    ws.Cells[2, 12].Value = "REMARK (JUSTIFIED OR NOT JUSTIFIED)";
//                    ws.Cells[2, 13].Value = "DEPARTMENT HANDLING THE COMPLAINT";
//                    int strtRw = 3;

//                    for (int i = 0; i < lsReportData.Count; i++)
//                    {
//                        var sub = lsReportData[i];

//                        ws.Cells[strtRw, 1].Value = i + 1;
//                        ws.Cells[strtRw, 2].Value = sub.Id;
//                        ws.Cells[strtRw, 3].Value = sub.DateCreated.Date.ToString("dddd, dd MMMM, yyyy") ?? "";
//                        ws.Cells[strtRw, 4].Value = $"{sub.Surname} {sub.FirstName}";
//                        //string dmf = "";
//                        //string issr = "";
//                        //string othersEnt = "";
//                        //for (int j = 0; j < sub.Details.Count; j++)
//                        //{
//                        //    var rxnDet = sub.Details[j];
//                        //    switch (rxnDet.EntityTypeEnum)
//                        //    {
//                        //        case EntityTypeEnum.DealingMember:
//                        //            if (!string.IsNullOrWhiteSpace(rxnDet.FirmName))
//                        //            {
//                        //                dmf += rxnDet.FirmName?.Trim() + ";";

//                        //            }

//                        //            break;
//                        //        case EntityTypeEnum.Issuer:
//                        //            if (!string.IsNullOrWhiteSpace(rxnDet.FirmName))
//                        //            {
//                        //                issr += rxnDet.FirmName?.Trim() + ";";
//                        //            }

//                        //            break;
//                        //        case EntityTypeEnum.Others:
//                        //            if (!string.IsNullOrWhiteSpace(rxnDet.FirmName))
//                        //            {
//                        //                othersEnt += rxnDet.FirmName?.Trim() + ";";
//                        //            }

//                        //            break;
//                        //        default:
//                        //            break;
//                        //    }
//                        //}
//                        ws.Cells[strtRw, 5].Value = sub.Dmf;
//                        ws.Cells[strtRw, 6].Value = sub.LicenseStatus;
//                        ws.Cells[strtRw, 7].Value = sub.SubmissionType;
//                        ws.Cells[strtRw, 8].Value = sub.ShareName;
//                        ws.Cells[strtRw, 9].Value = sub.ShareVolume;
//                        ws.Cells[strtRw, 10].Value = sub.SubmissionStatus;
//                        ws.Cells[strtRw, 11].Value = sub.DateLastModified.Date.ToString("dddd, dd MMMM, yyyy") ?? "";
//                        ws.Cells[strtRw, 12].Value = "";

//                        ws.Cells[strtRw, 13].Value = sub.SubmissionOwnerEnum;
//                        strtRw++;
//                    }

//                    using (ExcelRange r = ws.Cells[$"A1:A{strtRw}"])
//                    {
//                        r.Style.Font.Bold = true;
//                        r.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
//                        r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
//                    }
//                    using (ExcelRange r = ws.Cells[$"A2:M{strtRw}"])
//                    {

//                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
//                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;
//                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
//                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;

//                        r.Style.Border.BorderAround(ExcelBorderStyle.Thick);
//                    }

//                    using (ExcelRange r = ws.Cells["A2:M2"])
//                    {
//                        r.Style.Font.Bold = true;
//                        r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

//                        r.Style.Fill.PatternType = ExcelFillStyle.Solid;
//                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(178, 190, 181));
//                        r.Style.Font.Color.SetColor(Color.Black);
//                        r.Style.Font.Bold = true;
//                    }
//                    ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells

//                    using (ExcelRange r = ws.Cells[$"I3:S{strtRw}"])
//                    {
//                        r.Style.Numberformat.Format = "0.00";
//                    }
//                    p.SaveAs(memoryStream);
//                }

//                return memoryStream;
//            }
//            catch (Exception ex)
//            {
//                return null;
//            }
//        }
//        [NonAction]
//        private Stream ProcessDataForExcel(DateTime strtDt, DateTime endDt, List<SubmissionModel> lsReportData)
//        {
//            try
//            {
//                var memoryStream = new MemoryStream();

//                //Creates a blank workbook. Use the using statment, so the package is disposed when we are done.
//                using (var p = new ExcelPackage())
//                {
//                    var user = User.Identity.Name;
//                    var today = DateTime.Now.ToString("dddd, dd MMMM, yyyy");
//                    var time = DateTime.Now.ToString("hh:mm tt");
//                    var ws = p.Workbook.Worksheets.Add($"Complaint REPORT");
//                    ws.Cells["A1"].Value = $"COMPLAINT REPORT SUMMARY Generated On {today} At {time} By {user}";
//                    ws.Cells["A1:M1"].Merge = true;
//                    ws.Cells[2, 1].Value = "S/N";
//                    ws.Cells[2, 2].Value = "Ref";
//                    ws.Cells[2, 3].Value = "Date Received";
//                    ws.Cells[2, 4].Value = "Complaint Type";
//                    ws.Cells[2, 5].Value = "Dealing Member";
//                    ws.Cells[2, 6].Value = "Issuer";
//                    ws.Cells[2, 7].Value = "Others";
//                    ws.Cells[2, 8].Value = "Description";

//                    int strtRw = 3;
//                    for (int i = 0; i < lsReportData.Count; i++)
//                    {
//                        var sub = lsReportData[i];

//                        ws.Cells[strtRw, 1].Value = i + 1;
//                        ws.Cells[strtRw, 2].Value = sub.Id;
//                        ws.Cells[strtRw, 3].Value = sub.DateCreated.ToShortDateString();
//                        ws.Cells[strtRw, 4].Value = sub.SubmissionType;
//                        string dmf = "";
//                        string issr = "";
//                        string othersEnt = "";

//                        for (int j = 0; j < sub.Details.Count; j++)
//                        {
//                            var rxnDet = sub.Details[j];
//                            switch (rxnDet.EntityTypeEnum)
//                            {
//                                case EntityTypeEnum.DealingMember:
//                                    if (!string.IsNullOrWhiteSpace(rxnDet.FirmName))
//                                    {
//                                        dmf += rxnDet.FirmName?.Trim() + ";";
//                                    }

//                                    break;
//                                case EntityTypeEnum.Issuer:
//                                    if (!string.IsNullOrWhiteSpace(rxnDet.FirmName))
//                                    {
//                                        issr += rxnDet.FirmName?.Trim() + ";";
//                                    }

//                                    break;
//                                case EntityTypeEnum.Others:
//                                    if (!string.IsNullOrWhiteSpace(rxnDet.FirmName))
//                                    {
//                                        othersEnt += rxnDet.FirmName?.Trim() + ";";
//                                    }

//                                    break;
//                                default:
//                                    break;
//                            }
//                        }
//                        ws.Cells[strtRw, 5].Value = dmf;
//                        ws.Cells[strtRw, 6].Value = issr;
//                        ws.Cells[strtRw, 7].Value = othersEnt;

//                        var lsRxn = sub.Reactions.FindLast(r => r.ReactionSrc == SubmissionReactionSourceEnum.Exchange);
//                        ws.Cells[strtRw, 8].Value = lsRxn?.StatusComment ?? "";

//                        strtRw++;
//                    }

//                    using (ExcelRange r = ws.Cells[$"A1:A{strtRw}"])
//                    {
//                        r.Style.Font.Bold = true;
//                        r.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
//                        r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
//                    }
//                    using (ExcelRange r = ws.Cells[$"A2:M{strtRw}"])
//                    {

//                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;
//                        r.Style.Border.Right.Style = ExcelBorderStyle.Thin;
//                        r.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
//                        r.Style.Border.Top.Style = ExcelBorderStyle.Thin;

//                        r.Style.Border.BorderAround(ExcelBorderStyle.Thick);
//                    }

//                    using (ExcelRange r = ws.Cells["A2:M2"])
//                    {
//                        r.Style.Font.Bold = true;
//                        r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

//                        r.Style.Fill.PatternType = ExcelFillStyle.Solid;
//                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(178, 190, 181));
//                        r.Style.Font.Color.SetColor(Color.Black);
//                        r.Style.Font.Bold = true;
//                    }
//                    ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells


//                    p.SaveAs(memoryStream);
//                }

//                return memoryStream;
//            }
//            catch (Exception ex)
//            {
//                return null;
//            }
//        }


//        [NonAction]
//        public Task RunNotif(byte[] fileArray, string username, string useremail)
//        {
//            MemoryStream ms = new MemoryStream(fileArray);

//            _opNotifRepo.FileRenditionReportNotification(ms, "Report.xlsx", username, useremail, new List<EmailAddress>());

//            return Task.FromResult(0);
//        }

//    }
//}
