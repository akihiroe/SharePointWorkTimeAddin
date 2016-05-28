using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Linq.Expressions;
using Microsoft.LightSwitch.Security.Server;
using Microsoft.LightSwitch;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Diagnostics;
using SharePoint.WorkTimeAddin.SpreadsheetML;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LightSwitchApplication
{
    public class ExcelManager
    {
        const string emailAddressCell = "S6";
        const string startDateCell = "A9";
        const int startPos = 9;
        const string sheetName = "作業時間";

        public void Import(string filename)
        {
            using (var helper = new SpreadsheetHelper(filename))
            {
                if (!helper.MoveWorksheet(sheetName)) throw new ApplicationException("不正なテンプレートです。" + sheetName + "のシートがありません");
                var worksheet = helper.CurrentSheet;

                using (var serverContext = ServerApplicationContext.CreateContext())
                {
                    var date = (DateTime)helper.GetCellValue(startDateCell, CellValues.Date);
                    var userId = helper.GetCellValue(emailAddressCell) as string;
                    if (string.IsNullOrEmpty(userId)) throw new ApplicationException("ユーザが指定されていません");
                    var sb = new StringBuilder();
                    for (int rowIdx = startPos; rowIdx < startPos + 31; rowIdx++)
                    {
                        using (var workspace = serverContext.Application.CreateDataWorkspace())
                        {
                            try
                            {
                                string sickHoliday = helper.GetCellValue(4, rowIdx) as string;
                                var startTimeObject = helper.GetCellValue(9, rowIdx);
                                string startTime = startTimeObject is DateTime ? ((DateTime)startTimeObject).ToString("HH:mm") : startTimeObject as string;
                                var endTimeObject = helper.GetCellValue(11, rowIdx);
                                string endTime = endTimeObject is DateTime ? ((DateTime)endTimeObject).ToString("HH:mm") : endTimeObject as string;
                                var item = workspace.ApplicationData.WorkTimeSet.AddNew();
                                string remark = helper.GetCellValue(16, rowIdx) as string;

                                if (string.IsNullOrEmpty(startTime) && string.IsNullOrEmpty(endTime) && string.IsNullOrEmpty(remark) && string.IsNullOrEmpty(sickHoliday)) continue;

                                item.UserId = userId;
                                item.SickHolidy = sickHoliday;
                                item.WorkDate = date.AddDays(rowIdx - startPos);
                                item.StartTime = startTime;
                                item.EndTime = endTime;
                                item.Remark = remark;
                                workspace.ApplicationData.SaveChanges();
                            }
                            catch (Microsoft.LightSwitch.ValidationException vex)
                            {
                                foreach (var error in vex.ValidationResults)
                                {
                                    sb.AppendLine(date.AddDays(rowIdx - startPos).ToString("d") + ":" + error.Message);
                                }

                            }
                        }
                    }
                    if (sb.Length > 0)
                    {
                        throw new ApplicationException(sb.ToString());
                    }
                }

            }
        }

        public void Export(string filename, string email, int year, int month)
        {
            using (var helper = new SpreadsheetHelper(filename))
            {
                if (!helper.MoveWorksheet(sheetName)) throw new ApplicationException("不正なテンプレートです。" + sheetName + "のシートがありません");
                var worksheet = helper.CurrentSheet;
                helper.SetCellValue(emailAddressCell,email);
                using (var serverContext = ServerApplicationContext.CreateContext())
                {
                    var startDate = new DateTime(year, month, 1);
                    helper.SetCellValue(startDateCell, startDate);
                    var sb = new StringBuilder();

                    using (var workspace = serverContext.Application.CreateDataWorkspace())
                    {

                        foreach (WorkTime item in workspace.ApplicationData.WorkTimeSet.Where(x => x.UserId == email && x.WorkDate.Year == year && x.WorkDate.Month == month))
                        {
                            var rowIdx = item.WorkDate.Day - 1 + startPos;
                            helper.SetCellValue(4,rowIdx, item.SickHolidy);
                            helper.SetCellValue(9, rowIdx, item.StartTime);
                            helper.SetCellValue(11, rowIdx, item.EndTime);
                            helper.SetCellValue(16, rowIdx, item.Remark);
                        }
                    }
                }
                helper.Save(filename);
            }

        }

        public void Submit(string filename, string email, int year, int month)
        {
            using (var serverContext = ServerApplicationContext.CreateContext())
            {

                using (var workspace = serverContext.Application.CreateDataWorkspace())
                {
                    var entried = workspace.ApplicationData.WorkTimeSet.Where(x => x.UserId == email && x.WorkDate.Year == year && x.WorkDate.Month == month).Select(x => x.WorkDate).Execute().ToList();
                    var holidays = HolidyManager.GenerateHoliday(year);
                    foreach (var day in Enumerable.Range(1, DateTime.DaysInMonth(year, month)))
                    {
                        var targetDay = new DateTime(year, month, day);
                        if (targetDay.DayOfWeek == DayOfWeek.Sunday || targetDay.DayOfWeek == DayOfWeek.Saturday) continue;
                        if (holidays.ContainsKey(targetDay)) continue;
                        if (!entried.Contains(targetDay))
                        {
                            throw new ApplicationException(targetDay.ToString("d") + "の入力がありません。");
                        }
                    }

                }
            }

            Export(filename, email, year, month);

            using (var serverContext = ServerApplicationContext.CreateContext())
            {
                var appWebContext = serverContext.Application.SharePoint;

                using (var ctx = appWebContext.GetAppWebClientContext())
                {
                    var list = ctx.Web.Lists.GetByTitle("WorkTimeSheet");
             
                    var rootFolder = list.RootFolder;
                    ctx.Load(rootFolder, x=>x.Folders, x=>x.ServerRelativeUrl);
                    ctx.ExecuteQuery();
                    var subFolderName = year.ToString("0000") + month.ToString("00");
                    var subFolder = list.RootFolder.Folders.Where(x => x.Name == subFolderName).FirstOrDefault();
                    if (subFolder == null)
                    {
                        subFolder = rootFolder.Folders.Add(rootFolder.ServerRelativeUrl +"/" + subFolderName);
                        ctx.Load(subFolder);
                        ctx.ExecuteQuery();
                    }
                    using (var st = new FileStream(filename, FileMode.Open))
                    {
                        var info = new FileCreationInformation();
                        info.ContentStream = st;
                        info.Overwrite = true;
                        info.Url = subFolder.ServerRelativeUrl +"/" + email.Replace("@","_") + ".xlsx";
                        var file = subFolder.Files.Add(info);
                        ctx.ExecuteQuery();
                    }
                }
            }

        }
    }
}