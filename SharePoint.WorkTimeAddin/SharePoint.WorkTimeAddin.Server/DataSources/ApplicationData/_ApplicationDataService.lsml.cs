using System.Linq.Expressions;
using Microsoft.LightSwitch.Security.Server;
using Microsoft.LightSwitch;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System;

namespace LightSwitchApplication
{
    public partial class ApplicationDataService
    {
        partial void WorkTimeSet_Inserting(WorkTime entity)
        {
            entity.WorkDate = entity.WorkDate.Date;
            var replaceEntity = this.WorkTimeSet.Where(x => x.UserId == entity.UserId && x.WorkDate == entity.WorkDate).FirstOrDefault();
            if (replaceEntity != null)
            {
                replaceEntity.Delete();
            }
        }

        partial void CurrentUserWorkTime_PreprocessQuery(int? year, int? month, ref IQueryable<WorkTime> query)
        {
            query = query.Where(x => x.UserId == this.Application.User.Email);
            if (year.HasValue && month.HasValue)
            {
                query = query.Where(x => x.WorkDate.Year == year.Value && x.WorkDate.Month == month.Value);
            }
        }

        partial void WorkTimeSet_Validate(WorkTime entity, EntitySetValidationResultsBuilder results)
        {
            DateTime startTime;
            DateTime endTime;
            if (!string.IsNullOrEmpty(entity.StartTime) && !DateTime.TryParseExact(entity.StartTime, "HH:mm", null, System.Globalization.DateTimeStyles.None, out startTime))
            {
                results.AddPropertyError("開始時刻形式が不正", entity.Details.Properties.StartTime);
            }
            if (!string.IsNullOrEmpty(entity.EndTime) && !DateTime.TryParseExact(entity.EndTime, "HH:mm", null, System.Globalization.DateTimeStyles.None, out endTime))
            {
                results.AddPropertyError("終了時刻形式が不正", entity.Details.Properties.EndTime);
            }
        }
    }
}