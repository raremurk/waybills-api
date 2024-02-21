using WaybillsAPI.Interfaces;

namespace WaybillsAPI.Services
{
    public class DateService : IDateService
    {
        public (int, int) GetSalaryPeriod()
        {
            var currentDate = DateOnly.FromDateTime(DateTime.UtcNow);
            if (currentDate.Day < 16)
            {
                currentDate = currentDate.AddMonths(-1);
            }
            return (currentDate.Year, currentDate.Month);
        }
    }
}
