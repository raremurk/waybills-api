using WaybillsAPI.Interfaces;

namespace WaybillsAPI.Services
{
    public class SalaryPeriodService : ISalaryPeriodService
    {
        public (int Year, int Month) GetSalaryPeriod()
        {
            var currentDate = DateTime.UtcNow;
            if (currentDate.Day < 16)
            {
                currentDate = currentDate.AddMonths(-1);
            }
            return (currentDate.Year, currentDate.Month);
        }

        public DateOnly GetMaxWaybillDate()
        {
            var (Year, Month) = GetSalaryPeriod();
            return new DateOnly(Year, Month, DateTime.DaysInMonth(Year, Month));
        }
    }
}
