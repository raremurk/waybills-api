namespace WaybillsAPI.Interfaces
{
    public interface ISalaryPeriodService
    {
        public (int Year, int Month) GetSalaryPeriod();
        public DateOnly GetMaxWaybillDate();
    }
}
