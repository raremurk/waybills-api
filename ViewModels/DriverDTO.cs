using WaybillsAPI.Models;

namespace WaybillsAPI.ViewModels
{
    public class DriverDTO
    {
        public int Id { get; init; }
        public string FirstName { get; init; } = "";
        public string MiddleName { get; init; } = "";
        public string LastName { get; init; } = "";
        public int PersonnelNumber { get; init; }

        private DriverDTO() { }
        public DriverDTO(Driver driver)
        {
            ArgumentNullException.ThrowIfNull(driver);
            Id = driver.Id;
            FirstName = driver.FirstName;
            MiddleName = driver.MiddleName;
            LastName = driver.LastName;
            PersonnelNumber = driver.PersonnelNumber;
        }
    }
}
