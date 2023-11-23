using WaybillsAPI.Models;

namespace WaybillsAPI.Context
{
    public class Data
    {
        public static List<Driver> Drivers() =>
        [
            new Driver { Id = 1, FirstName = "Василий", MiddleName = "Васильевич", LastName = "Богатко", PersonnelNumber = 5008 },
            new Driver { Id = 2, FirstName = "Виктор", MiddleName = "Викторович", LastName = "Куделич", PersonnelNumber = 1585 }
        ];


        public static List<Transport> Transports() =>
        [
            new Transport { Id = 1, Name = "МТЗ-82 №6", Code = 5937, Coefficient = 4.9 },
            new Transport { Id = 2, Name = "МТЗ-922.3 №47", Code = 5990, Coefficient = 4.9 }
        ];
    }
}
