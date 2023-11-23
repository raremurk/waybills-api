namespace WaybillsAPI.Models
{
    public class Driver
    {
        public int Id { get; set; }

        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public int PersonnelNumber { get; set; }

        public List<Waybill> Waybills { get; set; }

        public string ShortFullName()
        {
            var firstNameInitial = string.IsNullOrEmpty(FirstName) ? "" : $" {FirstName[0]}. ";
            var middleNameInitial = string.IsNullOrEmpty(MiddleName) ? "" : $"{MiddleName[0]}.";
            return string.Concat(LastName, firstNameInitial, middleNameInitial);
        }
    }
}
