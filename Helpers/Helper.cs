using System.Text.RegularExpressions;

#pragma warning disable SYSLIB1045

namespace WaybillsAPI.Helpers
{
    public static class Helper
    {
        public static string PadNumbers(string input)
        {
            return Regex.Replace(input, "[0-9]+", match => match.Value.PadLeft(10, '0'));
        }
    }
}
