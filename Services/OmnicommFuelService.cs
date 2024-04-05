using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Omnicomm;
using WaybillsAPI.Omnicomm.OmnicommModels;

namespace WaybillsAPI.Services
{
    public class OmnicommFuelService : IOmnicommFuelService
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly string url = "https://online.omnicomm.ru/";
        private readonly string statisticsRoute = "ls/api/v1/reports/statistics";

        private readonly JsonSerializerOptions jsonSerializerOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

        private readonly string JWTFileNameUser1 = "OmnicommLoginData/JWTUser1.json";
        private readonly string JWTFileNameUser2 = "OmnicommLoginData/JWTUser2.json";
        private readonly JWT JWTUser1 = new();
        private readonly JWT JWTUser2 = new();

        private readonly string loginDataFileNameUser1 = "OmnicommLoginData/LoginDataUser1.json";
        private readonly string loginDataFileNameUser2 = "OmnicommLoginData/LoginDataUser2.json";
        private readonly LoginData loginData1 = new();
        private readonly LoginData loginData2 = new();

        public OmnicommFuelService(IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
            GetJWTFromFile(JWTFileNameUser1, JWTUser1);
            GetJWTFromFile(JWTFileNameUser2, JWTUser2);
            GetLoginDataFromFile(loginDataFileNameUser1, loginData1);
            GetLoginDataFromFile(loginDataFileNameUser2, loginData2);
        }

        public async Task<OmnicommFuelReport> GetOmnicommFuelReport(DateOnly date, int omnicommId)
        {
            var jwt = omnicommId < 1000000000 ? JWTUser1 : JWTUser2;
            var userId = omnicommId < 1000000000 ? 1 : 2;
            var startUnix = new DateTimeOffset(date, TimeOnly.MinValue, TimeSpan.Zero).ToUnixTimeSeconds();
            var endUnix = new DateTimeOffset(date, TimeOnly.MaxValue, TimeSpan.Zero).ToUnixTimeSeconds();

            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("JWT", jwt.Jwt);
            var response = await client.GetAsync(url + statisticsRoute + $"?timeBegin={startUnix}&timeEnd={endUnix}&dataGroups=[fuel]&vehicles=[{omnicommId}]");

            if (response.StatusCode == HttpStatusCode.Unauthorized)
            {
                await GetActualJWT(userId, jwt);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("JWT", jwt.Jwt);
                response = await client.GetAsync(url + statisticsRoute + $"?timeBegin={startUnix}&timeEnd={endUnix}&dataGroups=[fuel]&vehicles=[{omnicommId}]");
            }

            return await response.Content.ReadFromJsonAsync<OmnicommFuelReport>();
        }

        private async Task GetActualJWT(int userId, JWT jwt)
        {
            var loginData = userId == 1 ? loginData1 : loginData2;
            var fileName = userId == 1 ? JWTFileNameUser1 : JWTFileNameUser2;
            HttpResponseMessage response;
            if (jwt.Refresh != "")
            {
                response = await RefreshJWT(jwt.Refresh);
                if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    response = await Authenticate(loginData);
                }
            }
            else
            {
                response = await Authenticate(loginData);
            }

            var responseBody = await response.Content.ReadAsStringAsync();
            var jwtFromResponse = JsonSerializer.Deserialize<JWT>(responseBody, jsonSerializerOptions);
            if (jwtFromResponse != null)
            {
                jwt.Jwt = jwtFromResponse.Jwt;
                jwt.Refresh = jwtFromResponse.Refresh;
                File.WriteAllText(fileName, responseBody);
            }
        }

        private Task<HttpResponseMessage> Authenticate(LoginData loginData)
        {
            var client = _httpClientFactory.CreateClient();
            return client.PostAsync($"{url}auth/login?jwt=1", JsonContent.Create(loginData));
        }

        private Task<HttpResponseMessage> RefreshJWT(string refreshJwt)
        {
            var client = _httpClientFactory.CreateClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("JWT", refreshJwt);
            return client.PostAsync($"{url}auth/refresh", null);
        }

        private void GetJWTFromFile(string fileName, JWT jwt)
        {
            try
            {
                var fileData = File.ReadAllText(fileName);
                var jwtFromFile = JsonSerializer.Deserialize<JWT>(fileData, jsonSerializerOptions);
                if (jwtFromFile != null)
                {
                    jwt.Jwt = jwtFromFile.Jwt;
                    jwt.Refresh = jwtFromFile.Refresh;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private void GetLoginDataFromFile(string fileName, LoginData loginData)
        {
            try
            {
                var fileData = File.ReadAllText(fileName);
                var loginDataFromFile = JsonSerializer.Deserialize<LoginData>(fileData, jsonSerializerOptions);
                if (loginDataFromFile != null)
                {
                    loginData.Login = loginDataFromFile.Login;
                    loginData.Password = loginDataFromFile.Password;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
