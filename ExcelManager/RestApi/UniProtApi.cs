using System;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Duck.OfficeAutomationModule.RestApi
{
    static public class UniProtApi // 정적 클래스 VS 그냥 클래스 고민...
    {
        static readonly string PROTEINS_URL = "https://www.ebi.ac.uk/proteins/api/proteins";
        static readonly string COORDINATES_URL = "https://www.ebi.ac.uk/proteins/api/coordinates";

        private static readonly HttpClient _client = new HttpClient();

        public static async Task<string> GetProteinDataOrNullAsync(string accession)
        {
            string apiUrl = $"{PROTEINS_URL}?accession={accession}";
            HttpResponseMessage response = await _client.GetAsync(apiUrl);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }
            return null;
        }

        public static async Task<string> GetCoordinatesDataOrNullAsync(string accession)
        {
            string apiUrl = $"{COORDINATES_URL}?accession={accession}";
            HttpResponseMessage response = await _client.GetAsync(apiUrl);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }
            return null;
        }

        public static async Task<string> GetFilteredProteinDataOrNullAsync(string accession)
        {
            string apiUrl = $"{PROTEINS_URL}?accession={accession}"; // TODO: URL에 조건을 넣어 필요한 data만 가져오도록 수정
            HttpResponseMessage response = await _client.GetAsync(apiUrl);

            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }
            return null;
        }
    }
}
