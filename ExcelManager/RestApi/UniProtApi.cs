using System;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Duck.OfficeAutomationModule.RestApi
{
    static public class UniProtApi // 정적 클래스 VS 그냥 클래스 고민...
    {
        static readonly string PROTEINS_URL = "https://www.ebi.ac.uk/proteins/api/proteins";
        static readonly string COORDINATES_URL = "https://www.ebi.ac.uk/proteins/api/coordinates";

        private static readonly HttpClient CLIENT = new HttpClient();

        public static async Task<string> GetProteinDataOrNullAsync(string accession)
        {
            Debug.Assert(accession != null);

            try
            {
                string apiUrl = $"{PROTEINS_URL}?accession={accession}";
                HttpResponseMessage response = await CLIENT.GetAsync(apiUrl);

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }


        public static async Task<string> GetCoordinatesDataOrNullAsync(string accession)
        {
            Debug.Assert(accession != null);

            try
            {
                string apiUrl = $"{COORDINATES_URL}?accession={accession}";
                HttpResponseMessage response = await CLIENT.GetAsync(apiUrl);

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }

        public static async Task<string> GetFilteredProteinDataOrNullAsync(string accession)
        {
            Debug.Assert(accession != null);

            try
            {
                string apiUrl = $"{PROTEINS_URL}?accession={accession}";
                HttpResponseMessage response = await CLIENT.GetAsync(apiUrl);

                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }
    }
}
