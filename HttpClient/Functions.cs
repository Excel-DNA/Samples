using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace HttpClientSample
{
    public static class Functions
    {
        static HttpClient httpClient = new HttpClient();

        [ExcelFunction]
        public static async Task<string> httpGetString(string uri)
        {
            return await httpClient.GetStringAsync(uri);
        }
    }
}
