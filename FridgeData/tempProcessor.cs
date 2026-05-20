using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace FridgeData
{
    public class tempProcessor
    {
        
        public static async Task<temp> LoadInformation(string a, string b )
        {
            
            
            string url = "http://"+a+":"+b+"/api/v1/school/status";
            
            

            using (HttpResponseMessage reponse = await ApiHelper.client.GetAsync(url))
            {
                

                if (reponse.IsSuccessStatusCode)
                {
                    tempResults TEMPERATURE_MAIN = await reponse.Content.ReadAsAsync<tempResults>();

                    return TEMPERATURE_MAIN.TEMPERATURE_MAIN;





                }
                else
                {
                    throw new Exception(reponse.ReasonPhrase);
                }

            }




        }

    }
}
