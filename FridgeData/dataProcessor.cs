using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
namespace FridgeData
{
    public class dataProcessor
    {
        public static async Task<datas> LoadInformation(string a, string b)
        {
            string url = "http://"+a+":"+b+"/api/v1/school/status";



            using (HttpResponseMessage reponse = await ApiHelper.client.GetAsync(url))
            {

                if (reponse.IsSuccessStatusCode)
                {
                    datas result = await reponse.Content.ReadAsAsync<datas>();
                   
                    return result;
                    
                    

                    

                }
                else
                {
                    throw new Exception(reponse.ReasonPhrase);
                }

            }
            



        }

        public static async Task<datas> LoadInformation2()
        {
            string url = "http://192.168.31.250:56000/api/v1/school/status";



            using (HttpResponseMessage reponse = await ApiHelper.client.GetAsync(url))
            {

                if (reponse.IsSuccessStatusCode)
                {
                   
                    TEMPERATURE_MAIN TEMPERATURE_MAIN = await reponse.Content.ReadAsAsync<TEMPERATURE_MAIN>();
                    return TEMPERATURE_MAIN.datas;





                }
                else
                {
                    throw new Exception(reponse.ReasonPhrase);
                }

            }




        }

    }
}
