using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using System.Windows.Forms;

namespace WindowsFormsApp8
{
    class API_class
    {
        public string getInformationAPI (string erdpouuu)
        {
            string url = "https://clarity-project.info/api/edr.info/" + erdpouuu;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            string response;
                using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
                {
                    response = streamReader.ReadToEnd();
                }

            NameInformation nameInformation = JsonConvert.DeserializeObject<NameInformation>(response);

            string livePlace = nameInformation.address.postal + ", " + nameInformation.address.locality + ", " + nameInformation.address.address;
            //MessageBox.Show(livePlace);
            return livePlace;
        }

        public string getInformationAPIname(string erdpouuu)
        {
            string url = "https://clarity-project.info/api/edr.info/" + erdpouuu;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            string response;
            using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
            {
                response = streamReader.ReadToEnd();
            }

            NameInformation nameInformation = JsonConvert.DeserializeObject<NameInformation>(response);

            string nameFind = nameInformation.Name;
            //MessageBox.Show(nameFind);
            return nameFind;
        }

        public string getInformationAPIshortName (string erdpouuu)
        {
            string url = "https://clarity-project.info/api/edr.info/" + erdpouuu;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            string response;
            using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
            {
                response = streamReader.ReadToEnd();
            }

            NameInformation nameInformation = JsonConvert.DeserializeObject<NameInformation>(response);

            string addresses = nameInformation.edr_data.shortname;

            return addresses;
        }

        public string getInformationAPIaddressF (string erdpouuu)
        {
            string url = "https://clarity-project.info/api/edr.info/" + erdpouuu;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            string response;
            using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
            {
                response = streamReader.ReadToEnd();
            }

            NameInformation nameInformation = JsonConvert.DeserializeObject<NameInformation>(response);

            string addresses = nameInformation.edr_data.address_parts.full_address;

            return addresses;
        }

        public object getInformationDirector (string erdpouuu)
        {
            string url = "https://clarity-project.info/api/edr.info/" + erdpouuu;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            string response;
            using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
            {
                response = streamReader.ReadToEnd();
            }

            NameInformation nameInformation = JsonConvert.DeserializeObject<NameInformation>(response);

            var nameof = new Dictionary<string, string>();
            string firtst = "hell";
            string addresses = nameInformation.edr_data.address_parts.postal_code;

            string firtst2 = "hell2";
            string addresses2 = nameInformation.edr_data.address_parts.postal_code;

            string firtst3 = "hell3";
            string addresses3 = nameInformation.edr_data.address_parts.postal_code;


            nameof.Add(firtst, addresses);
            nameof.Add(firtst2, addresses2);
            nameof.Add(firtst3, addresses3);

            //MessageBox.Show(nameof[firtst]);

            return nameof;
        }
    }

    public class NameInformation
    {
        public string Name { get; set; }
        public string Edr { get; set; }
        public Address address { get; set; }
        public Edr_data edr_data { get; set; }
    }

    public class Edr_data
    {
        public string shortname { get; set; }
        public Address_parts address_parts { get; set; } 
    }

    public class Address_parts
    {
        public string full_address { get; set; }
        public string postal_code { get; set; }
    }

    public class Address
    {
        public string locality { get; set; }
        public string address { get; set; }
        public string postal { get; set; }
    }
}
