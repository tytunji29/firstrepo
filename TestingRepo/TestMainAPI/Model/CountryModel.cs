using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TestAPI.Model
{
    public class CountryModel
    {

        public int Id { get; set; }
        public string Name { get; set; }
        public string Capital { get; set; }
        public string DialCode { get; set; }
        public string Currency { get; set; }
        public string Region { get; set; }
        public string Subregion { get; set; }
        public string Demonym { get; set; }
        public string TimeZone { get; set; }
        public DateTime DateCreated { get; set; }
    }

}
