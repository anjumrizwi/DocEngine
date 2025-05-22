using Bogus.DataSets;
using NameGenerator.Generators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace DocEngine.Data
{
    public class DataHelper
    {
        public static List<Dictionary<string, string>> GetRecepientData(int num)
        {
            var lstRecepient = new List<Dictionary<string, string>>();
            RealNameGenerator Generator = new RealNameGenerator();
            
            for (int i = 0; i < num; i++)
            {
                var recepient =  new Dictionary<string, string>
                {
                    { "«First»", Generator.Generate() },
                    { "«Last»", Generator.Generate() },
                    { "«Title»", Generator.Generate() },
                    { "«Department»", "Development" },
                    { "«Company»", "Valtech India" },
                    { "«Display_name»", "Valtech US" },
                    { "«Address»", "Orlando" },
                    { "«Business_Address_City»", "Cichago" }
                   
                };

                lstRecepient.Add(recepient);
            }

            return lstRecepient;
           
        }
        public static List<Dictionary<string, string>> GetRecepientData()
        {
            return new List<Dictionary<string, string>>()
            {
                new Dictionary<string, string>
                {
                    { "FullName", "Anjum Rizwi" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "382, 6th Cross, SK Garden, Bengaluru" },
                    { "Company", "Valtech India" }
                },
                new Dictionary<string, string>
                {
                    { "FullName", "Arun" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "382, 6th Cross, JP Nagar, Bengaluru" },
                    { "Company", "Valtech India" }
                },
                new Dictionary<string, string>
                {
                    { "FullName", "John Doe" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "123 Test Street, Bengaluru" },
                    { "Company", "Valtech US" }
                }
            };
        }
    }
}
