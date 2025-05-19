using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMergeTestApp
{
    internal class DataHelper
    {
        public static List<Dictionary<string, string>> GetRecepientData()
        {
            return new List<Dictionary<string, string>>()
            {
                new Dictionary<string, string>
                {
                    { "Name", "Anjum Rizwi" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "382, 6th Cross, SK Garden, Bengaluru" },
                    { "Company", "Valtech India" }
                },
                new Dictionary<string, string>
                {
                    { "Name", "Arun" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "382, 6th Cross, JP Nagar, Bengaluru" },
                    { "Company", "Valtech India" }
                },
                new Dictionary<string, string>
                {
                    { "Name", "John Doe" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "123 Test Street, Bengaluru" },
                    { "Company", "Valtech US" }
                }
            };
        }
    }
}
