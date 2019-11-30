using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupProjecto
{
    class Holiday
    {
        public DateTime HolidayDate { get; set; }
        public string HolidayDescription { get; set; }

        public Holiday(DateTime holidaydate, string holidaydescription)
        {
            HolidayDate = holidaydate;
            HolidayDescription = holidaydescription;
        }
    }
}
