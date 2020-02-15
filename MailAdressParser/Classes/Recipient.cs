using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailAdressParser
{
    /// <summary>
    /// получатель писем
    /// </summary>
    [Serializable]
    public class Recipient
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public string Company { get; set; }
        public string Occupation { get; set; }
        public string INN { get; set; }
        public string Phone { get; set; }

        public string Address { get; set; }

        public bool WasSent { get; set; }

        public DateTime? SendDateTime { get ; set ; }

    }
}
