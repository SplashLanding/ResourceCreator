using FileHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Resource_Creator
{
    //Went with '&' because commas are aboundent in languages. 
    [DelimitedRecord("|")]
    public class TranslationModel
    {        
        public string Name { get; set; }
        public string Value { get; set; }
    }
}
