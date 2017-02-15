using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace w2e_conversion_test
{
    public class InstructionSanitizer
    {
        private string text;

        public string InstructionSanitizer(string text)
        {
            text = "This is a testy test";
            return text;
            //this.text = text;
        }
    }
}
