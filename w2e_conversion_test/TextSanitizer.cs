using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace w2e_conversion_test
{
    static class TextSanitizer
    {
        public static string InstructionSanitizer(string text)
        {
            text.Replace("☐", "<input type=\"checkbox\"/> ");
            return text;
        }
    }
}
