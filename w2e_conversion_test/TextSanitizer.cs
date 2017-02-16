using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace w2e_conversion_test
{
    static class TextSanitizer
    {
        public static string StandardSanitizer(string text)
        {
            int colonIdx = text.IndexOf(":");
            text = "<b>STANDARD:</b> " + (text.Substring(colonIdx + 2)).Trim();
            return text;
        }

        public static string InstructionSanitizer(string text)
        {
            var list = text.Split();
            List<string> cbList = new List<string>();
            for (int i = 0; i < list.Length; i++)
            {
                if (list[i].Contains("☐"))
                {
                    list[i] = list[i].Replace("☐", "\n<input type=\"checkbox\"/> ");
                }
            }
            text = string.Join(" ", list);
            return text;
        }
    }
}
