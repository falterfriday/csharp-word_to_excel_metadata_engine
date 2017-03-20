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
            text = string.Join(" ", text.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));
            var checkBoxList = text.Split();
            for (int i = 0; i < checkBoxList.Length; i++)
            {
                if (checkBoxList[i].Contains("☐"))
                {
                    checkBoxList[i] = checkBoxList[i].Replace("☐", "<br />\n<input type=\"checkbox\" id=\"CB\"/> ");
                }
                if (checkBoxList[i].Contains("Assess"))
                {
                    checkBoxList[i] = checkBoxList[i].Replace("Assess", "<br />\nAssess");
                }
            }
            text = string.Join(" ", checkBoxList);
            return text;
        }
    }
}
