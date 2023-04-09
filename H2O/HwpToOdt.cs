using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public class HwpToOdt
    {
        public HwpToOdt()
        {
        }

        public void Convert()
        {
            //HwpToJson
            HwpToJson hj = new HwpToJson();
            hj.Run();

            //JsonToOdt
            JsonToOdt jo = new JsonToOdt();
            jo.setJson(hj.getJson());

            jo.Run();

            string[] files = new string[] { "content.xml", "settings.xml", "styles.xml" };
            foreach (string filename in files)
            {
                string rewrite = System.IO.File.ReadAllText(Application.StartupPath + @"\New File\" + filename);
                int index = rewrite.IndexOf("\n");
                string subString1 = rewrite.Substring(0, index + 1);
                string subString2 = rewrite.Substring(index + 1);
                rewrite = subString1 + Regex.Replace(subString2, @">\r\n( )*<", "><");
                //  (, ) 가 json에 (,\n+공백)로 줄바꿈 입력됨 이유는 모름 나중에 수정가능성
                rewrite = Regex.Replace(rewrite, @",\n             ", ", ");
                System.IO.File.WriteAllText(Application.StartupPath + @"\New File\" + filename, rewrite);
            }
        }
    }
}