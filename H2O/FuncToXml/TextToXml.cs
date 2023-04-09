using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace WindowsFormsApp1.FuncToXml
{
    public class TextToXml
    {
        public TextToXml()
        {
        }

        public int bitcal(int i, int shift, byte b) //대부분 property에 있는 value를 bit별로 나누어서 넣을때 
        /** i = value 값  shift = 이동갯수 b = &연산할 값 */
        {
            byte temp = (byte)(i >> shift);
            return (int)(temp & b);
        }

        public void Run(XmlManager xm, JToken json, JToken docJson, bool zeroCheck)
        {

            // 문단 ID 가져오기
            int shapeId = json["header"]["paraShapeId"].Value<int>();
            int sID = shapeId;

            //전체 텍스트 가져오기
            string pcontent = ""; // p text
            try
            {
                object obj = json["text"].Value<object>();

                if (obj != null)
                    pcontent = json["text"].Value<string>();
                if (pcontent.Equals(""))
                    throw new Exception();
            }
            /////////////////////////////////////////////////////////////글자가 없을 때
            catch (Exception e)
            {
                int nullPStyle = json["charShape"]["positionShapeIdPairList"][0]["shapeId"].Value<int>();
                float nullPFontSize = docJson["charShapeList"][nullPStyle]["baseSize"].Value<float>() / 100;
                string nullPName = xm.AddContentP(); // para null 인 경우 처리
                int property = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();
                xm.SetFontSize(nullPName, nullPFontSize);

                //줄 간격
                int lineSpace = docJson["paraShapeList"][sID]["lineSpace"].Value<int>();
                xm.Paragraph.SetLineSpace(nullPName, (XmlDocument)xm.docs["content.xml"], lineSpace, property, (int)nullPFontSize * 100);

                //문단 테두리 간격
                double topborderSpace = docJson["paraShapeList"][sID]["topBorderSpace"].Value<double>();
                double bottomBorderSpace = docJson["paraShapeList"][sID]["bottomBorderSpace"].Value<double>();
                double leftBorderSpace = docJson["paraShapeList"][sID]["leftBorderSpace"].Value<double>();
                double rightBorderSpace = docJson["paraShapeList"][sID]["rightBorderSpace"].Value<double>();

                topborderSpace *= 0.01;
                bottomBorderSpace *= 0.01;
                leftBorderSpace *= 0.01;
                rightBorderSpace *= 0.01;

                //xm.Paragraph.SetBorderSpace(nullPName, (XmlDocument)xm.docs["content.xml"], topborderSpace, bottomBorderSpace, leftBorderSpace, rightBorderSpace);

                //정렬
                int temp = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();

                int paraAlign = bitcal(temp, 2, 0x7);
                if (paraAlign.Equals(2))
                    xm.SetPAlign(nullPName, "end");
                else if (paraAlign.Equals(0))
                    xm.SetPAlign(nullPName, "justify");
                else if (paraAlign.Equals(3))
                    xm.SetPAlign(nullPName, "center");
                else if (paraAlign.Equals(4))
                    xm.SetPAlign(nullPName, "Divide");
                else if (paraAlign.Equals(5))
                    xm.SetPAlign(nullPName, "Distribute");
                else if (paraAlign.Equals(1))
                    xm.SetPAlign(nullPName, "Distribute");

                //첫줄 들여쓰기
                //margin
                double indent = docJson["paraShapeList"][sID]["indent"].Value<double>();
                double topspace = docJson["paraShapeList"][sID]["topParaSpace"].Value<double>();
                double bottomspace = docJson["paraShapeList"][sID]["bottomParaSpace"].Value<double>();
                double leftmargin = docJson["paraShapeList"][sID]["leftMargin"].Value<double>();
                double rightmargin = docJson["paraShapeList"][sID]["rightMargin"].Value<double>();
                if (indent != 0 || topspace != 0 || bottomspace != 0 || leftmargin != 0 || rightmargin != 0)
                {
                    xm.SetPIndent(nullPName, (float)(indent / 200 * 0.0353));
                    if (indent < 0)
                        leftmargin = -indent + leftmargin;
                    xm.SetPMargin(nullPName, (float)(leftmargin / 200 * 0.0353), (float)(rightmargin / 200 * 0.0353), (float)(topspace / 200 * 0.0353), (float)(bottomspace / 200 * 0.0353));
                }

                return;
            }
            /////////////////////////////////////////////////////////////////

            // 스타일별 텍스트 개수
            int spancount = json["charShape"]["positionShapeIdPairList"].Count(); // p 안 text style count
                                                                                  // 스타일별 텍스트 아이디
            int pstyle = json["charShape"]["positionShapeIdPairList"][0]["shapeId"].Value<int>();

            int current_position;
            int next_position;

            string pname = "";
            string name;
            Boolean hasControl = false;
            int controlListCount = 0;
            // 텍스트별 위치 비교해서 자르기
            for (int j = 0; j < spancount; j++)
            {
                //스타일이 시작되는 위치
                current_position = json["charShape"]["positionShapeIdPairList"][j]["position"].Value<int>();

                string subcontent = "";
                if (j < spancount - 1)
                {
                    next_position = json["charShape"]["positionShapeIdPairList"][j + 1]["position"].Value<int>();
                    if (zeroCheck)
                    {
                        current_position = current_position - 16 < 0 ? 0 : current_position - 16;
                        next_position -= 16;
                    }
                    try
                    {
                        controlListCount = json["controlList"].Count();
                        hasControl = true;

                    }
                    catch (System.ArgumentNullException e) { hasControl = false; }
                    if (hasControl)
                        for (int i = 0; i < controlListCount; i++)
                        {
                            if (627600491 == json["controlList"][i]["header"]["ctrlId"].Value<int>() || 1651469165
 == json["controlList"][i]["header"]["ctrlId"].Value<int>())
                                subcontent = pcontent.Substring(current_position - 8 < 0 ? 0 : current_position - 8, current_position - 8 < 0 ? next_position - 8 - current_position : next_position - 8 - (current_position - 8)); //처음포지션부터 글자 수만큼 자르기
                            else subcontent = pcontent.Substring(current_position, next_position - current_position);
                        }
                    else subcontent = pcontent.Substring(current_position, next_position - current_position);                                                                             //
                }
                else //마지막 텍스트
                {
                    if (zeroCheck && j != 0)
                    {
                        current_position -= 16;
                    }

                    if (hasControl)
                        for (int i = 0; i < controlListCount; i++)
                        {
                            int count = json["charShape"]["positionShapeIdPairList"].Count() - 1;
                            if (627600491 == json["controlList"][i]["header"]["ctrlId"].Value<int>() || 1651469165
 == json["controlList"][i]["header"]["ctrlId"].Value<int>())
                                subcontent = pcontent.Substring(current_position - 8 * count); //처음포지션부터 글자 수만큼 자르기
                            else subcontent = pcontent.Substring(current_position);
                        }
                    else subcontent = pcontent.Substring(current_position);

                }

                //스타일 추가 및 내용 추가
                //스타일의 아이디
                int currentstyle = json["charShape"]["positionShapeIdPairList"][j]["shapeId"].Value<int>();

                /////////////////////////////////////
                /*if (j == 0)
                    name = xm.AddContentP(subcontent);
                else if (currentstyle == pstyle)
                    xm.AddContentP(i, subcontent); //pstyle과 같으면 텍스트만 추가
                else
                    name = xm.AddContentSpan(pname, subcontent); //pstyle과 다르면 span 생성 후 텍스트 추가*/
                ///////////////////////////////////////////////추가 수정 필요
                if (spancount == 1)
                {
                    pname = xm.AddContentP(pcontent);
                    name = pname;
                }
                else if (j == 0)
                {
                    pname = xm.AddContentP("");
                    name = xm.AddContentSpan(pname, subcontent);
                }
                else
                    name = xm.AddContentSpan(pname, subcontent);
                //


                //스타일 속성 추가
                //문단 속성 추가
                double baseSize = docJson["charShapeList"][currentstyle]["baseSize"].Value<double>(); // pt * 100 값

                if (j == 0)
                {
                    //줄 간격
                    int lineSpace = docJson["paraShapeList"][sID]["lineSpace"].Value<int>();
                    int property = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();
                    xm.Paragraph.SetLineSpace(pname, (XmlDocument)xm.docs["content.xml"], lineSpace, property, (int)baseSize);

                    //문단 테두리 간격
                    double topborderSpace = docJson["paraShapeList"][sID]["topBorderSpace"].Value<double>();
                    double bottomBorderSpace = docJson["paraShapeList"][sID]["bottomBorderSpace"].Value<double>();
                    double leftBorderSpace = docJson["paraShapeList"][sID]["leftBorderSpace"].Value<double>();
                    double rightBorderSpace = docJson["paraShapeList"][sID]["rightBorderSpace"].Value<double>();

                    topborderSpace *= 0.01;
                    bottomBorderSpace *= 0.01;
                    leftBorderSpace *= 0.01;
                    rightBorderSpace *= 0.01;

                    //xm.Paragraph.SetBorderSpace(pname, (XmlDocument)xm.docs["content.xml"], topborderSpace, bottomBorderSpace, leftBorderSpace, rightBorderSpace);

                    //줄 나눔
                    //한글이 ByWord일 경우 적용 (한글, 영어 모두)



                    int temp = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();

                    Byte byWord = (Byte)(temp >> 7);
                    if (byWord.Equals(1))
                    {
                        xm.Paragraph.SetByWord(pname, (XmlDocument)xm.docs["content.xml"]);
                    }

                    //외톨이줄 보호 여부
                    int isProtectLoner = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();
                    int protemp = bitcal(isProtectLoner, 15, 1);
                    if (protemp == 1)
                        xm.Paragraph.SetisProtectLoner(pname, (XmlDocument)xm.docs["content.xml"]);

                    //다음 문단과 함께 여부
                    int isTogetherNextPara = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();
                    int togtemp = bitcal(isTogetherNextPara, 16, 1);
                    if (togtemp == 1)
                        xm.Paragraph.SetisTogetherNextPara(pname, (XmlDocument)xm.docs["content.xml"]);

                    //문단 보호 여부
                    int isProtectPara = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();
                    int paratemp = bitcal(isProtectPara, 17, 1);
                    if (paratemp == 1)
                        xm.Paragraph.SetisProtectPara(pname, (XmlDocument)xm.docs["content.xml"]);

                    //한글과 영어 간격을 자동 조절 여부
                    int isAutoAdjustGapHangulEnglish = docJson["paraShapeList"][sID]["property2"]["value"].Value<int>();
                    int autotemp = bitcal(isAutoAdjustGapHangulEnglish, 4, 0x1);
                    if (autotemp == 1)
                        xm.Paragraph.SetisAutoAdjustGapHangulEnglish(pname, (XmlDocument)xm.docs["content.xml"]);


                    //정렬
                    int aligntemp = docJson["paraShapeList"][sID]["property1"]["value"].Value<int>();
                    int paraAlign = bitcal(aligntemp, 2, 0x7);
                    if (paraAlign.Equals(2))
                        xm.SetPAlign(pname, "end");
                    else if (paraAlign.Equals(0))
                        xm.SetPAlign(pname, "justify");
                    else if (paraAlign.Equals(3))
                        xm.SetPAlign(pname, "center");
                    else if (paraAlign.Equals(5))
                        xm.SetPAlign(pname, "Divide");
                    else if (paraAlign.Equals(4))
                        xm.SetPAlign(pname, "Distribute");

                    //첫줄 들여쓰기
                    //margin
                    double indent = docJson["paraShapeList"][sID]["indent"].Value<double>();
                    double topspace = docJson["paraShapeList"][sID]["topParaSpace"].Value<double>();
                    double bottomspace = docJson["paraShapeList"][sID]["bottomParaSpace"].Value<double>();
                    double leftmargin = docJson["paraShapeList"][sID]["leftMargin"].Value<double>();
                    double rightmargin = docJson["paraShapeList"][sID]["rightMargin"].Value<double>();
                    if (indent != 0 || topspace != 0 || bottomspace != 0 || leftmargin != 0 || rightmargin != 0)
                    {
                        xm.SetPIndent(pname, (float)(indent / 200 * 0.0353));
                        if (indent < 0)
                            leftmargin = -indent + leftmargin;
                        xm.SetPMargin(pname, (float)(leftmargin / 200 * 0.0353), (float)(rightmargin / 200 * 0.0353), (float)(topspace / 200 * 0.0353), (float)(bottomspace / 200 * 0.0353));
                    }



                }
                int charPro = docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>();

                int bold = bitcal(charPro, 1, 0x1);
                int italic = bitcal(charPro, 0, 0x1);
                int underline = bitcal(charPro, 2, 0x3);
                //bool kerning = docJson["charShapeList"][currentstyle]["Property"]["isKerning"].Value<bool>();

                int fontcolor = docJson["charShapeList"][currentstyle]["charColor"]["value"].Value<int>();
                int strikeline = bitcal(charPro, 18, 0x7);



                //진하게
                if (bold == 1)
                    xm.SetBold(name);
                //기울임
                if (italic == 1)
                    xm.SetItalic(name);

                //커닝 odt에서 지원하지 않는 기능
                /*
                if (kerning)
                {
                    baseSize = docJson["charShapeList"][currentstyle]["baseSize"].Value<int>(); // pt * 100 값
                    double kerningSpace = docJson["charShapeList"][currentstyle]["CharSpacebyLanguage"]["Hangul"].Value<int>(); // kerning percent 값
                    double value = baseSize * (kerningSpace / 100) / 100 * 0.0353; //pt로 환산 후 cm로 환산
                    xm.SetKerning(name, baseSize.ToString() + "cm");
                }
                */

                //폰트 사이즈, 폰트 색
                xm.SetFontSize(name, docJson["charShapeList"][currentstyle]["baseSize"].Value<float>() / 100);
                if (fontcolor != 0)
                {
                    byte[] bit = BitConverter.GetBytes(fontcolor);
                    Array.Reverse(bit);
                    fontcolor = BitConverter.ToInt32(bit, 0);
                    xm.SetFontColor(name, "#" + fontcolor.ToString("X8").Substring(0, 6));
                }


                //윗줄 밑줄
                if (underline.Equals(1))
                {
                    int templineshape = docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>();
                    int lineshape = bitcal(templineshape, 4, 0xf);
                    int linecolor = docJson["charShapeList"][currentstyle]["underLineColor"]["value"].Value<int>();
                    byte[] bit = BitConverter.GetBytes(linecolor);
                    Array.Reverse(bit);
                    linecolor = BitConverter.ToInt32(bit, 0);
                    ///lineshape 값? 왜 이렇게들어가느지
                    xm.SetUnderline(name, lineshape, "#" + linecolor.ToString("X8").Substring(0, 6));
                }
                else if (underline.Equals(3))
                {
                    int templineshape = docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>();
                    int lineshape = bitcal(templineshape, 4, 0xf);
                    int linecolor = docJson["charShapeList"][currentstyle]["underLineColor"]["value"].Value<int>();
                    byte[] bit = BitConverter.GetBytes(linecolor);
                    Array.Reverse(bit);
                    linecolor = BitConverter.ToInt32(bit, 0);
                    xm.SetUnderline(name, lineshape, "#" + linecolor.ToString("X8").Substring(0, 6));
                }

                //취소선 odt에서는 종류 제한, 색상 선택 불가
                if (strikeline > 0)
                {
                    int templineshape = docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>();
                    int lineshape = bitcal(templineshape, 4, 0xf);
                    xm.SetThroughline(name, lineshape);
                }

                //외곽선 한글에 종류 여러개지만 odt에서는 한개
                if (bitcal((docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 8, 0x7) != 0)
                {
                    xm.SetOutline(name);
                }

                //그림자 odt한종류
                if (bitcal((docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 11, 0x3) != 0)
                {
                    xm.SetShadow(name);
                }

                //음각 양각
                if (bitcal((docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 13, 0x1) != 0)
                {
                    xm.SetRelief(name);
                }
                else if (bitcal((docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 14, 0x1) != 0)
                {
                    xm.SetRelief(name, "engraved");
                }

                //위첨자
                //아래첨자
                if (bitcal((docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 15, 0x1) != 0)
                {
                    xm.SetSuper(name);
                }
                else if (bitcal((docJson["charShapeList"][currentstyle]["property"]["value"].Value<int>()), 16, 0x1) != 0)
                {
                    xm.SetSub(name);
                }

                //글꼴 적용
                int fontID = docJson["charShapeList"][currentstyle]["faceNameIds"]["array"][0].Value<int>();
                string fontName = docJson["hangulFaceNameList"][fontID]["name"].Value<string>();
                xm.SetFont(name, fontName);

                //글자간격
                double charspace = docJson["charShapeList"][currentstyle]["charSpaces"]["array"][0].Value<double>();
                if (charspace != 0)
                    xm.SetLetterSpace(name, (float)(baseSize * 0.01 * charspace * 0.01 * 0.0353));

            }
        }
    }
}