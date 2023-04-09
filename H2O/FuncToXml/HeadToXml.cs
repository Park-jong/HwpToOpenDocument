using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace WindowsFormsApp1.FuncToXml
{
    public class HeadToXml
    {
        public HeadToXml()
        {

        }

        JToken jsonbinData;
        JToken jsoncontrol;
        //JToken jsonHeader;


        private void setData()
        {
        }

        public int bitcal(int i, int shift, byte b) //대부분 property에 있는 value를 bit별로 나누어서 넣을때 
        /** i = value 값  shift = 이동갯수 b = &연산할 값 */
        {
            byte temp = (byte)(i >> shift);
            return (int)(temp & b);
        }
        Image StringToImage(sbyte[] _Image)
        {

            sbyte[] bitmapData = new sbyte[_Image.Length];
            MemoryStream ms = new MemoryStream((byte[])((Array)_Image));
            Image returnImage = Image.FromStream(ms);
            return returnImage;

        }


        public void Run(XmlManager xm, JToken binJson, JToken bodyJson, JToken docJson, bool zeroCheck)
        {
            //Image있는지 체크

            bool hasImg = false;
            int hasImgControlNum;
            int controlListCount = 0;
            try
            {
                controlListCount = bodyJson["controlList"].Count();
            }
            catch (System.ArgumentNullException e)
            {
                controlListCount = 0;
            }

            //머리말이미지
            //머리말은 스타일에들어감

            bool headerhasImg = false;
            bool hasHeader = false;
            int headerImgcontrolListCount = 0;
            int hasHeadernum = 0;
            int headerhasImgControlNum = 0;


            for (int controlList = 0; controlList < controlListCount; controlList++)
            {
                int Id = bodyJson["controlList"][controlList]["header"]["ctrlId"].Value<int>();
                if (Id == 1751474532)
                {
                    hasHeader = true;
                    hasHeadernum = controlList;
                    for (int i = 0; i < bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"].Count(); i++)
                        for (int j = 0; j < bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"].Count(); j++)
                        {
                            try
                            {

                                object existImg = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"].Value<object>();
                                headerhasImg = true;
                                headerhasImgControlNum = controlList;

                            }
                            catch (System.ArgumentNullException e)
                            {
                                headerhasImg = false;
                            }

                            //image가 존재할때만 실행
                            if (headerhasImg)
                            {
                                // int borderColor = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<int>();
                                int ID = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["pictureInfo"]["binItemID"].Value<int>();
                                int IDindex = 0;
                                int imgWidth = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["width"].Value<int>();
                                int imgHeight = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["height"].Value<int>();
                                // String borderProperty = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<String>();
                                double X = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["xOffset"].Value<int>();
                                double Y = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["yOffset"].Value<int>();
                                int property = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["property"]["value"].Value<int>();
                                int VertiRelTo = bitcal(property, 3, 0x3);
                                int VertiRelToarray = bitcal(property, 5, 0x7);
                                int HorzRelTo = bitcal(property, 8, 0x3);
                                int HorzRelToarray = bitcal(property, 10, 0x7);
                                int geulja = bitcal(property, 0, 0x1);
                                int through = bitcal(property, 21, 0x7);
                                int zindex = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["zOrder"].Value<int>();
                                int linevertical = 0;
                                if (bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"] != null)
                                    linevertical = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"]["lineSegItemList"][0]["lineVerticalPosition"].Value<int>();

                                for (int k = 0; k < binJson["embeddedBinaryDataList"].Count(); k++)
                                {
                                    String temp2 = binJson["embeddedBinaryDataList"][k]["name"].Value<String>();
                                    String tt = temp2.Substring(3, 4);

                                    if (ID == Convert.ToInt16(tt, 16))
                                        IDindex = k;

                                }
                                String name = binJson["embeddedBinaryDataList"][IDindex]["name"].Value<String>();

                                JArray temp = binJson["embeddedBinaryDataList"][IDindex]["data"].Value<JArray>();
                                sbyte[] items = temp.ToObject<sbyte[]>();

                                //ImgNode 생성 및 Pictures폴더 child 설정
                                ImgNode node = setPicturesChild(xm, name);
                                node.img = StringToImage(items);

                                String extension = docJson["binDataList"][IDindex]["extensionForEmbedding"].Value<String>();


                                double width = Math.Round(imgWidth * 2.54 / 7200, 3);
                                double height = Math.Round((imgHeight) * 2.54 / 7200, 3);
                                string currentPath = "Pictures/" + name;
                                xm.imgstyle(VertiRelTo, VertiRelToarray, HorzRelTo, HorzRelToarray, through, true, false);
                                xm.makeimg(width, height, extension, currentPath, X, Y, geulja, zindex, linevertical, VertiRelTo, ID, true, false);

                            }
                        }
                }
                if (Id == 1718579060)
                {
                    hasHeader = true;
                    hasHeadernum = controlList;
                    for (int i = 0; i < bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"].Count(); i++)
                        for (int j = 0; j < bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"].Count(); j++)
                        {
                            try
                            {

                                object existImg = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"].Value<object>();
                                headerhasImg = true;
                                headerhasImgControlNum = controlList;

                            }
                            catch (System.ArgumentNullException e)
                            {
                                headerhasImg = false;
                            }

                            //image가 존재할때만 실행
                            if (headerhasImg)
                            {
                                // int borderColor = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<int>();
                                int ID = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["pictureInfo"]["binItemID"].Value<int>();
                                int IDindex = 0;
                                int imgWidth = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["width"].Value<int>();
                                int imgHeight = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["height"].Value<int>();
                                // String borderProperty = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["shapeComponentPicture"]["borderColor"].Value<String>();
                                double X = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["xOffset"].Value<int>();
                                double Y = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["yOffset"].Value<int>();
                                int property = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["property"]["value"].Value<int>();
                                int VertiRelTo = bitcal(property, 3, 0x3);
                                int VertiRelToarray = bitcal(property, 5, 0x7);
                                int HorzRelTo = bitcal(property, 8, 0x3);
                                int HorzRelToarray = bitcal(property, 10, 0x7);
                                int geulja = bitcal(property, 0, 0x1);
                                int through = bitcal(property, 21, 0x7);
                                int zindex = bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["controlList"][j]["header"]["zOrder"].Value<int>();

                                int linevertical = 0;
                                if (bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"] != null)
                                    for (int m = 0; m < bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"]["lineSegItemList"].Count(); m++)
                                        linevertical += bodyJson["controlList"][hasHeadernum]["paragraphList"]["paragraphList"][i]["lineSeg"]["lineSegItemList"][m]["lineVerticalPosition"].Value<int>();
                                for (int k = 0; k < binJson["embeddedBinaryDataList"].Count(); k++)
                                {
                                    String temp2 = binJson["embeddedBinaryDataList"][k]["name"].Value<String>();
                                    String tt = temp2.Substring(3, 4);
                                    UInt16 ta = Convert.ToUInt16(tt, 16);
                                    if (ID == ta)
                                        IDindex = k;

                                }
                                String name = binJson["embeddedBinaryDataList"][IDindex]["name"].Value<String>();

                                JArray temp = binJson["embeddedBinaryDataList"][IDindex]["data"].Value<JArray>();
                                sbyte[] items = temp.ToObject<sbyte[]>();

                                //ImgNode 생성 및 Pictures폴더 child 설정
                                ImgNode node = setPicturesChild(xm, name);
                                node.img = StringToImage(items);

                                String extension = docJson["binDataList"][IDindex]["extensionForEmbedding"].Value<String>();


                                double width = Math.Round(imgWidth * 2.54 / 7200, 3);
                                double height = Math.Round((imgHeight) * 2.54 / 7200, 3);
                                string currentPath = "Pictures/" + name;
                                xm.imgstyle(VertiRelTo, VertiRelToarray, HorzRelTo, HorzRelToarray, through, false, true);
                                xm.makeimg(width, height, extension, currentPath, X, Y, geulja, zindex, (int)linevertical, VertiRelTo, ID, false, true);

                            }
                        }
                }
            }
        }

        public ImgNode setPicturesChild(XmlManager xm, string name)
        {
            FolderNode pictures = (FolderNode)xm.root.child["Pictures"];
            ImgNode node = new ImgNode(name, pictures);
            node.path = Application.StartupPath + @"\New File\Pictures\" + name;

            return node;
        }

    }
}