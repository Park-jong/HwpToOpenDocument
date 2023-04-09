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
    public class ImgToXml
    {
        public ImgToXml()
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
            for (int controlList = 0; controlList < controlListCount; controlList++)
            {
                try
                {
                    object existImg = bodyJson["controlList"][controlList]["shapeComponentPicture"].Value<object>();
                    hasImg = true;
                    hasImgControlNum = controlList;
                }
                catch (System.ArgumentNullException e)
                {
                    hasImg = false;
                }
                //image가 존재할때만 실행
                if (hasImg)
                {
                    // int borderColor = bodyjson["controlList"][controlList]["shapeComponentPicture"]["borderColor"].Value<int>();
                    int ID = bodyJson["controlList"][controlList]["shapeComponentPicture"]["pictureInfo"]["binItemID"].Value<int>();
                    int IDindex = 0;
                    int imgWidth = bodyJson["controlList"][controlList]["header"]["width"].Value<int>();
                    int imgHeight = bodyJson["controlList"][controlList]["header"]["height"].Value<int>();
                    // String borderProperty = bodyJson["controlList"][controlList]["shapeComponentPicture"]["borderColor"].Value<String>();
                    double X = bodyJson["controlList"][controlList]["header"]["xOffset"].Value<int>();
                    double Y = bodyJson["controlList"][controlList]["header"]["yOffset"].Value<int>();
                    int property = bodyJson["controlList"][controlList]["header"]["property"]["value"].Value<int>();
                    int VertiRelTo = bitcal(property, 3, 0x3);
                    int VertiRelToarray = bitcal(property, 5, 0x7);
                    int HorzRelTo = bitcal(property, 8, 0x3);
                    int HorzRelToarray = bitcal(property, 10, 0x7);
                    int geulja = bitcal(property, 0, 0x1);
                    int through = bitcal(property, 21, 0x7);
                    int zindex = bodyJson["controlList"][controlList]["header"]["zOrder"].Value<int>();
                    int linevertical = 0;
                    if (bodyJson["lineSeg"] != null)
                        linevertical = bodyJson["lineSeg"]["lineSegItemList"][0]["lineVerticalPosition"].Value<int>();
                    for (int i = 0; i < binJson["embeddedBinaryDataList"].Count(); i++)
                    {
                        String temp2 = binJson["embeddedBinaryDataList"][i]["name"].Value<String>();
                        String tt = temp2.Substring(3, 4);
                        if (ID == Convert.ToInt16(tt, 16))
                            IDindex = i;

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
                    xm.imgstyle(VertiRelTo, VertiRelToarray, HorzRelTo, HorzRelToarray, through, false, false);
                    xm.makeimg(width, height, extension, currentPath, X, Y, geulja, zindex, linevertical, VertiRelTo, ID, false, false);


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