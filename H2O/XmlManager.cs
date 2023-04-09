using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Collections;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;

namespace WindowsFormsApp1
{
    public class XmlManager
    {
        public FileNode root;
        private const string header_style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        private const string header_fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        private const string header_office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private const string header_text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private const string header_svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        private const string header_table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        private const string header_draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        private const string header_xlink = "http://www.w3.org/1999/xlink";
        private const string header_manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        private const string header_loext = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
        private static int numP = 0;
        private static int numSpan = 0;
        private static int numMP = 0;
        private static int numMT = 0;
        private static int numTable = 0;

        private static int numImage = 0;
        private static int imageStylenum = 0;
        private static int numheader = 0;
        private static int numfooter = 0;
        public bool ContentXml { get; set; }//본문인지 아닌지 판별

        public Hashtable docs = new Hashtable();

        public ParagraphManager Paragraph;

        public XmlManager()
        {
            Paragraph = new ParagraphManager();
            numP = 0;
            numSpan = 0;
            numMP = 0;
            numMT = 0;
            numTable = 0;

        }

        public void CreateODT()
        {
            string path = Application.StartupPath;

            root = new FolderNode("New File");
            root.path = path + @"\" + root.name;


            //Create File

            string[] list1 = { "Configurations2", "META-INF", "Thumbnails", "Pictures" };

            foreach (string s in list1)
            {
                FolderNode node = new FolderNode(s, root);
            }

            FolderNode Configurations2 = (FolderNode)root.child["Configurations2"];

            string[] list2 = { "accelerator", "floater", "images", "menubar", "popupmenu", "progressbar", "statusbar", "toolbar", "toolpanel" };

            foreach (string s in list2)
            {
                FolderNode node = new FolderNode(s, Configurations2);
            }

            FolderNode images = (FolderNode)Configurations2.child["images"];

            FolderNode Bitmaps = new FolderNode("Bitmaps", images);


            //CreateXML

            string[] file1 = { "content.xml", "manifest.xml", "meta.xml", "mimetype", "settings.xml", "styles.xml" };

            foreach (string s in file1)
            {
                XmlNode node = new XmlNode(s, root);

                node.LoadXml(path + @"\data" + @"\" + node.name);

                docs.Add(node.name, node.doc);
            }

            FolderNode META_INF = (FolderNode)root.child["META-INF"];

            XmlNode manifest = new XmlNode("manifest.xml", META_INF);
            manifest.LoadXml(path + @"\data" + @"\META-INF" + @"\" + manifest.name);


            FolderNode accelerator = (FolderNode)Configurations2.child["accelerator"];

            XmlNode current = new XmlNode("current.xml", accelerator);
            current.LoadXml(path + @"\data" + @"\Configurations2" + @"\accelerator" + @"\" + current.name);

        }

        public void SaveODT(FileNode node)
        {
            if (node.GetType() == typeof(FolderNode))
            {
                DirectoryInfo di = new DirectoryInfo(node.path);
                if (!di.Exists) { di.Create(); }
            }
            else if (node.GetType() == typeof(XmlNode))
            {
                if (((XmlNode)node).name == "mimetype")
                {
                    StreamWriter sw = File.CreateText(((XmlNode)node).path);
                    sw.Write("application/vnd.oasis.opendocument.text");
                    sw.Close();
                }
                else
                {
                    string path;

                    if (((XmlNode)node).path.Contains("New File\\manifest.xml"))
                    {
                        path = ((XmlNode)node).path.Replace(".xml", ".rdf");
                    }
                    else
                    {
                        path = ((XmlNode)node).path;
                    }

                    ((XmlNode)node).doc.Save(path);
                }
            }
            else if (node.GetType() == typeof(ImgNode))
            {
                SaveImage(((ImgNode)node));
            }

            if (node.child == null)
                return;

            foreach (FileNode child in node.child.Values)
            {
                SaveODT(child);
            }

        }

        void SaveImage(ImgNode node)
        {
            Image image = node.img;
            string filePath = node.path;

            string fileExtension = Path.GetExtension(filePath);
            switch (fileExtension.ToLower())
            {

                case ".bmp": image.Save(filePath, ImageFormat.Bmp); break;
                case ".exif": image.Save(filePath, ImageFormat.Exif); break;
                case ".gif": image.Save(filePath, ImageFormat.Gif); break;
                case ".jpg": case ".jpeg": image.Save(filePath, ImageFormat.Jpeg); break;
                case ".png": image.Save(filePath, ImageFormat.Png); break;
                case ".tif": case ".tiff": image.Save(filePath, ImageFormat.Tiff); break;
            }
        }



        public void SetFontSize(string name, float size)
        {
            {
                XmlNode content;
                if (ContentXml == true)
                    content = (XmlNode)root.child["content.xml"];
                else
                    content = (XmlNode)root.child["styles.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("style", header_style);

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {

                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("font-size", header_fo, size + "pt");
                        e1.SetAttribute("font-size-asian", header_style, size + "pt");
                        e1.SetAttribute("font-size-complex", header_style, size + "pt");

                        e.AppendChild(e1);
                        break;
                    }
                }
            }
        }

        public void SetFontColor(string name, string color)
        {
            {
                XmlNode content;
                if (ContentXml == true)
                    content = (XmlNode)root.child["content.xml"];
                else
                    content = (XmlNode)root.child["styles.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("style", header_style);

                ((XmlElement)doc.GetElementsByTagName("document-content", header_office).Item(0)).SetAttribute("xmlns:loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0");

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("color", header_fo, color);
                        e1.SetAttribute("opacity", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0", "100%");

                        e.AppendChild(e1);
                        break;
                    }
                }
            }
        }

        public void SetLetterSpace(string name, float space)
        {
            {
                XmlNode content;
                if (ContentXml == true)
                    content = (XmlNode)root.child["content.xml"];
                else
                    content = (XmlNode)root.child["styles.xml"];

                XmlDocument doc = content.doc;

                XmlNodeList list = doc.GetElementsByTagName("style", header_style);

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("letter-spacing", header_fo, space + "cm");

                        e.AppendChild(e1);
                        break;
                    }
                }
            }
        }

        //글꼴 적용
        public void SetFont(string name, string font)
        {
            {
                XmlDocument doc;
                if (ContentXml == true)
                    doc = (XmlDocument)docs["content.xml"];
                else
                    doc = (XmlDocument)docs["styles.xml"];

                //office:font-face-decls에 폰트 중복 체크
                XmlNodeList list = doc.GetElementsByTagName("font-face", header_style);
                for (int i = 0; i < list.Count; i++)
                {
                    string fontname = ((XmlElement)list.Item(i)).GetAttribute("name", header_style);
                    if (fontname.Equals(font))
                        break;
                    if (i == list.Count - 1)
                    {
                        XmlNodeList fontlist = doc.GetElementsByTagName("font-face-decls", header_office);
                        XmlElement addfont = (XmlElement)fontlist.Item(0);
                        XmlElement e1 = doc.CreateElement("style:font-face", header_style);

                        e1.SetAttribute("name", header_style, font);
                        e1.SetAttribute("font-family", header_svg, font);
                        e1.SetAttribute("font-family-generic", header_style, "system");
                        e1.SetAttribute("font-pitch", header_style, "variable");
                        //panose 제외

                        addfont.AppendChild(e1);

                    }

                }

                list = doc.GetElementsByTagName("style", header_style);

                foreach (XmlElement e in list)
                {
                    string check_name = e.GetAttribute("name", header_style);

                    if (check_name.Equals(name))
                    {
                        XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                        if (e1 == null)
                        {
                            e1 = doc.CreateElement("style:text-properties", header_style);
                        }
                        e1.SetAttribute("font-name", header_style, font);
                        e1.SetAttribute("font-name-asian", header_style, font);
                        e1.SetAttribute("font-name-complex", header_style, font);
                        //font-name-asian 제외
                        //font-name-complex 제외

                        e.AppendChild(e1);
                        break;
                    }
                }

                XmlNode styles = (XmlNode)root.child["styles.xml"];
                doc = styles.doc;

                list = doc.GetElementsByTagName("font-face", header_style);
                for (int i = 0; i < list.Count; i++)
                {
                    string fontname = ((XmlElement)list.Item(i)).GetAttribute("name", header_style);
                    if (fontname.Equals(font))
                        break;
                    if (i == list.Count - 1)
                    {
                        XmlNodeList fontlist = doc.GetElementsByTagName("font-face-decls", header_office);
                        XmlElement addfont = (XmlElement)fontlist.Item(0);
                        XmlElement e1 = doc.CreateElement("style:font-face", header_style);

                        e1.SetAttribute("name", header_style, font);
                        e1.SetAttribute("font-family", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0", font);
                        e1.SetAttribute("font-family-generic", header_style, "system");
                        e1.SetAttribute("font-pitch", header_style, "variable");

                        addfont.AppendChild(e1);

                    }
                }
            }
        }

        public void SetBold(string name)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("font-weight", header_fo, "bold");
                    e1.SetAttribute("font-weight-asian", header_style, "bold");
                    e1.SetAttribute("font-weight-complex", header_style, "bold");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetItalic(string name)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("font-style", header_fo, "italic");
                    e1.SetAttribute("font-style-asian", header_style, "italic");
                    e1.SetAttribute("font-style-complex", header_style, "italic");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetKerning(string name, string kerningValue)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);
            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);
                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }

                    e1.SetAttribute("letter-spacing", header_fo, kerningValue);
                    e1.SetAttribute("letter-kerning", header_style, "true");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetUnderline(string name, int stylenum = 0, string color = "font-color")
        {
            string style = "solid";
            string width = "auto";
            switch (stylenum)
            {
                case 0:
                    style = "solid";
                    break;
                case 1:
                    style = "dash";
                    break;
                case 2:
                    style = "dotted";
                    break;
                case 3:
                    style = "dot-dash";
                    break;
                case 4:
                    style = "dot-dot-dash";
                    break;
                case 5:
                    style = "long-dash";
                    break;
                case 6:
                    style = "dash";
                    width = "bold";
                    break;
                case 7:
                    style = "solid";
                    break;
                case 11:
                    style = "wave";
                    break;
                case 12:
                    style = "wave";
                    break;
                default:
                    break;
            }
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-underline-style", header_style, style);
                    e1.SetAttribute("text-underline-width", header_style, width);
                    e1.SetAttribute("text-underline-color", header_style, color);
                    if (stylenum == 7 || stylenum == 12)
                    {
                        e1.SetAttribute("text-underline-type", header_style, "double");
                    }

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetOverline(string name, int stylenum = 0, string color = "font-color")
        {
            string style = "solid";
            string width = "auto";
            switch (stylenum)
            {
                case 0:
                    style = "solid";
                    break;
                case 1:
                    style = "dash";
                    break;
                case 2:
                    style = "dotted";
                    break;
                case 3:
                    style = "dot-dash";
                    break;
                case 4:
                    style = "dot-dot-dash";
                    break;
                case 5:
                    style = "long-dash";
                    break;
                case 6:
                    style = "dash";
                    width = "bold";
                    break;
                case 7:
                    style = "solid";
                    break;
                case 11:
                    style = "wave";
                    break;
                case 12:
                    style = "wave";
                    break;
                default:
                    break;
            }

            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-overline-style", header_style, style);
                    e1.SetAttribute("text-overline-width", header_style, width);
                    e1.SetAttribute("text-overline-color", header_style, color);
                    if (stylenum == 7 || stylenum == 12)
                    {
                        e1.SetAttribute("text-overline-type", header_style, "double");
                    }

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetThroughline(string name, int stylenum = 0)
        {
            string style = "solid";
            string type = "single";
            switch (stylenum)
            {
                case 0:
                    style = "solid";
                    break;
                case 7:
                    style = "solid";
                    type = "double";
                    break;
                default:
                    break;
            }

            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-line-through-style", header_style, style);
                    e1.SetAttribute("text-line-through-type", header_style, type);

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetOutline(string name)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-outline", header_style, "true");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetShadow(string name)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-shadow", header_fo, "1pt 1pt");

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetRelief(string name, string type = "embossed")
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("font-relief", header_style, type);//engraved, embossed

                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetSuper(string name)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-position", header_style, "super 58%");
                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetSub(string name)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("text-properties", header_style).Item(0);

                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:text-properties", header_style);
                    }
                    e1.SetAttribute("text-position", header_style, "sub 58%");
                    e.AppendChild(e1);
                    break;
                }
            }
        }

        public void SetPMargin(string name, float left, float right, float top, float bottom)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:paragraph-properties", header_style);
                    }
                    e1.SetAttribute("margin-left", header_fo, left + "cm");
                    e1.SetAttribute("margin-right", header_fo, right + "cm");
                    e1.SetAttribute("margin-top", header_fo, top + "cm");
                    e1.SetAttribute("margin-bottom", header_fo, bottom + "cm");

                    e.PrependChild(e1);
                    break;
                }
            }
        }

        public void SetPIndent(string name, float ident)
        {
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:paragraph-properties", header_style);
                    }
                    e1.SetAttribute("text-indent", header_fo, ident + "cm");

                    e.PrependChild(e1);
                    break;
                }
            }
        }

        public void SetPAlign(string name, string align)
        {
            string set_align = align;
            if (align.Equals("Divide") || align.Equals("Distribute"))
            {
                set_align = "justify";
            }
            XmlNode content;
            if (ContentXml == true)
                content = (XmlNode)root.child["content.xml"];
            else
                content = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);

            foreach (XmlElement e in list)
            {
                string check_name = e.GetAttribute("name", header_style);

                if (check_name.Equals(name))
                {
                    XmlElement e1 = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);
                    if (e1 == null)
                    {
                        e1 = doc.CreateElement("style:paragraph-properties", header_style);
                    }
                    e1.SetAttribute("text-align", header_fo, set_align);
                    e1.SetAttribute("justify-single-word", header_style, "false");
                    if (align.Equals("Divide"))
                    {
                        e1.SetAttribute("text-align-last", header_fo, set_align);
                    }
                    else if (align.Equals("Distribute"))
                    {
                        e1.SetAttribute("text-align-last", header_fo, set_align);
                        e1.SetAttribute("justify-single-word", header_style, "true");
                    }

                    e.PrependChild(e1);
                    break;
                }
            }
        }

        // p 생성
        public string AddContentP(string text)
        {
            string pname = "P" + (numP + 1).ToString();

            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            // 문단 스타일 생성
            XmlElement e1 = doc.CreateElement("style:style", header_style);

            e1.SetAttribute("name", header_style, pname);
            e1.SetAttribute("family", header_style, "paragraph");
            e1.SetAttribute("parent-style-name", header_style, "Standard");

            e.AppendChild(e1);

            // 문단 내용 생성
            list = doc.GetElementsByTagName("text", header_office);

            e = (XmlElement)list.Item(0);

            XmlElement text_element = doc.CreateElement("text:p", header_text);

            text_element.SetAttribute("style-name", header_text, pname);
            foreach (char eachtext in text)
            {
                if (eachtext.Equals(' '))
                {
                    XmlElement s = doc.CreateElement("text:s", header_text);
                    text_element.AppendChild(s);
                }
                else
                {
                    XmlText t = doc.CreateTextNode(eachtext.ToString());
                    text_element.AppendChild(t);
                }
            }
            e.AppendChild(text_element);

            numP++;

            return pname;
        }

        // p 생성 : null
        public string AddContentP()
        {
            string pname = "P" + (numP + 1).ToString();

            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;


            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            // 문단 스타일 생성
            XmlElement e1 = doc.CreateElement("style:style", header_style);
            e1.SetAttribute("name", header_style, pname);
            e1.SetAttribute("family", header_style, "paragraph");
            e1.SetAttribute("parent-style-name", header_style, "Standard");
            e.AppendChild(e1);

            list = doc.GetElementsByTagName("text", header_office);
            e = (XmlElement)list.Item(0);

            XmlElement text_element = doc.CreateElement("text:p", header_text);

            text_element.SetAttribute("style-name", header_text, pname);
            e.AppendChild(text_element);

            numP++;

            return pname;
        }

        // text 추가
        public void AddContentP(string pname, string text)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("p", header_text);

            XmlText xmltext = doc.CreateTextNode(text);
            foreach (XmlElement e in list)
            {
                string namecheck = e.GetAttribute("name", header_style);
                if (namecheck.Equals(pname))
                {
                    e.AppendChild(xmltext);
                    break;
                }
            }

        }

        // span 추가
        public string AddContentSpan(string pname, string text)
        {
            string spanname = "T" + (numSpan + 1).ToString();
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement e1 = doc.CreateElement("style:style", header_style);

            e1.SetAttribute("name", header_style, spanname);
            e1.SetAttribute("family", header_style, "text");

            e.AppendChild(e1);

            list = doc.GetElementsByTagName("p", header_text);

            //현재는 모든 p가 이름이 다르기 때문에 가능 수정?
            foreach (XmlElement element in list)
            {
                string name_check = element.GetAttribute("style-name", header_text);
                if (name_check.Equals(pname))
                {
                    XmlElement text_element = doc.CreateElement("text:span", header_text);

                    text_element.SetAttribute("style-name", header_text, spanname);
                    foreach (char eachtext in text)
                    {
                        if (eachtext.Equals(' '))
                        {
                            XmlElement s = doc.CreateElement("text:s", header_text);
                            text_element.AppendChild(s);
                        }
                        else
                        {
                            XmlText t = doc.CreateTextNode(eachtext.ToString());
                            text_element.AppendChild(t);
                        }
                    }
                    element.AppendChild(text_element);
                    numSpan++;
                    break;
                }
            }
            return spanname;
        }

        public void SetPageLayout(double left, double right, double top, double bottom)
        {
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlElement pagelayout = (XmlElement)doc.GetElementsByTagName("page-layout-properties", header_style).Item(0);
            pagelayout.SetAttribute("margin-left", header_fo, left + "cm");
            pagelayout.SetAttribute("margin-right", header_fo, right + "cm");
            pagelayout.SetAttribute("margin-top", header_fo, top + "cm");
            pagelayout.SetAttribute("margin-bottom", header_fo, bottom + "cm");
        }

        public void SetHeader(double headermargin, double left, double right, double bottom)
        {
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlElement headerstyle = (XmlElement)doc.GetElementsByTagName("header-style", header_style).Item(0);
            XmlElement header_footer_properties = doc.CreateElement("style:header-footer-properties", header_style);
            header_footer_properties.SetAttribute("height", header_svg, headermargin + "cm");
            header_footer_properties.SetAttribute("margin-left", header_fo, left + "cm");
            header_footer_properties.SetAttribute("margin-right", header_fo, right + "cm");
            header_footer_properties.SetAttribute("margin-bottom", header_fo, bottom + "cm");
            header_footer_properties.SetAttribute("dynamic-spacing", header_style, "false");
            header_footer_properties.SetAttribute("background-color", header_fo, "transparent");
            header_footer_properties.SetAttribute("fill", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", "none");
            headerstyle.AppendChild(header_footer_properties);
        }

        public void SetFooter(double footermargin, double left, double right, double top)
        {
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlElement footerstyle = (XmlElement)doc.GetElementsByTagName("footer-style", header_style).Item(0);
            XmlElement header_footer_properties = doc.CreateElement("style:header-footer-properties", header_style);
            header_footer_properties.SetAttribute("height", header_svg, footermargin + "cm");
            header_footer_properties.SetAttribute("margin-left", header_fo, left + "cm");
            header_footer_properties.SetAttribute("margin-right", header_fo, right + "cm");
            header_footer_properties.SetAttribute("margin-bottom", header_fo, top + "cm");
            header_footer_properties.SetAttribute("dynamic-spacing", header_style, "false");
            header_footer_properties.SetAttribute("background-color", header_fo, "transparent");
            header_footer_properties.SetAttribute("fill", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", "none");
            footerstyle.AppendChild(header_footer_properties);
        }

        public string AddHeader()
        {
            string name = "MP" + (numMP + 1).ToString();
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            XmlElement element = (XmlElement)list.Item(0);

            XmlElement styleElement = doc.CreateElement("style:style", header_style);
            styleElement.SetAttribute("name", header_style, name);
            styleElement.SetAttribute("family", header_style, "paragraph");
            styleElement.SetAttribute("parent-style-name", header_style, "Header");
            element.PrependChild(styleElement);

            XmlElement text_properties = doc.CreateElement("style:text-properties", header_style);
            styleElement.AppendChild(text_properties);

            XmlElement master_page = (XmlElement)doc.GetElementsByTagName("master-page", header_style).Item(0);
            XmlElement header = (XmlElement)master_page.GetElementsByTagName("header", header_style).Item(0);
            if (header == null)
            {
                header = doc.CreateElement("style:header", header_style);
                master_page.AppendChild(header);
            }
            XmlElement p = doc.CreateElement("text:p", header_text);
            p.SetAttribute("style-name", header_text, name);
            header.AppendChild(p);
            numMP++;

            return name;
        }

        public string AddFooter()
        {
            string name = "MP" + (numMP + 1).ToString();
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            XmlElement element = (XmlElement)list.Item(0);

            XmlElement styleElement = doc.CreateElement("style:style", header_style);
            styleElement.SetAttribute("name", header_style, name);
            styleElement.SetAttribute("family", header_style, "paragraph");
            styleElement.SetAttribute("parent-style-name", header_style, "Footer");
            element.PrependChild(styleElement);

            XmlElement text_properties = doc.CreateElement("style:text-properties", header_style);
            styleElement.AppendChild(text_properties);

            XmlElement master_page = (XmlElement)doc.GetElementsByTagName("master-page", header_style).Item(0);
            XmlElement footer = (XmlElement)master_page.GetElementsByTagName("footer", header_style).Item(0);
            if (footer == null)
            {
                footer = doc.CreateElement("style:footer", header_style);
                master_page.AppendChild(footer);
            }
            XmlElement p = doc.CreateElement("text:p", header_text);
            p.SetAttribute("style-name", header_text, name);
            footer.AppendChild(p);
            numMP++;

            return name;
        }

        public string AddHeaderFooterSpan(string p, string text)
        {
            string name = "MT" + (numMT + 1).ToString();
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            XmlElement element = (XmlElement)list.Item(0);

            XmlElement styleElement = doc.CreateElement("style:style", header_style);
            styleElement.SetAttribute("name", header_style, name);
            styleElement.SetAttribute("family", header_style, "text");
            element.AppendChild(styleElement);

            XmlElement text_properties = doc.CreateElement("style:text-properties", header_style);
            styleElement.AppendChild(text_properties);

            XmlElement master_page = (XmlElement)doc.GetElementsByTagName("master-page", header_style).Item(0);
            list = master_page.GetElementsByTagName("p", header_text);
            XmlElement e = (XmlElement)list.Item(list.Count - 1);
            XmlElement span = doc.CreateElement("text:span", header_text);
            span.SetAttribute("style-name", header_text, name);
            foreach (char eachtext in text)
            {
                if (eachtext.Equals(' '))
                {
                    XmlElement s = doc.CreateElement("text:s", header_text);
                    span.AppendChild(s);
                }
                else
                {
                    XmlText t = doc.CreateTextNode(eachtext.ToString());
                    span.AppendChild(t);
                }
            }
            e.AppendChild(span);
            numMT++;

            return name;
        }

        public void SetStandardMargin(float left, float right, float top, float bottom)
        {
            XmlNode styles = (XmlNode)root.child["styles.xml"];

            XmlDocument doc = styles.doc;

            XmlNodeList list = doc.GetElementsByTagName("style", header_style);
            foreach (XmlElement e in list)
            {
                string stylename = e.GetAttribute("name", header_style);
                if (stylename.Equals("Standard"))
                {
                    XmlElement standardmargin = (XmlElement)e.GetElementsByTagName("paragraph-properties", header_style).Item(0);

                    standardmargin.SetAttribute("margin-left", header_fo, left + "cm");
                    standardmargin.SetAttribute("margin-right", header_fo, right + "cm");
                    standardmargin.SetAttribute("margin-top", header_fo, top + "cm");
                    standardmargin.SetAttribute("margin-bottom", header_fo, bottom + "cm");
                }
            }
        }

        public void Update_Text(string text)
        {
            //update content
            XmlNodeList con_list = content.GetElementsByTagName("p", content_text);

            foreach (XmlElement e in con_list)
            {
                e.InnerText = text;
            }

            //update meta
            int paragraph_count = text.Split('\n').Length;
            int word_count = text.Split(' ').Length;
            int character_count = text.Length;

            XmlNodeList meta_list = meta.GetElementsByTagName("document-statistic", meta_meta);

            foreach (XmlElement e in meta_list)
            {
                XmlAttribute xa = e.GetAttributeNode("paragraph-count", meta_meta);
                xa.Value = paragraph_count.ToString();

                xa = e.GetAttributeNode("word-count", meta_meta);
                xa.Value = word_count.ToString();

                xa = e.GetAttributeNode("character-count", meta_meta);
                xa.Value = character_count.ToString();
            }

        }

        public string MakeTable(int row_n, int col_n)
        {
            String name = "표" + (numTable + 1).ToString();
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("text", header_office);

            XmlElement e = (XmlElement)list.Item(0);

            XmlElement table = doc.CreateElement("table:table", header_table);

            table.SetAttribute("name", header_table, name);
            table.SetAttribute("style-name", header_table, name);
            e.AppendChild(table);

            for (int i = 0; i < col_n; i++)
            {
                XmlElement column = doc.CreateElement("table:table-column", header_table);
                table.AppendChild(column);
            }
            for (int i = 0; i < row_n; i++)
            {
                XmlElement row = doc.CreateElement("table:table-row", header_table);
                table.AppendChild(row);
                for (int j = 0; j < col_n; j++)
                {
                    XmlElement cell = doc.CreateElement("table:table-cell", header_table);
                    row.AppendChild(cell);
                }
            }
            numTable++;
            return name;
        }

        public void setTable(string name, double width, double outterTopMargin, double outterLeftMargin, double outterRightMargin, double outterBottomMargin)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;


            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement table = doc.CreateElement("style:style", header_style);
            table.SetAttribute("name", header_style, name);
            table.SetAttribute("family", header_style, "table");

            XmlElement tablestyle = doc.CreateElement("style:table-properties", header_style);
            tablestyle.SetAttribute("width", header_style, width + "cm");
            tablestyle.SetAttribute("align", header_table, "center");

            tablestyle.SetAttribute("margin-top", header_fo, outterTopMargin + "cm");
            tablestyle.SetAttribute("margin-left", header_fo, outterLeftMargin + "cm");
            tablestyle.SetAttribute("margin-right", header_fo, outterRightMargin + "cm");
            tablestyle.SetAttribute("margin-bottom", header_fo, outterBottomMargin + "cm");

            table.AppendChild(tablestyle);

            e.AppendChild(table);
        }

        public void setCol(string name, int col_num, double width)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;


            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement col = doc.CreateElement("style:style", header_style);
            string col_name;
            if (col_num < 26)
                col_name = ((char)('A' + col_num)).ToString();
            else
            {
                col_name = ((char)('A' + (col_num / 26))).ToString() + ((char)('A' + (col_num % 26))).ToString();
            }
            col_name = name + "." + col_name;
            col.SetAttribute("name", header_style, col_name);
            col.SetAttribute("family", header_style, "table-column");

            XmlElement colStyle = doc.CreateElement("style:table-column-properties", header_style);
            colStyle.SetAttribute("column-width", header_style, width + "cm");

            col.AppendChild(colStyle);
            e.AppendChild(col);

            list = doc.GetElementsByTagName("table", header_table);
            e = (XmlElement)list.Item(numTable - 1);

            XmlElement column = (XmlElement)e.GetElementsByTagName("table-column", header_table).Item(col_num);
            column.SetAttribute("style-name", header_table, col_name);
        }

        public void setRow(string name, int row_num, double height)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;


            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement row = doc.CreateElement("style:style", header_style);
            string row_name = name + "." + row_num;

            row.SetAttribute("name", header_style, row_name);
            row.SetAttribute("family", header_style, "table-row");

            XmlElement rowStyle = doc.CreateElement("style:table-row-properties", header_style);
            rowStyle.SetAttribute("min-row-height", header_style, height*0.95 + "cm");
            rowStyle.SetAttribute("keep-together", header_fo, "auto");

            row.AppendChild(rowStyle);
            e.AppendChild(row);

            list = doc.GetElementsByTagName("table", header_table);
            e = (XmlElement)list.Item(numTable - 1);

            XmlElement column = (XmlElement)e.GetElementsByTagName("table-row", header_table).Item(row_num);
            column.SetAttribute("style-name", header_table, row_name);
        }

        public void SetCell(string name, double topThickness, double leftThickness, double rightThickness, double bottomThickness, int colSpan, int rowSpan, int column_index, int row_index, double height, double width, double margin_top, double margin_bottom, double margin_left, double margin_right, int Verticalalign, string background_color_to_bit, string top_line_type, string left_line_type, string right_line_type, string bottom_line_type)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;
            string cell_name = name + "." + row_index + "." + column_index;

            XmlNodeList list = doc.GetElementsByTagName("table", header_table);
            XmlElement e = (XmlElement)list.Item(numTable - 1);

            list = e.GetElementsByTagName("table-row", header_table);
            e = (XmlElement)list.Item(row_index);

            XmlElement cell;
            try
            {
                cell = (XmlElement)e.GetElementsByTagName("table-cell", header_table).Item(column_index);
            }
            catch (System.NullReferenceException exp)
            {
                return;
            }

            cell.SetAttribute("style-name", header_table, cell_name);
            cell.SetAttribute("value-type", header_office, "string");

            if (colSpan > 1)
            {
                string colSpanNum = colSpan.ToString();
                cell.SetAttribute("number-columns-spanned", header_table, colSpanNum);
            }
            if (rowSpan > 1)
            {
                string rowSpanNum = rowSpan.ToString();
                cell.SetAttribute("number-rows-spanned", header_table, rowSpanNum);

            }


            list = doc.GetElementsByTagName("automatic-styles", header_office);
            e = (XmlElement)list.Item(0);

            cell = doc.CreateElement("style:style", header_style);

            cell.SetAttribute("name", header_style, cell_name);
            cell.SetAttribute("family", header_style, "table-cell");

            XmlElement cellStyle = doc.CreateElement("style:table-cell-properties", header_style);
            //cellStyle.SetAttribute("width", header_style, width + "cm");
            //cellStyle.SetAttribute("height", header_style, height + "cm");
            cellStyle.SetAttribute("padding-top", header_fo, margin_top + "cm");
            cellStyle.SetAttribute("padding-bottom", header_fo, margin_bottom + "cm");
            cellStyle.SetAttribute("padding-left", header_fo, margin_left + "cm");
            cellStyle.SetAttribute("padding-right", header_fo, margin_right + "cm");

            switch (right_line_type)
            {
                case "Solid":
                    cellStyle.SetAttribute("border-right", header_fo, rightThickness + "pt " + "solid #000000");
                    break;
                case "Dot":
                    cellStyle.SetAttribute("border-right", header_fo, rightThickness + "pt " + "dotted #000000");
                    break;
                case "Dash":
                    cellStyle.SetAttribute("border-right", header_fo, rightThickness + "pt " + "dashed #000000");
                    break;
            }
            switch (bottom_line_type)
            {
                case "Solid":
                    cellStyle.SetAttribute("border-bottom", header_fo, bottomThickness + "pt " + "solid #000000");
                    break;
                case "Dot":
                    cellStyle.SetAttribute("border-bottom", header_fo, bottomThickness + "pt " + "dotted #000000");
                    break;
                case "Dash":
                    cellStyle.SetAttribute("border-bottom", header_fo, bottomThickness + "pt " + "dashed #000000");
                    break;
            }

            if (column_index == 0)
            {
                switch (left_line_type)
                {
                    case "Solid":
                        cellStyle.SetAttribute("border-left", header_fo, leftThickness + "pt " + "solid #000000");
                        break;
                    case "Dot":
                        cellStyle.SetAttribute("border-left", header_fo, leftThickness + "pt " + "dotted #000000");
                        break;
                    case "Dash":
                        cellStyle.SetAttribute("border-left", header_fo, leftThickness + "pt " + "dashed #000000");
                        break;
                }
            }

            if (row_index == 0)
            {
                switch (top_line_type)
                {
                    case "Solid":
                        cellStyle.SetAttribute("border-top", header_fo, topThickness + "pt " + "solid #000000");
                        break;
                    case "Dot":
                        cellStyle.SetAttribute("border-top", header_fo, topThickness + "pt " + "dotted #000000");
                        break;
                    case "Dash":
                        cellStyle.SetAttribute("border-top", header_fo, topThickness + "pt " + "dashed #000000");
                        break;
                }
            }

            if (Verticalalign == 1)
            {
                cellStyle.SetAttribute("vertical-align", header_style, "middle");
            }

            if (background_color_to_bit != "0")
            {
                string strR, strG, strB;
                if (background_color_to_bit.Length == 16)
                {
                    strR = background_color_to_bit.Substring(0, 5);
                    strG = background_color_to_bit.Substring(5, 6);
                    strB = background_color_to_bit.Substring(11, 5);
                }
                else
                {
                    strR = background_color_to_bit.Substring(0, 8);
                    strG = background_color_to_bit.Substring(8, 8);
                    strB = background_color_to_bit.Substring(16, 8);
                }

                int r = 0;
                int g = 0;
                int b = 0;

                try
                {
                    r = Convert.ToByte(strB, 2);
                    g = Convert.ToByte(strG, 2);
                    b = Convert.ToByte(strR, 2);

                }
                catch (FormatException)
                {
                    System.Console.WriteLine("Invalid String");
                }

                string theHexColor = "#" + r.ToString("X2") + g.ToString("X2") + b.ToString("X2");

                cellStyle.SetAttribute("background-color", header_fo, theHexColor);
                XmlElement backgroundImage = doc.CreateElement("style:background-image", header_style);
                cellStyle.AppendChild(backgroundImage);
            }

            cell.AppendChild(cellStyle);
            e.AppendChild(cell);
        }

        public void replaceP(int row_index, int col_index, int colSpan, int rowSpan, int childCellIndex)
        {
            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("table", header_table);
            XmlElement e = (XmlElement)list.Item(numTable - 1);

            list = e.GetElementsByTagName("table-row", header_table);
            XmlElement row = (XmlElement)list.Item(row_index);

            for (int changeIndex = 1; changeIndex < colSpan; changeIndex++)
            {
                XmlElement covered = doc.CreateElement("table:covered-table-cell", header_table);
                list = row.GetElementsByTagName("table-cell", header_table);
                XmlElement changed = (XmlElement)list.Item(col_index + 1);
                row.ReplaceChild(covered, changed);
            }
            for (int rowindex = 1; rowindex < rowSpan; rowindex++)
            {
                for (int changeIndex = 0; changeIndex < colSpan; changeIndex++)
                {
                    list = e.GetElementsByTagName("table-row", header_table);
                    row = (XmlElement)list.Item(row_index + rowindex);
                    XmlElement covered = doc.CreateElement("table:covered-table-cell", header_table);
                    list = row.ChildNodes;
                    XmlElement changed = (XmlElement)list.Item(changeIndex + childCellIndex);
                    row.ReplaceChild(covered, changed);
                }
            }
        }


        public string AddTableContentP(string text, int row, int cell)
        {
            string pname = "P" + (numP + 1).ToString();

            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;

            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            // 문단 스타일 생성
            XmlElement e1 = doc.CreateElement("style:style", header_style);

            e1.SetAttribute("name", header_style, pname);
            e1.SetAttribute("family", header_style, "paragraph");
            e1.SetAttribute("parent-style-name", header_style, "Standard");

            e.AppendChild(e1);

            // 문단 내용 생성
            list = doc.GetElementsByTagName("table", header_table);

            XmlElement table = (XmlElement)list.Item(numTable - 1);

            XmlElement tableRow = (XmlElement)table.GetElementsByTagName("table-row", header_table).Item(row);

            XmlElement tableCell = (XmlElement)tableRow.GetElementsByTagName("table-cell", header_table).Item(cell);

            XmlElement text_element = doc.CreateElement("text:p", header_text);

            text_element.SetAttribute("style-name", header_text, pname);
            foreach (char eachtext in text)
            {
                if (eachtext.Equals(' '))
                {
                    XmlElement s = doc.CreateElement("text:s", header_text);
                    text_element.AppendChild(s);
                }
                else
                {
                    XmlText t = doc.CreateTextNode(eachtext.ToString());
                    text_element.AppendChild(t);
                }
            }
            tableCell.AppendChild(text_element);

            numP++;

            return pname;
        }

        // p 생성 : null
        public string AddTableContentP(int row, int cell)
        {
            string pname = "P" + (numP + 1).ToString();

            XmlNode content = (XmlNode)root.child["content.xml"];
            XmlDocument doc = content.doc;


            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            // 문단 스타일 생성
            XmlElement e1 = doc.CreateElement("style:style", header_style);
            e1.SetAttribute("name", header_style, pname);
            e1.SetAttribute("family", header_style, "paragraph");
            e1.SetAttribute("parent-style-name", header_style, "Standard");
            e.AppendChild(e1);

            list = doc.GetElementsByTagName("table", header_table);

            XmlElement table = (XmlElement)list.Item(numTable - 1);

            XmlElement tableRow = (XmlElement)table.GetElementsByTagName("table-row", header_table).Item(row);

            XmlElement tableCell = (XmlElement)tableRow.GetElementsByTagName("table-cell", header_table).Item(cell);

            XmlElement text_element = doc.CreateElement("text:p", header_text);

            text_element.SetAttribute("style-name", header_text, pname);
            tableCell.AppendChild(text_element);

            numP++;

            return pname;
        }

        public void makeimgEntry(ImgNode img)
        {
            FolderNode meta = (FolderNode)root.child["META-INF"];
            XmlNode manifest = (XmlNode)meta.child["manifest.xml"];
            XmlDocument doc = manifest.doc;

            XmlNodeList list = doc.GetElementsByTagName("manifest", header_manifest);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement entry = doc.CreateElement("manifest:file-entry", header_manifest);
            entry.SetAttribute("full-path", header_manifest, "Pictures/" + img.name);
            entry.SetAttribute("media-type", header_manifest, "image/jpeg");

            e.AppendChild(entry);

        }

        public void makeimg(double Width, double Height, String extension, String imagePath, double x, double y, int g, int zindex, int linevertical, int verrel, int ID, Boolean isHeader, Boolean isFooter)
        {
            String geulja = "";
            String name = "이미지" + ID.ToString();
            String width = Width.ToString(); String height = Height.ToString();
            String svgx = ""; String svgy = "";
            XmlNode content;
            XmlNodeList list;
            XmlDocument doc;
            XmlElement e2; XmlElement e = null;
            if (isHeader == false && isFooter == false)
            {

                content = (XmlNode)root.child["content.xml"];
                doc = content.doc;
                list = doc.GetElementsByTagName("text", header_office);
                e = (XmlElement)list.Item(0);

            }
            else
            {
                content = (XmlNode)root.child["styles.xml"];
                doc = content.doc;

                list = doc.GetElementsByTagName("master-page", header_style);
                e2 = (XmlElement)list.Item(0);
                if (isFooter == true)
                {
                    if (numfooter == 0)
                        e = doc.CreateElement("style:footer", header_style);
                    else e = (XmlElement)(doc.GetElementsByTagName("footer", header_style).Item(0));
                }
                else if (isHeader == true)
                {
                    if (numheader == 0)
                        e = doc.CreateElement("style:header", header_style);
                    else e = (XmlElement)(doc.GetElementsByTagName("header", header_style).Item(0));
                }
                e2.AppendChild(e);
            }


            XmlElement image = doc.CreateElement("text:p", header_text);

            if (g == 0)
                geulja = "paragraph";
            else if (g == 1)
                geulja = "as-char";

            if (verrel == 2)
                svgy = (Math.Round((y + linevertical) * 2.54 / 7200, 3)).ToString();
            else
                svgy = (Math.Round(y * 2.54 / 7200, 3)).ToString(); ;

            svgx = (Math.Round(x * 2.54 / 7200, 3)).ToString();

            if (isHeader == false && isFooter == false)
                image.SetAttribute("style-name", header_text, "Standard");

            e.AppendChild(image);

            XmlElement draw = doc.CreateElement("draw:frame", header_draw);
            draw.SetAttribute("name", header_draw, name);
            if (isHeader == false && isFooter == false)
                draw.SetAttribute("style-name", header_draw, "fr" + imageStylenum);
            else
                draw.SetAttribute("style-name", header_draw, "Mfr" + imageStylenum); ;
            draw.SetAttribute("z-index", header_draw, zindex.ToString());
            draw.SetAttribute("width", header_svg, width + "cm");
            draw.SetAttribute("height", header_svg, height + "cm");
            draw.SetAttribute("y", header_svg, svgy + "cm");
            draw.SetAttribute("x", header_svg, svgx + "cm");

            draw.SetAttribute("anchor-type", header_text, geulja);

            image.AppendChild(draw);
            XmlElement ndraw = doc.CreateElement("draw:image", header_draw);
            ndraw.SetAttribute("mime-type", header_draw, "image/" + extension);
            ndraw.SetAttribute("actuate", header_xlink, "onLoad");
            ndraw.SetAttribute("show", header_xlink, "embed");
            ndraw.SetAttribute("type", header_xlink, "simple");
            ndraw.SetAttribute("href", header_xlink, imagePath);
            draw.AppendChild(ndraw);
            numImage++;
            if (isHeader == true && isFooter == false)
                numheader++;
            if (isHeader == false && isFooter == true)
                numfooter++;
        }
        public void imgstyle(int VertiRelTo,
        int VertiRelToarray,
        int HorzRelTo,
        int HorzRelToarray,
        int through, Boolean isHeader, Boolean isFooter)
        {
            String vertrel = "";
            String horzrel = "";
            String verto = "";
            String horzto = "";
            String thr = "";
            String wrap = "";
            String name = "";
            XmlNode content;

            if (isHeader == false && isFooter == false)
            {
                name = "fr" + (++imageStylenum).ToString();
                content = (XmlNode)root.child["content.xml"];
            }
            else
            {
                name = "Mfr" + (++imageStylenum).ToString();
                content = (XmlNode)root.child["styles.xml"];
            }
            XmlDocument doc = content.doc;
            XmlNodeList list = doc.GetElementsByTagName("automatic-styles", header_office);
            XmlElement e = (XmlElement)list.Item(0);

            XmlElement image = doc.CreateElement("style:style", header_style);
            if (VertiRelTo == 0)
                vertrel = "paper";
            else if (VertiRelTo == 1)
                vertrel = "page";
            else if (VertiRelTo == 2)
                vertrel = "paragraph";

            if (VertiRelToarray == 0)
                verto = "from-top";

            else if (VertiRelToarray == 1)
                if (VertiRelTo == 0 || VertiRelTo == 1)
                    verto = "middle";

                else if (VertiRelToarray == 2)
                    verto = "from-bottom";

                else if (VertiRelToarray == 3)
                    if (VertiRelTo == 0 || VertiRelTo == 1)
                        verto = "inside";

                    else if (VertiRelToarray == 4)
                        if (VertiRelTo == 0 || VertiRelTo == 1)
                            verto = "outside";
            if (HorzRelTo == 0)
                horzrel = "paper";
            else if (HorzRelTo == 1)
                horzrel = "page";
            else if (HorzRelTo == 2)
                horzrel = "paragraph";
            else if (HorzRelTo == 3)
                horzrel = "paragraph";

            if (HorzRelToarray == 0)
                horzto = "from-left";
            else if (HorzRelToarray == 1)
                if (HorzRelTo == 0 || HorzRelTo == 1)
                    horzto = "middle";

                else if (HorzRelToarray == 2)
                    horzto = "from-right";

                else if (HorzRelToarray == 3)
                    if (HorzRelTo == 0 || HorzRelTo == 1)
                        horzto = "inside";

                    else if (HorzRelToarray == 4)
                        if (HorzRelTo == 0 || HorzRelTo == 1)
                            horzto = "outside";

            image.SetAttribute("name", header_style, name);
            image.SetAttribute("family", header_style, "graphic");
            image.SetAttribute("parent-style-name", header_style, "Greaphics");
            e.AppendChild(image);
            XmlElement imgstyle = doc.CreateElement("style:graphic-properties", header_style);
            imgstyle.SetAttribute("color-mode", header_draw, "standard");
            imgstyle.SetAttribute("image-opacity", header_draw, "100%");
            imgstyle.SetAttribute("color-inversion", header_draw, "false");
            imgstyle.SetAttribute("color-gamma", header_draw, "100%");
            imgstyle.SetAttribute("color-blue", header_draw, "0%");
            imgstyle.SetAttribute("color-red", header_draw, "0%");
            imgstyle.SetAttribute("color-green", header_draw, "0%");
            imgstyle.SetAttribute("color-contrast", header_draw, "0%");
            imgstyle.SetAttribute("color-luminance", header_draw, "0%");

            imgstyle.SetAttribute("clip", header_fo, "rect(0cm,0cm,0cm,0cm)");
            imgstyle.SetAttribute("mirror", header_style, "none");
            imgstyle.SetAttribute("horizontal-rel", header_style, horzrel);
            imgstyle.SetAttribute("horizontal-pos", header_style, horzto);
            if (isFooter == true || isHeader == true)
                imgstyle.SetAttribute("vertical-rel", header_style, "page");
            else
                imgstyle.SetAttribute("vertical-rel", header_style, "page-content");
            imgstyle.SetAttribute("vertical-pos", header_style, verto);
            imgstyle.SetAttribute("wrap-contour", header_style, "false");

            if (through == 2)
            {
                wrap = "run-through";
                thr = "background";
                imgstyle.SetAttribute("wrap", header_style, wrap);
                imgstyle.SetAttribute("run-through", header_style, thr);

                imgstyle.SetAttribute("wrap-influence-on-position", header_draw, "once-concurrent");
                imgstyle.SetAttribute("number-wrapped-paragraphs", header_style, "no-limit");
            }
            else if (through == 3)
            {
                wrap = "run-through";
                thr = "foreground";
                imgstyle.SetAttribute("wrap", header_style, wrap);
                imgstyle.SetAttribute("run-through", header_style, thr);
                imgstyle.SetAttribute("number-wrapped-paragraphs", header_style, "no-limit");
                imgstyle.SetAttribute("wrap-influence-on-position", header_draw, "once-concurrent");
            }




            image.AppendChild(imgstyle);


        }


        static string content_text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        static string meta_meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
        private string mimetype;

        private XmlDocument content;
        private XmlDocument manifest;
        private XmlDocument meta;
        private XmlDocument settings;
        private XmlDocument styles;

        private XmlDocument manifest_META;

        private XmlDocument current;
    }
}
