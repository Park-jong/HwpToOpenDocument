using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml;

namespace WindowsFormsApp1
{
    public class XmlNode : FileNode
    {
        public XmlDocument doc;

        public XmlNode()
        {

        }

        public XmlNode(string name)
        {
            this.name = name;

            doc = new XmlDocument();
        }

        public XmlNode(string name, FileNode parent)
        {
            this.name = name;
            this.parent = parent;

            parent.AddChild(this);

            doc = new XmlDocument();
        }
        public void LoadXml(string loadPath)
        {
            if (loadPath.Contains("mimetype"))
                return;

            try
            {
                doc.Load(loadPath);
            }
            catch(XmlException e)
            {
                XmlElement root = doc.CreateElement("root");
                doc.AppendChild(root);
            }
        }
    }
}
