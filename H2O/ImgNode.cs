using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace WindowsFormsApp1
{
    public class ImgNode : FileNode
    {
        public Image img;

        public ImgNode()
        {
        }

        public ImgNode(string name, FileNode parent)
        {
            this.name = name;
            this.parent = parent;

            parent.AddChild(this);
        }
    }
}