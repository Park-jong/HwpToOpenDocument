using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class FileNode
    {
        public FileNode parent;
        public Hashtable child;

        public string path;

        public string name;

        public FileNode()
        {
        }

        public FileNode(string name)
        {
            this.name = name;
        }

        public FileNode(string name, FileNode parent)
        {
            this.name = name;
            this.parent = parent;

            parent.AddChild(this);
        }

        public void AddChild(FileNode node)
        {
            if (child == null) { child = new Hashtable(); }
            if (!child.ContainsKey(node.name))
            {
                child.Add(node.name, (FileNode)node);
                node.path = this.path + @"\" + node.name;
            }
        }
    }
}