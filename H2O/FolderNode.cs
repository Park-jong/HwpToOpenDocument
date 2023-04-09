using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class FolderNode : FileNode
    {
        public FolderNode()
        {

        }

        public FolderNode(string name)
        {
            this.name = name;
        }

        public FolderNode(string name, FileNode parent)
        {
            this.name = name;
            this.parent = parent;

            parent.AddChild(this);
        }
    }
}
