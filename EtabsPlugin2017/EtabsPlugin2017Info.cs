using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace EtabsPlugin2017
{
    public class EtabsPlugin2017Info : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "EtabsPlugin2017";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("2c669eb9-0a55-4af6-bad8-5094f40ff117");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "";
            }
        }
    }
}
