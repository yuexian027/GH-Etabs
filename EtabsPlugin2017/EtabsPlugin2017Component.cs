using System;
using System.Collections.Generic;
using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using ETABSv17;
using Rhino.Geometry;



// In order to load the result of this wizard, you will also need to
// add the output bin/ folder of this project to the list of loaded
// folder in Grasshopper.
// You can use the _GrasshopperDeveloperSettings Rhino command for that.

namespace EtabsPlugin2017
{
    public class EtabsPlugin2017Component : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public EtabsPlugin2017Component()
          : base("EtabsPlugin2017", "EtabsV17",
              "Set up an Etabs model in V17",
              "CsiPlugin", "Etabs")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("OpenEtabs", "OpebEtabs", "Open Etabs Application", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Attached", "Attached", "Attached to an existing model", GH_ParamAccess.item);
            pManager.AddTextParameter("FilePath", "FilePath", "File path of the target Etabs model", GH_ParamAccess.item);
            pManager.AddTextParameter("ModelName", "ModelName", "Model name of the etabs file", GH_ParamAccess.item);
            pManager.AddTextParameter("ExcelPath", "ExcelPath", "File path of the Excel", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Return", "Return", "output a datatree with building info", GH_ParamAccess.tree);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            Boolean openEtabs = false;
            Boolean attaced = false;
            string filePath = "";
            string modelName = "";
            string excelPath = "";

            double x = double.NaN;
            List<string>[] modalData;


            DA.GetData(0, ref openEtabs);
            DA.GetData(1, ref attaced);
            DA.GetData(2, ref filePath);
            DA.GetData(3, ref modelName);
            DA.GetData(4, ref excelPath);

            

            // set a as true to create an etabs model
            if (openEtabs == true)
            {
                // read data from excel
                Excel myExcel = new Excel(excelPath);
                modalData = myExcel.GetExcel();

                DataTree<string> GHmodalData = ToDataTree(modalData);
                x = Etabs.GetModel(attaced, "C:\\Program Files\\Computers and Structures\\ETABS 17\\ETABS.exe", filePath, modelName,GHmodalData);

                DA.SetDataTree(0, GHmodalData);

            }
           
            

        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                // You can add image files to your project resources and access them like this:
                //return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("eff17ca9-c706-4e9e-b21b-60dd0d487709"); }
        }

        public  DataTree<string> ToDataTree(List<string>[] Data)
        {
            DataTree<string> dTree = new DataTree<string>();
            // define numbers of branches of data tree
            int rows = Data[0].Count;
            // define numbers of items in each branch
            int cols = Data.GetLength(0);
            for(int irow = 0; irow < rows; irow++)
            {
                for (int icol = 0; icol < cols; icol++)
                {
                    if(Data[icol] == null)
                    {
                        continue;
                    }
                    //dTree.Add(Data[icol][irow], new GH_Path(irow));
                    try
                    {
                        dTree.Add(Data[icol][irow].ToString(), new GH_Path(irow));
                    }
                    catch
                    {
                        continue;
                    }
                }
              
            }
            return dTree ;
        }
    }
    //CALCULATE COLUMNS OF EACH ROW
        
         
   
}


