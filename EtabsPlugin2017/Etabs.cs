using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ETABSv17;
using System.ComponentModel;
using System.Data;
using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;


namespace EtabsPlugin2017
{
    public static class Etabs
    {


        #region etbas model setup, handling whether it is a existing model
        public static int GetModel(bool AttachToInstance, string ProgramPath, string ModelDirectory, string ModelName, DataTree<string> GHmodalData)
        {
            bool SpecifyPath = true;
            //full path to the model 
            //set it to an already existing folder 
            string ModelPath = ModelDirectory + System.IO.Path.DirectorySeparatorChar + ModelName;
            try
            {
                System.IO.Directory.CreateDirectory(ModelDirectory);
            }
            catch (Exception ex)
            {
                throw new InvalidProgramException($"Could not create directory:{ModelDirectory}{ex.Message}");
            }
            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero) 
            int ret = 1;
            //dimension the ETABS Object as cOAPI type
            ETABSv17.cOAPI myETABSObject = null;


            if (AttachToInstance)
            {
                //attach to a running instance of ETABS 
                try
                {
                    //get the active ETABS object
                    myETABSObject = (ETABSv17.cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.ETABS.API.ETABSObject");

                }
                catch (Exception ex)
                {
                    throw new InvalidProgramException($"No running instance of the program found or failed to attach");
                }
            }
            else
            {
                //create API helper object
                ETABSv17.cHelper myHelper;
                try
                {
                    myHelper = new ETABSv17.Helper();
                }
                catch (Exception ex)
                {
                    throw new InvalidProgramException($"Cannot create an instance of the Helper object");

                }

                if (SpecifyPath)
                {
                    //'create an instance of the ETABS object from the specified path
                    try
                    {
                        //create ETABS object


                        myETABSObject = myHelper.CreateObject(ProgramPath);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidProgramException($"Cannot start a new instance of the program from  { ProgramPath }");

                    }
                }
                else
                {
                    //'create an instance of the ETABS object from the latest installed ETABS
                    try
                    {
                        //create ETABS object
                        myETABSObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidProgramException($"Cannot start a new instance of the program.");

                    }
                }

                //start ETABS application
                ret = myETABSObject.ApplicationStart();
            }
            #endregion

            ETABSv17.cSapModel mySapModel = default(ETABSv17.cSapModel);
            mySapModel = myETABSObject.SapModel;

            //Initialize model
            ret = mySapModel.InitializeNewModel();
            // set unit
            ret = mySapModel.SetPresentUnits(eUnits.kip_ft_F);

            #region build model
            //List<string> storyName = new List<string>() { "1","2","3","4" };
            


           // for (int i = 1; i<= 3;i++)
            //{
                //storyName[i - 1] = i.ToString();
            //}


           // ret = mySapModel.Story.SetStories_2(0, 4, ref storyName,ref storyHeight, ref toggle,ref similarToStory, ref toggle,ref storyHeight,ref color);
            int limit = GHmodalData.BranchCount;
            
            string name = string.Empty;
            string frames = string.Empty;
            //string column = string.Empty;
           
            DataTree<string> PointTree = new DataTree<string>();
            
            //string[] SlabArray = null;

            for (int i = 0; i < limit; i++)
            {
                List<string> nameList = new List<string>();
                
                for (int j = 0; j < 2; j++)
                {
                    
                    string endtPoint = GHmodalData.Branch(i)[3+j];
                    endtPoint = endtPoint.Trim(new char[] { '{', '}' });
                    string[] startPointCoord = endtPoint.Split(new char[] { ',' });
                    double x = double.Parse(startPointCoord[0]);
                    double y = double.Parse(startPointCoord[1]);
                    double z = double.Parse(startPointCoord[2]);
                    //string endPoint = GHmodalData.Branch(i)[4];
                    //double x = double.Parse(GHmodalData.Branch(i)[3]) + double.Parse(GHmodalData.Branch(i)[5] )* Math.Cos(2 * Math.PI * ivertice / vertices) / 2;
                   
                    ret = mySapModel.PointObj.AddCartesian(x, y, z, ref name);
                    nameList.Add(name);
                    
                }
                string index = GHmodalData.Branch(i)[0];
                string section = GHmodalData.Branch(i)[1];
                PointTree.AddRange(nameList, new GH_Path(i));
                ret = mySapModel.FrameObj.AddByPoint(PointTree.Branch(i)[0], PointTree.Branch(i)[1], ref frames, section,index);
                //    if( i > 1 && i <= limit)
                //    {
                //        int k = i - 1;
                //        for (int j = 0; j < vertices; j++)
                //        {
                //           
                //            //SlabTree.Add(PointTree.Branch(k-1)[j]);
                //            SlabTree.Add(PointTree.Branch(k - 1)[j], new GH_Path(i-1));
                //        }
                //        SlabArray = SlabTree.Branch(k-1).ToArray();
                //        ret = mySapModel.AreaObj.AddByPoint(vertices, ref SlabArray, ref slab);
                //    }
            }

            
            #endregion


            //Save model
            ret = mySapModel.File.Save(ModelPath);

            //Run analysis
            //ret = mySapModel.Analyze.RunAnalysis();

            //Close ETABS
            myETABSObject.ApplicationExit(false);

            return ret;
        }
       

    }
}

