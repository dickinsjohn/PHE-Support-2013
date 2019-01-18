using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;

using Security_Check;

namespace PHE_Supports
{
    //Transaction assigned as automatic
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Automatic)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    //Creating an external command to provide supports
    public class PHESupports : IExternalCommand
    {
        //instances to store application and the document
        UIDocument m_document = null;

        //for storing the selected elements
        ElementSet eleSet = null;

        //tostore the family name in config file, corresopnding family and its family symbol
        string FamilyName = null;
        FamilySymbol symbol = null;
        Family family = null;

        //to check for the security
        bool security = false;

        //for storing the rod length computed, offest fro ends and required spacing
        double rodLength = 0.0, spacing = 0.0;

        //level property of the pipe selected
        Level pipeLevel = null;

        //variable to store element created
        Element tempEle = null;

        //to store the pipe curve of the pipe under selection
        Curve pipeCurve = null;

        //lists to store, spacing details, points for placing instances and list of created instances
        List<XYZ> points = new List<XYZ>();
        List<string> specsData = new List<string>();

        //failure lists
        List<ElementId> FailedToPlace = new List<ElementId>();
        List<ElementId> createdElements = new List<ElementId>();

        //default execute method required by the IExternalCommand class
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                //call to the security check method to check for authentication
                security = SecurityLNT.Security_Check();
                if (security == false)
                {
                    return Result.Succeeded;
                }


                //get the application data
                Autodesk.Revit.ApplicationServices.Application app = commandData.Application.Application;

                //read the config file
                SpecificationData configFileData = UtilityMethods.ReadConfig();

                //check for discipline
                if (configFileData.discipline != "PHE")
                {
                    TaskDialog.Show("Failed!", "Sorry! Plug-in Not intended for your discipline.");
                    return Result.Succeeded;
                }

                //exception handled
                if (configFileData.offset == -1.0)
                {
                    MessageBox.Show("Configuration data not found!");
                    return Result.Succeeded;
                }

                //get all data from the specification file
                specsData = UtilityMethods.GetAllSpecs(configFileData);

                //exception handled
                if (specsData == null)
                {
                    MessageBox.Show("Specifications not found!");
                    return Result.Succeeded;
                }

                //open  the active document in revit
                m_document = commandData.Application.ActiveUIDocument;

                //get the selected element set
                eleSet = m_document.Selection.Elements;

                if (eleSet.IsEmpty)
                {
                    MessageBox.Show("Please select pipes before executing the Add-in!");
                    return Result.Succeeded;
                }

                //call to method to get the transform required
                Transform transform = UtilityMethods.GetInverseTransform(m_document.Document, commandData.Application.Application);

                if (transform == null)
                {
                    MessageBox.Show("Sorry! Couldn't find a possible transform.");
                    return Result.Succeeded;
                }

                //get family name from config data
                FamilyName = configFileData.selectedFamily;

                //check if family exists
                family = UtilityMethods.FindElementByName(m_document.Document, typeof(Family), FamilyName) as Family;

                //if existing
                if (family == null)
                {
                    MessageBox.Show("Please load the family into the project and re-run the Add-in.");
                    return Result.Succeeded;
                }

                //get the family symbol
                symbol = UtilityMethods.GetFamilySymbol(family);

                //exception handled
                if (family == null)
                {
                    MessageBox.Show("No family symbol for the family you specified!");
                    return Result.Succeeded;
                }

                //create a logical filter for all the structural elements
                LogicalOrFilter filter = UtilityMethods.GetStructuralFilter();

                //if no filter is returned
                if (filter == null)
                {
                    return Result.Succeeded;
                }

                //get the structural elements from the documents
                List<Element> structuralElements = new List<Element>();

                structuralElements = UtilityMethods.GetStructuralElements(app, filter);

                if (structuralElements == null)
                {
                    MessageBox.Show("Sorry! No structural element found");
                    return Result.Succeeded;
                }

                //list to store all the planar botom faces of the structural elements
                List<PlanarFace> pFace = new List<PlanarFace>();

                //find and add bottom planar faces to the list
                foreach (Element e in structuralElements)
                {
                    PlanarFace pf = UtilityMethods.GetBottomFace(e);
                    if (pf != null)
                        pFace.Add(pf);
                }

                //clear the structural elements list as it is no longer required
                structuralElements.Clear();

                //control variable for removing unwanted elements from the planar faces list
                int flag = 0;

                //iterate through all the selected elements
                foreach (Element ele in eleSet)
                {
                    //check whether the selected element is of type pipe
                    if (ele is Pipe)
                    {
                        //get the location curve of the pipe
                        pipeCurve = ((LocationCurve)ele.Location).Curve;

                        //if the length of pipe curve obtained is zero skip that pipe
                        if (pipeCurve.Length == 0)
                        {
                            FailedToPlace.Add(ele.Id);
                            continue;
                        }

                        //remove unwanted planes from the list
                        if (flag == 0)
                        {
                            //remove unwanted faces from the list (this works only once)
                            pFace = UtilityMethods.GetPossiblePlanes(pFace, pipeCurve.get_EndPoint(0), transform);
                            flag = -1;
                        }
                        
                        //if no plane is found for intersect to work
                        if (pFace.Count == 0)
                        {
                            MessageBox.Show("Sorry! No structural element found");
                            return Result.Succeeded;
                        }

                        //from the specification file, get the spacing corresponding to the pipe diameter
                        spacing = UtilityMethods.GetSpacing(ele, specsData);

                        //check if the spacing returned is -1
                        if (spacing == -1)
                        {
                            FailedToPlace.Add(ele.Id);
                            continue;
                        }

                        //get the points for placing the family instances
                        points = UtilityMethods.GetPlacementPoints(spacing, pipeCurve,
                            1000 * configFileData.offset, 1000 * configFileData.minSpacing);

                        //check if the points is null exception
                        if (points == null)
                        {
                            FailedToPlace.Add(ele.Id);
                            continue;
                        }

                        //get the pipe level
                        pipeLevel = ele.Level;

                        //iterate through all the points for placing the family instances
                        foreach (XYZ point in points)
                        {
                            try
                            {
                                //create the instances at each points
                                tempEle = m_document.Document.Create.NewFamilyInstance
                                    (point, symbol, ele, pipeLevel, StructuralType.NonStructural);
                                createdElements.Add(tempEle.Id);
                            }
                            catch
                            {
                                FailedToPlace.Add(ele.Id);
                                continue;
                            }

                            //find the rod length required 
                            rodLength = UtilityMethods.ReturnLeastZ_Value(m_document.Document, pFace, point, transform);

                            if (rodLength == -1)
                            {
                                FailedToPlace.Add(ele.Id);
                                createdElements.Remove(tempEle.Id);
                                m_document.Document.Delete(tempEle.Id);
                                continue;
                            }

                            //adjust the newly created element properties based on the rodlength, 
                            //orientation and dia of pipe
                            if (!UtilityMethods.AdjustElement(m_document.Document,
                                tempEle, point, (Pipe)ele, rodLength, pipeCurve))
                            {
                                FailedToPlace.Add(ele.Id);
                                createdElements.Remove(tempEle.Id);
                                m_document.Document.Delete(tempEle.Id);
                                continue;
                            }
                        }
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
            throw new NotImplementedException();
        }
    }
}
