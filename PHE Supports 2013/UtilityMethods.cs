using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;

namespace PHE_Supports
{
    public static class UtilityMethods
    {
        //method to align created element
        public static bool AdjustElement(Document m_document, Element createdInstance, 
            XYZ instancePoint, Pipe pipeElement, double rodLength, Curve pipeCurve)
        {
            bool rod = false, radius = false;

            try
            {
                if (pipeCurve is Line)
                {
                    Line pipeLine = (Line)pipeCurve;

                    //axis to find the y angle
                    Line yAngleAxis = m_document.Application.Create.NewLineBound(pipeLine.get_EndPoint(0), new XYZ(pipeLine.get_EndPoint(1).X, pipeLine.get_EndPoint(1).Y, pipeLine.get_EndPoint(0).Z));

                    double yAngle = XYZ.BasisY.AngleTo(yAngleAxis.Direction);

                    //axis of rotation
                    Line axis = m_document.Application.Create.NewLineBound(instancePoint, new XYZ(instancePoint.X, instancePoint.Y, instancePoint.Z + 10));

                    if (pipeCurve.get_EndPoint(0).Y > pipeCurve.get_EndPoint(1).Y)
                    {
                        if (pipeCurve.get_EndPoint(0).X > pipeCurve.get_EndPoint(1).X)
                        {
                            //rotate the created family instance to align with the pipe
                            ElementTransformUtils.RotateElement(m_document, createdInstance.Id, axis, Math.PI + yAngle);
                        }
                        else if (pipeCurve.get_EndPoint(0).X < pipeCurve.get_EndPoint(1).X)
                        {
                            //rotate the created family instance to align with the pipe
                            ElementTransformUtils.RotateElement(m_document, createdInstance.Id, axis, 2 * Math.PI - yAngle);
                        }
                    }
                    else if (pipeCurve.get_EndPoint(0).Y < pipeCurve.get_EndPoint(1).Y)
                    {
                        if (pipeCurve.get_EndPoint(0).X > pipeCurve.get_EndPoint(1).X)
                        {
                            //rotate the created family instance to align with the pipe
                            ElementTransformUtils.RotateElement(m_document, createdInstance.Id, axis, Math.PI + yAngle);
                        }
                        else if (pipeCurve.get_EndPoint(0).X < pipeCurve.get_EndPoint(1).X)
                        {
                            //rotate the created family instance to align with the pipe
                            ElementTransformUtils.RotateElement(m_document, createdInstance.Id, axis, 2 * Math.PI - yAngle);
                        }
                    }
                    else
                    {
                        //rotate the created family instance to align with the pipe
                        ElementTransformUtils.RotateElement(m_document, createdInstance.Id, axis, yAngle);
                    }

                    ParameterSet parameters = createdInstance.Parameters;

                    //set the Nominal radius and Rod height parameters
                    foreach (Parameter para in parameters)
                    {
                        if (para.Definition.Name == "PIPE RADIUS")
                        {
                            para.Set(pipeElement.get_Parameter("Outside Diameter").AsDouble() / 2.0);
                            radius = true;
                        }
                        else if (para.Definition.Name == "ROD LENGTH")
                        {
                            para.Set(rodLength);
                            rod = true;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }

            if (rod == true && radius == true)
                return true;
            else
                return false;
        }


        //method to get the points for family placement
        public static List<XYZ> GetPlacementPoints(double spacing, Curve pipeCurve, double offset, double minSpacing)
        {
            List<XYZ> points = new List<XYZ>();
            try
            {
                if (spacing == 0)
                {
                    return null;
                }

                double length = 304.8 * pipeCurve.Length;

                //ratios are required for computing the points using eveluate methods 
                //ratios should be within [0,1]
                double startRatio = offset / length;
                double endRatio = (length - offset) / length;

                points.Clear();

                double param = 0;

                //check whether the spacing is less than or greater than length of pipe
                if (length < spacing)
                {
                    //check if spacing minus offsets will be less than or 
                    //greater than minimum spacing between pipes
                    if ((length - 2 * offset) < minSpacing)
                    {
                        points.Add(pipeCurve.Evaluate(0.5, true));
                    }
                    else
                    {
                        points.Add(pipeCurve.Evaluate(startRatio, true));
                        points.Add(pipeCurve.Evaluate(endRatio, true));
                    }
                }
                else
                {
                    //compute number of splits required
                    int splits = (int)length / (int)spacing;

                    //convert splits into fractions so taht it can be compared with [0,1] range
                    param = (double)1 / (splits + 1);

                    //increment param so find required points within the range of [0,1]
                    for (double i = param; i < 1; i += param)
                    {
                        points.Add(pipeCurve.Evaluate(i, true));
                    }

                    points.Add(pipeCurve.Evaluate(startRatio, true));
                    points.Add(pipeCurve.Evaluate(endRatio, true));
                }
            }
            catch
            {
                return null;
            }
            return points;
        }


        //method to get family symbol
        public static FamilySymbol GetFamilySymbol(Family family)
        {
            FamilySymbol symbol = null;
            try
            {
                foreach (FamilySymbol s in family.Symbols)
                {
                    symbol = s;
                    break;
                }
            }
            catch
            {
                return null;
            }
            return symbol;
        }


        //method to Read data from the configuration file
        public static SpecificationData ReadConfig()
        {
            SpecificationData fileData = new SpecificationData();
            fileData.offset = -1.0;

            string configurationFile = null, assemblyLocation = null;
            try
            {
                //open configuration file
                assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;

                //convert the active file path into directory name
                if (File.Exists(assemblyLocation))
                {
                    assemblyLocation = new FileInfo(assemblyLocation).Directory.ToString();
                }

                //get parent directory of the current directory
                if (Directory.Exists(assemblyLocation))
                {
                    assemblyLocation = Directory.GetParent(assemblyLocation).ToString();
                }
            }
            catch
            {
                fileData.offset = -1.0;
                return fileData;
            }

            try
            {
                configurationFile = assemblyLocation + @"\Spacing Configuration\Configuration.txt";

                if (!(File.Exists(configurationFile)))
                {
                    MessageBox.Show("Configuration file doesn't exist!");
                    return fileData;
                }
                else
                {
                    //read all the contents of the file into a string array
                    string[] fileContents = File.ReadAllLines(configurationFile);

                    for (int i = 0; i < fileContents.Count(); i++)
                    {
                        if (fileContents[i].Contains("SelectedFamily: "))
                        {
                            fileData.selectedFamily = fileContents[i].Replace("SelectedFamily: ", "").Trim();
                        }
                        else if (fileContents[i].Contains("Discipline: "))
                        {
                            fileData.discipline = fileContents[i].Replace("Discipline: ", "").Trim();
                        }
                        else if (fileContents[i].Contains("Offest: "))
                        {
                            fileData.offset = double.Parse(fileContents[i].Replace("Offest: ", "").Trim());
                        }
                        else if (fileContents[i].Contains("Spacing: "))
                        {
                            fileData.minSpacing = double.Parse(fileContents[i].Replace("Spacing: ", "").Trim());
                        }
                        else if (fileContents[i].Contains("Support Type: "))
                        {
                            fileData.supportType = fileContents[i].Replace("Support Type: ", "").Trim();
                        }
                        else if (fileContents[i].Contains("File Location: "))
                        {
                            fileData.specsFile = fileContents[i].Replace("File Location: ", "").Trim();
                        }
                    }
                }
            }
            catch
            {
                fileData.offset = -1.0;
                return fileData;
            }
            return fileData;
        }


        //method to find the element by family name
        public static Element FindElementByName(Document m_document, Type targetType, string targetName)
        {
            try
            {
                return new FilteredElementCollector(m_document).OfClass(targetType)
                    .FirstOrDefault<Element>(e => e.Name.Equals(targetName));
            }
            catch
            {
                return null;
            }
        }


        //get all data from specification file
        public static List<string> GetAllSpecs(SpecificationData confidFileData)
        {
            try
            {
                List<string> specsData = new List<string>();
                specsData = File.ReadAllLines(confidFileData.specsFile).ToList();
                return specsData;
            }
            catch
            {
                return null;
            }
        }


        //method to retun the spacing required based on diameter
        public static double GetSpacing(Element ele, List<string> specData)
        {
            double dia = 0;
            double spacing = -1;
            Curve pipeCurve = null;

            try
            {
                if (!((Math.Round(((LocationCurve)ele.Location).Curve.get_EndPoint(0).X, 5)
                    == Math.Round(((LocationCurve)ele.Location).Curve.get_EndPoint(1).X, 5))
                    && (Math.Round(((LocationCurve)ele.Location).Curve.get_EndPoint(0).Y, 5)
                    == Math.Round(((LocationCurve)ele.Location).Curve.get_EndPoint(1).Y, 5))))
                {
                    pipeCurve = ((LocationCurve)ele.Location).Curve;

                    //converting dia dn length into millimeters
                    dia = 304.8 * ele.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();

                    //get the material of the pipe
                    string materialName = null;

                    if (ele.get_Parameter("Material") != null)
                        materialName = ele.get_Parameter("Material").AsValueString();
                    else
                        return -1;

                    for (int i = 0; i < specData.Count; i++)
                    {
                        if (specData[i].Contains(materialName))
                        {
                            string[] splitLines = null;

                            for (int j = i + 1; j < specData.Count; j++)
                            {
                                if (!specData[j].Contains(" "))
                                    return -1;
                                splitLines = specData[j].Split(' ');
                                if (splitLines[0] != "" && splitLines[1] != "")
                                {
                                    if (dia.ToString() == splitLines[0].Trim())
                                    {
                                        spacing = double.Parse(splitLines[1].Trim());
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    return -1;
                }
            }
            catch
            {
                return -1;
            }
            return spacing;
        }

        //methdo to get the bottom face of element passed to it
        public static PlanarFace GetBottomFace(Element ele)
        {
            try
            {
                Options opt = new Options();
                GeometryElement geoEle = ele.get_Geometry(opt);
                IEnumerator enu = geoEle.GetEnumerator();
                enu.Reset();

                while (enu.MoveNext())
                {
                    Solid solid = enu.Current as Solid;
                    if (null != solid)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            PlanarFace pf = face as PlanarFace;
                            if (null != pf)
                            {
                                if (pf.Normal.Z < 0)
                                {
                                    return pf;
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        //method to find the intersection of planarface with a line from the point passed
        public static double ReturnLeastZ_Value(Document doc, IList<PlanarFace> pf, XYZ xyz, Transform trans)
        {
            try
            {
                List<XYZ> intersectionPoints = new List<XYZ>();

                Line line = doc.Application.Create.NewLineBound(trans.OfPoint(xyz), new XYZ(trans.OfPoint(xyz).X, trans.OfPoint(xyz).Y, trans.OfPoint(xyz).Z + 10));

                IntersectionResultArray resultArray = null;

                foreach (PlanarFace face in pf)
                {
                    IntersectionResult iResult = null;
                    SetComparisonResult result = new SetComparisonResult();

                    try
                    {
                        result = face.Intersect(line, out resultArray);
                    }
                    catch
                    {
                        continue;
                    }

                    if (result != SetComparisonResult.Disjoint)
                    {
                        try
                        {
                            iResult = resultArray.get_Item(0);
                            intersectionPoints.Add(iResult.XYZPoint);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                XYZ minPoint = intersectionPoints.First();

                foreach (XYZ point in intersectionPoints)
                {
                    if (minPoint.Z > point.Z)
                    {
                        minPoint = point;
                    }
                }

                return (minPoint.Z - trans.OfPoint(xyz).Z);
            }
            catch
            {
                return -1;
            }
        }


        //method to return the inverse of transform from one linked file to another
        public static Transform GetInverseTransform(Document doc, Autodesk.Revit.ApplicationServices.Application app)
        {
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(RevitLinkInstance));

                Transform transform = null;

                IEnumerator enumerate = collector.GetEnumerator();

                enumerate.Reset();

                while (enumerate.MoveNext())
                {
                    RevitLinkInstance instance = enumerate.Current as RevitLinkInstance;

                    RevitLinkType linkType = (RevitLinkType)doc.GetElement(instance.GetTypeId());

                    foreach (Document docu in app.Documents)
                    {
                        if (docu.PathName.Contains(linkType.Name))
                        {
                            transform = instance.GetTotalTransform().Inverse;
                            return transform;
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        //method to reduce the list of planar faces into a smaller list where the intersection is posible
        public static List<PlanarFace> GetPossiblePlanes(List<PlanarFace> pf, XYZ xyz, Transform trans)
        {
            try
            {
                List<PlanarFace> pface = new List<PlanarFace>();

                foreach (PlanarFace face in pf)
                {
                    if (face.Origin.Z < trans.OfPoint(xyz).Z + 10 && face.Origin.Z > trans.OfPoint(xyz).Z - 10)
                    {
                        pface.Add(face);
                    }
                }
                return pface;
            }
            catch
            {
                return null;
            }
        }


        //get all structural floors ceilings and beams from the linked documents
        public static List<Element> GetStructuralElements(Autodesk.Revit.ApplicationServices.Application app, LogicalOrFilter filter)
        {
            try
            {
                List<Element> ElementsLinked = new List<Element>();

                foreach (Document d in app.Documents)
                {
                    FilteredElementCollector elements = new FilteredElementCollector(d).WherePasses(filter);
                    if (elements.Count() != 0)
                    {
                        ElementsLinked.AddRange(elements);
                    }
                }
                return ElementsLinked;
            }
            catch
            {
                return null;
            }
        }


        //filter method for structural elements
        public static LogicalOrFilter GetStructuralFilter()
        {
            //columns, beams and foundations are family instances and not types unlike walls, floor, 
            //ceiling, point load, continuoue footing, area load or line load.
            //so a logical filter is created to filter out all required structural elements

            BuiltInCategory[] builtInCategories = new BuiltInCategory[] { BuiltInCategory.OST_StructuralFraming };

            IList<ElementFilter> listOfElementFilter = new List<ElementFilter>(builtInCategories.Count());

            foreach (BuiltInCategory builtInCategory in builtInCategories)
            {
                listOfElementFilter.Add(new ElementCategoryFilter(builtInCategory));
            }

            LogicalOrFilter categoryFilter = new LogicalOrFilter(listOfElementFilter);

            LogicalAndFilter familyInstanceFilter = new LogicalAndFilter(categoryFilter, new ElementClassFilter(typeof(FamilyInstance)));

            IList<ElementFilter> typeFilter = new List<ElementFilter>();

            typeFilter.Add(new ElementClassFilter(typeof(Floor)));

            typeFilter.Add(new ElementClassFilter(typeof(Ceiling)));

            typeFilter.Add(familyInstanceFilter);

            LogicalOrFilter classFilter = new LogicalOrFilter(typeFilter);

            return classFilter;
        }
    }
}
