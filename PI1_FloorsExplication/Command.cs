using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using PI1_CORE;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PI1_FloorsExplication
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    // Start command class.
    public class Command : IExternalCommand
    {
        #region public methods

        /// <summary>
        /// Overload this method to implement and external command within Revit.
        /// </summary>
        /// <param name="commandData">An ExternalCommandData object which contains reference to Application and View
        /// needed by external command.</param>
        /// <param name="message">Error message can be returned by external command. This will be displayed only if the command status
        /// was "Failed".  There is a limit of 1023 characters for this message; strings longer than this will be truncated.</param>
        /// <param name="elements">Element set indicating problem elements to display in the failure dialog.  This will be used
        /// only if the command status was "Failed".</param>
        /// <returns>
        /// The result indicates if the execution fails, succeeds, or was canceled by user. If it does not
        /// succeed, Revit will undo any changes made by the external command.
        /// </returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Common settings for operations.
            ElementCategoryFilter floorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
            ThinLinesOptions.AreThinLinesEnabled = false;
            XYZ vectorZ = XYZ.BasisZ.Negate();

            // Create options for saving view as image.
            ImageExportOptions options = new ImageExportOptions();
            options.ExportRange = ExportRange.CurrentView;
            options.ImageResolution = ImageResolution.DPI_600;
            options.ZoomType = ZoomFitType.Zoom;
            options.Zoom = 100;

            // Get floor types of the currrent document.
            var floorTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsElementType()
                .Where(floor => floor.Name.Contains("Тип"));

            // Get rooms of the current document.
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms);

            // Get any 3D View.
            View3D view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .ToList()
                .First();

            // Get floor type scheme legends.
            var legends = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(view => view.ViewType == ViewType.Legend && view.Name.Contains("(P)_FL_Тип"));

            // Get all images names in project.
            var legendImageNames = from imageView in 
                        new FilteredElementCollector(doc)
                        .OfClass(typeof(ImageType))
                        .Cast<ImageType>()
                        .Where(imageView => imageView.Name.Contains("(P)_FL_Тип")) 
                        select imageView.Name;

            #region create floor scheme

            // Create visualized views.
            foreach (View legend in legends)
            {
                // Check if there already exists image of the scheme.
                if (!legendImageNames.Contains(legend.Name))
                {
                    // Open view for visualization.
                    uidoc.ActiveView = legend;
                    using (Transaction t = new Transaction(doc, $"Создание изображения {legend.Name}"))
                    {
                        t.Start();

                        // Get list of opened views.
                        IList<UIView> openedAllViews = uidoc.GetOpenUIViews();

                        // Set name and create view. 
                        options.ViewName = doc.ActiveView.Name;
                        doc.SaveToProjectAsImage(options);

                        // Close opened view.
                        foreach (UIView uiView in openedAllViews)
                        {
                            if (uiView.ViewId == legend.Id)
                            {
                                uiView.Close();
                            }
                        }

                        t.Commit();
                    }
                }
            }

            #endregion

            using (Transaction t = new Transaction(doc, "Создание экспликации полов"))
            {
                t.Start();

                foreach (Element floorType in floorTypes)
                {

                    #region set material description

                    // Get structural layers of the floor type.
                    HostObjAttributes hoas = floorType as HostObjAttributes;
                    IList<CompoundStructureLayer> layers = hoas.GetCompoundStructure().GetLayers();

                    string allMatForSpec = string.Empty;
                    int c = 1;
                    foreach (CompoundStructureLayer layer in layers)
                    {
                        // Get layer width and material description of the layer.
                        double width = UnitUtils.ConvertFromInternalUnits(layer.Width, DisplayUnitType.DUT_MILLIMETERS);
                        Element material = doc.GetElement(layer.MaterialId);
                        string matDescription = material.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString();

                        // Check the width of the layer.
                        string strForInserting = string.Empty;
                        if (width == 0)
                        {
                            strForInserting = matDescription;
                        }
                        else
                        {
                            strForInserting = matDescription + " - " + Convert.ToString(width) + "мм";
                        }

                        // Create name of the needed parameter and get it.
                        string parName = "Dyn_Материал_" + Convert.ToString(c);
                        Parameter parameter = floorType.LookupParameter(parName);

                        // Check if parameter exist and if not, close the app.
                        if (parameter == null)
                        {
                            TaskDialog.Show("Предупреждение", $"Парметр {parName} не создан для категории Перекрытия");
                            return Result.Cancelled;
                        }

                        // Set parameter value.
                        floorType.LookupParameter(parName).Set(strForInserting);

                        // Add material description with width to one string.
                        allMatForSpec += Convert.ToString(c) + ". " + strForInserting + "\n";

                        c++;
                    }

                    // Add substring in string with all materials and set common material layer description to parameter.
                    string lastMeterial = Convert.ToString(layers.Count + 1) + ". Ж/б плита см. раздел КР";
                    floorType.LookupParameter("Dyn_Материалы_Спецификация").Set(allMatForSpec + lastMeterial);

                    #endregion

                    #region set room numbers

                    // Get all room numbers list of room that placed on current floor type.
                    List<string> roomNumbers = new List<string>();
                    foreach (Element room in rooms)
                    {
                        // Get setting point of the room.
                        LocationPoint locPoint = room.Location as LocationPoint;
                        XYZ point = locPoint.Point;

                        // Get floor type that intersects with vector from setting point of the room in -Z direction.
                        ReferenceIntersector intersector = new ReferenceIntersector(floorFilter, FindReferenceTarget.Element, view3D);
                        Reference reference = intersector.FindNearest(point, vectorZ).GetReference();

                        if (reference == null)
                        {
                            continue;
                        }

                        Element elementType = doc.GetElement(doc.GetElement(reference).GetTypeId());

                        // Check getted flor type if it is current floor type.
                        if (elementType.Name == floorType.Name)
                        {
                            string roomNumber = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();
                            roomNumbers.Add(roomNumber);
                        }
                    }

                    // Sort room numbers ascending.
                    roomNumbers.OrderBy(st => st);

                    // Create a string of room numbers.
                    string roomNumbersValue = string.Join(", ", roomNumbers);
                    
                    // Check if there is a room numbers for current floor type
                    // and if it is set the string in parameter.
                    if (roomNumbersValue != string.Empty)
                    {
                        Parameter par = floorType.LookupParameter("Dyn_НомерПомещения");
                        par.Set(roomNumbersValue);
                    }

                    #endregion

                    #region set floor scheme

                    // Get created images.
                    var legendImages = new FilteredElementCollector(doc)
                        .OfClass(typeof(ImageType))
                        .Cast<ImageType>()
                        .Where(imageView => imageView.Name.Contains("(P)_FL_Тип"));

                    // Parse the name of floor type and image to define binding.
                    string floorTypeNumber = (floorType.Name.Split(new string[] { "Тип" }, StringSplitOptions.None))[1];
                    foreach (ImageType legendImage in legendImages)
                    {
                        string legendTypeNumber = (legendImage.Name.Split(new string[] { "Тип" }, StringSplitOptions.None))[1];
                        if (floorTypeNumber == legendTypeNumber)
                        {
                            // Set the image to floot type.
                            floorType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_IMAGE).Set(legendImage.Id);
                        }
                    }

                    #endregion


                }

                #region create schedule

                var floorSchedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Floors));
                var fsDefinition = floorSchedule.Definition;

                floorSchedule.get_Parameter(BuiltInParameter.VIEW_NAME).Set("(P)_Fl_Экспликация полов");
                fsDefinition.IsItemized = false;

                foreach (SchedulableField field1 in fsDefinition.GetSchedulableFields())
                {
                    if (field1.GetName(doc) == "Dyn_НомерПомещения")
                    {
                        var Field1 = fsDefinition.AddField(field1);
                        Field1.ColumnHeading = "Номер помещения";
                        Field1.SheetColumnWidth = UnitUtils.ConvertToInternalUnits(25, DisplayUnitType.DUT_MILLIMETERS);
                        Field1.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                        foreach (SchedulableField field2 in fsDefinition.GetSchedulableFields())
                        {
                            if (field2.GetName(doc) == "Маркировка типоразмера")
                            {
                                var Field2 = fsDefinition.AddField(field2);
                                Field2.ColumnHeading = "Тип пола";
                                Field2.SheetColumnWidth = UnitUtils.ConvertToInternalUnits(15, DisplayUnitType.DUT_MILLIMETERS);
                                Field2.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                                foreach (SchedulableField field3 in fsDefinition.GetSchedulableFields())
                                {
                                    if (field3.GetName(doc) == "Изображение типоразмера")
                                    {
                                        var Field3 = fsDefinition.AddField(field3);
                                        Field3.ColumnHeading = "Схема пола";
                                        Field3.SheetColumnWidth = UnitUtils.ConvertToInternalUnits(50, DisplayUnitType.DUT_MILLIMETERS);
                                        foreach (SchedulableField field4 in fsDefinition.GetSchedulableFields())
                                        {
                                            if (field4.GetName(doc) == "Dyn_Материалы_Спецификация")
                                            {
                                                var Field4 = fsDefinition.AddField(field4);
                                                Field4.ColumnHeading = "Данные элементов пола (наименование, толщина, основание и др.), мм";
                                                Field4.SheetColumnWidth = UnitUtils.ConvertToInternalUnits(85, DisplayUnitType.DUT_MILLIMETERS);
                                                foreach (SchedulableField field5 in fsDefinition.GetSchedulableFields())
                                                {
                                                    if (field5.GetName(doc) == "Площадь")
                                                    {
                                                        var Field5 = fsDefinition.AddField(field5);
                                                        Field5.ColumnHeading = "Площадь, м²";
                                                        Field5.SheetColumnWidth = UnitUtils.ConvertToInternalUnits(20, DisplayUnitType.DUT_MILLIMETERS);
                                                        Field5.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                                                        Field5.DisplayType = ScheduleFieldDisplayType.Totals;
                                                        foreach (SchedulableField field6 in fsDefinition.GetSchedulableFields())
                                                        {
                                                            if (field6.GetName(doc) == "Группа модели")
                                                            {
                                                                var Field6 = fsDefinition.AddField(field6);
                                                                Field6.IsHidden = true;
                                                                var groupFilter = new ScheduleFilter(Field6.FieldId, ScheduleFilterType.NotEqual, "Перекрытие");
                                                                fsDefinition.AddFilter(groupFilter);
                                                                break;
                                                            }
                                                        }

                                                        var areaFilter = new ScheduleFilter(Field5.FieldId, ScheduleFilterType.GreaterThan, 0.00);
                                                        fsDefinition.AddFilter(areaFilter);

                                                        break;
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }

                                var markSort = new ScheduleSortGroupField(Field2.FieldId, ScheduleSortOrder.Ascending);
                                fsDefinition.AddSortGroupField(markSort);

                                break;
                            }
                        }
                        break;
                    }
                }

                TableData fsTableData = floorSchedule.GetTableData();
                TableSectionData header = fsTableData.GetSectionData(SectionType.Header);
                header.SetCellText(0, 0, "Экспликация полов");

                #endregion

                t.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Gets the path of the current command.
        /// </summary>
        /// <returns></returns>
        public static string GetPath()
        {
            return typeof(Command).Namespace + "." + nameof(Command);
        }

        #endregion
    }
}
