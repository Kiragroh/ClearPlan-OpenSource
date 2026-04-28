using ClearPlan.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using VMS.TPS.Common.Model.API;

using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;

using VMS.TPS.Common.Model.Types;
using Series = OxyPlot.Series.Series;

namespace ClearPlan
{

    public class MainViewModel : ViewModelBase
    {
        public Patient Patient { get; set; }
        public User User { get; set; }

        //public List<string> SelectedStructures { get; set; }
        public PlanSetup PlanSetup { get; set; }
        public PlanSum PlanSum{ get; set; }
        public Image Image { get; set; }
        public StructureSet StructureSet { get; set; }
        //public Window MainWindow { get; set; }
        public string ScriptVersion { get; set; }
        public string Title { get; set; }
        public ConstraintViewModel ActiveConstraintPath { get; set; }
        public PlanningItemViewModel ActivePlanningItem { get; set; }
        public List<ErrorViewModel> ErrorGrid { get; set; }
        public List<RefViewModel> RefGrid { get; set; }

        //public List<DVHstatViewModel> DVHstatGrid { get; set; }

        public ObservableCollection<PQMSummaryViewModel> PqmSummaries { get; set; }

       // public ObservableCollection<StructureViewModel> _FoundStructureList;
       // public ObservableCollection<StructureViewModel> FoundStructureList
       // {
        //    get { return _FoundStructureList; }
        //    set
         //   {
         //       _FoundStructureList = value;
        //        NotifyPropertyChanged("FoundStructureList");
       //     }
       // }
        PQMSummaryViewModel[] Objectives { get; set; }
        public ObservableCollection<PlanningItemViewModel> PlanningItemList { get; set; }
        public ObservableCollection<ConstraintViewModel> ConstraintComboBoxList { get; set; }
        public ObservableCollection<PlanningItemDetailsViewModel> PlanningItemSummaries { get; set; }
        public ObservableCollection<StructureViewModel> StructureList { get; set; }
        public double SliderValue { get; set; }
        public Model3DGroup ModelGroup { get; set; }
        public Point3D isoctr { get; set; }
        public Point3D cameraPosition { get; set; }
        public Vector3D upDir { get; set; }
        public Vector3D lookDir { get; set; }


        public MainViewModel(User user, Patient patient, string scriptVersion, ObservableCollection<PlanningItemViewModel> planningItemList, PlanningItemViewModel planningItem, PlanningItem pItem)
        {
            if (pItem is PlanSetup)
            {
                _plan = pItem as PlanSetup;
            }
            else
            {
                _psum = pItem as PlanSum;
            }

            Structures = GetPlanStructures();
            PlotModel = CreatePlotModel();
            
            ActivePlanningItem = planningItem;
            Patient = patient;
            User = user;
            Image = ActivePlanningItem.PlanningItemImage;
            StructureSet = ActivePlanningItem.PlanningItemStructureSet;
            //int Fx = ActivePlanningItem.
            var settings = ClearPlanSettings.Load();
            DirectoryInfo constraintDir = new DirectoryInfo(settings.GetConstraintTemplatesDirectory());
            string defaultConstraintPath = SelectDefaultConstraintPath(planningItem, settings, constraintDir);
            ActiveConstraintPath = new ConstraintViewModel(defaultConstraintPath);
            PlanningItemList = planningItemList;
            StructureList = StructureSetListViewModel.GetStructureList(StructureSet);
            ConstraintComboBoxList = ConstraintListViewModel.GetConstraintList(constraintDir.ToString());
            //GetPQMSummaries(ActiveConstraintPath, ActivePlanningItem, Patient);
            //PqmSummaries = new ObservableCollection<PQMSummaryViewModel>();
            ErrorGrid = GetErrors(ActivePlanningItem, Patient);
            RefGrid = GetRefs(ActivePlanningItem, Patient);
            //DVHstatGrid = GetDVHstats(ActivePlanningItem, Patient);
            Title = GetTitle(patient, scriptVersion);
            ModelGroup = new Model3DGroup();
            SliderValue = 0;
            upDir = new Vector3D(0, -1, 0);
            lookDir = new Vector3D(0, 0, 1);
            isoctr = new Point3D(0, 0, 0);  //just to initalize
            cameraPosition = new Point3D(0, 0, -4500);
            PlanningItemSummaries = GetPlanningItemSummary(ActivePlanningItem, PlanningItemList);
            //NotifyPropertyChanged("Structure");
        }

        private static string SelectDefaultConstraintPath(PlanningItemViewModel planningItem, ClearPlanSettings settings, DirectoryInfo constraintDir)
        {
            foreach (string fileName in GetPreferredConstraintTemplates(planningItem, settings))
            {
                string fullPath = Path.Combine(constraintDir.FullName, fileName);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            string firstAvailableTemplate = constraintDir.Exists
                ? Directory.EnumerateFiles(constraintDir.FullName, "*.csv")
                    .OrderBy(file => Path.GetFileName(file), StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault()
                : null;

            return firstAvailableTemplate ?? settings.GetConstraintTemplatePath(settings.Paths.DefaultConventionalTemplate);
        }

        private static IEnumerable<string> GetPreferredConstraintTemplates(PlanningItemViewModel planningItem, ClearPlanSettings settings)
        {
            if (planningItem.PlanningItemFx == "PlanSum")
            {
                return new[]
                {
                    settings.Paths.DefaultPlanSumTemplate,
                    settings.Paths.DefaultConventionalTemplate
                };
            }

            if (IsHypofractionatedCandidate(planningItem))
            {
                return new[]
                {
                    settings.Paths.DefaultHypofractionatedTemplate,
                    settings.Paths.DefaultConventionalTemplate
                };
            }

            return new[]
            {
                settings.Paths.DefaultConventionalTemplate,
                settings.Paths.DefaultHypofractionatedTemplate,
                settings.Paths.DefaultPlanSumTemplate
            };
        }

        private static bool IsHypofractionatedCandidate(PlanningItemViewModel planningItem)
        {
            int numberOfFractions;
            if (int.TryParse(planningItem.PlanningItemFx, out numberOfFractions) && numberOfFractions <= 5)
            {
                return true;
            }

            return planningItem.PlanningItemDoseFx >= 4;
        }
        #region PlanSelector
        #endregion PlanSelector
        #region dvh
        private readonly PlanSetup _plan;
        private readonly PlanSum _psum;

        
        public IEnumerable<Structure> Structures { get; private set; }

        public PlotModel PlotModel { get; private set; }

        public void AddDvhCurve(Structure structure)
        {
            var dvh = CalculateDvh(structure);
            PlotModel.Series.Add(CreateDvhSeries(structure.Id, dvh));
            UpdatePlot();
        }

        public void RemoveDvhCurve(Structure structure)
        {
            var series = FindSeries(structure.Id);
            //System.Windows.MessageBox.Show(series.Tag.ToString());
            PlotModel.Series.Remove(series);
            UpdatePlot();
        }

        public void ExportPlotAsPdf(string filePath)
        {
            using (var stream = File.Create(filePath))
            {
                PdfExporter.Export(PlotModel, stream, 650, 300);
            }
        }

        //public void ExportPlotAsBitmap(string filePath)
        //{
        //    using (var stream = File.Create(filePath))
        //    {
        //        PdfExporter.Export(PlotModel, stream, 600, 300);
        //    }
        //}

        private IEnumerable<Structure> GetPlanStructures()
        {
            if (_psum == null)
            {
                return _plan.StructureSet != null
                    ? _plan.StructureSet.Structures.Where(x => !x.IsEmpty && x.HasSegment && x.DicomType.ToUpper() != "SUPPORT").OrderBy(x => x.Id)
                    : null;
            }
            else
            {
                return _psum.StructureSet != null
                   ? _psum.StructureSet.Structures.Where(x => !x.IsEmpty && x.HasSegment && x.DicomType.ToUpper() != "SUPPORT").OrderBy(x => x.Id)
                   : null;
            }
        }

        private PlotModel CreatePlotModel()
        {
            var plotModel = new PlotModel
            {
                PlotAreaBackground = OxyColor.FromAColor(120, OxyColors.LightGray)
            };
            AddAxes(plotModel);
            SetupLegend(plotModel);
            plotModel.Padding = new OxyThickness(0,10,5,0);
            return plotModel;
        }
        private void SetupLegend(PlotModel pm)
        {
            pm.LegendBorder = OxyColors.Black;
            pm.LegendBackground = OxyColor.FromAColor(120, OxyColors.LightGray);
            pm.LegendPosition = LegendPosition.RightTop;
            pm.LegendOrientation = LegendOrientation.Vertical;
            pm.LegendPlacement = LegendPlacement.Outside;
            //pm.LegendMaxHeight = 400;            
            
        }

        private static void AddAxes(PlotModel plotModel)
        {
            plotModel.Axes.Add(new LinearAxis
            {
                Title = "Dose [Gy]",
                TitleFontSize = 14,
                TitleFontWeight = FontWeights.Bold,
                AxisTitleDistance = 15,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Solid,
                Position = AxisPosition.Bottom,
                Minimum = 0,
                AbsoluteMinimum = 0,
               
            });

            plotModel.Axes.Add(new LinearAxis
            {
                Title = "Volume [%]",
                TitleFontSize = 14,
                TitleFontWeight = FontWeights.Bold,
                AxisTitleDistance = 15,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Solid,
                Position = AxisPosition.Left,
                Minimum = 0,
                Maximum = 100.5,
                AbsoluteMinimum = 0,
                IsZoomEnabled = false
        });
        }

        private DVHData CalculateDvh(Structure structure)
        {
            if (_psum == null)
            {
                return _plan.GetDVHCumulativeData(structure,
                DoseValuePresentation.Absolute,
                VolumePresentation.Relative, 0.01);
            }
            else
            {
                return _psum.GetDVHCumulativeData(structure,
                DoseValuePresentation.Absolute,
                VolumePresentation.Relative, 0.01);
            }
        }

        private Series CreateDvhSeries(string structureId, DVHData dvh)
        {
            var series = new LineSeries
            {
                Title = structureId.Length>16?structureId.Substring(0,14)+"..": structureId,
                Tag = structureId.Length > 26 ? structureId.Substring(0, 14) + ".." : structureId,
                Color = GetStructureColor(structureId),
                StrokeThickness = GetLineThickness(structureId),
                LineStyle = GetLineStyle(structureId)
            };
            var points = dvh.CurveData.Select(CreateDataPoint);
            series.Points.AddRange(points);
            return series;
        }

        private OxyColor GetStructureColor(string structureId)
        {

            var structures = _psum == null ? _plan.StructureSet.Structures : _psum.StructureSet.Structures;

            var structure = structures.First(x => x.Id == structureId);
            var color = structure.Color;
            return OxyColor.FromRgb(color.R, color.G, color.B);
        }
        private double GetLineThickness(string structureId)
        {
            if (structureId.ToUpper().StartsWith("PTV"))
                return 2;
            return 1.5;
        }
        private LineStyle GetLineStyle(string structureId)
        {
            if (structureId.ToUpper().StartsWith("Z"))
                return LineStyle.Dash;
            return LineStyle.Solid;
        }

        private DataPoint CreateDataPoint(DVHPoint p)
        {
            return new DataPoint(p.DoseValue.Dose, p.Volume);
        }

        private Series FindSeries(string structureId)
        {
            return PlotModel.Series.FirstOrDefault(x =>
                (string)x.Tag == structureId);
        }

        private void UpdatePlot()
        {
            PlotModel.InvalidatePlot(true);
        }
        #endregion dvh

        public string GetTitle(Patient patient, string scriptVersion)
        {
            Title = patient.Name + " - " + "ClearPlan v." + scriptVersion;
            return Title;
        }

        public void GetPQMSummaries(ConstraintViewModel constraintPath, PlanningItemViewModel planningItem, Patient patient)
        {
            PqmSummaries = new ObservableCollection<PQMSummaryViewModel>();
            StructureSet structureSet = planningItem.PlanningItemStructureSet;
            //PlanSetup plaSetup = planningItem.PlanningItemId;
            Structure evalStructure;
            ObservableCollection<PQMSummaryViewModel> pqmSummaries = new ObservableCollection<PQMSummaryViewModel>();
            ObservableCollection<StructureViewModel> foundStructureList = new ObservableCollection<StructureViewModel>();
            var calculator = new PQMSummaryCalculator();
            Objectives = calculator.GetObjectives(constraintPath);
            if (planningItem.PlanningItemObject is PlanSum)
            {
                var waitWindowPQM = new WaitWindowPQM();
                PlanSum plansum = (PlanSum)planningItem.PlanningItemObject;
                if (plansum.IsDoseValid() == true)
                {
                    waitWindowPQM.ShowInTaskbar = false;
                    waitWindowPQM.Show();
                    foreach (PQMSummaryViewModel objective in Objectives)
                    {
                        evalStructure = calculator.FindStructureFromAlias(structureSet, planningItem, objective.TemplateId, objective.TemplateAliases, objective.TemplateCodes, objective.TemplateType);
                        if (evalStructure != null)
                        {
                            var evalStructureVM = new StructureViewModel(evalStructure);
                            var obj = calculator.GetObjectiveProperties(objective, planningItem, structureSet, evalStructureVM);
                            PqmSummaries.Add(obj);
                            NotifyPropertyChanged("Structure");
                        }
                    }
                    waitWindowPQM.Close();
                }
            }
            if (planningItem.PlanningItemObject is PlanSetup) //is plansetup
            {
                var waitWindowPQM = new WaitWindowPQM();

                PlanSetup planSetup = (PlanSetup)planningItem.PlanningItemObject;
                if (planSetup.IsDoseValid() == true)
                {
                    waitWindowPQM.ShowInTaskbar = false;
                    waitWindowPQM.Show();
                    //waitWindowPQM.Topmost = true;
                    foreach (PQMSummaryViewModel objective in Objectives)
                    {
                        evalStructure = calculator.FindStructureFromAlias(structureSet, planningItem, objective.TemplateId, objective.TemplateAliases, objective.TemplateCodes, objective.TemplateType);
                        if (evalStructure != null)
                        {
                            /*if (evalStructure.StructureCodeInfos.FirstOrDefault().Code != null)
                            {
                                if (evalStructure.StructureCodeInfos.FirstOrDefault().Code.Contains("PTV") == true)
                                {
                                    foreach (Structure s in structureSet.Structures)
                                    {
                                        if (s.Id == planSetup.TargetVolumeID)
                                        {
                                            evalStructure = s;
                                        }

                                    }
                                }
                            }*/
                            if (objective.TemplateId == "Tumor" && planSetup.TargetVolumeID != "")
                            {
                                var targetVolume = planSetup.StructureSet.Structures.FirstOrDefault(y => y.Id == planSetup.TargetVolumeID)?.Volume;
                                if (targetVolume == null)
                                {
                                    evalStructure = null;
                                    return;
                                }

                                //evalStructure = structureSet.Structures.Where(x => (x.DicomType == "ITV" || x.DicomType == "CTV" || x.DicomType == "GTV") & !x.IsEmpty & !x.Id.ToLower().StartsWith("z") & !x.Id.ToLower().StartsWith("h") && planSetup.StructureSet.Structures.Where(y => y.Id == planSetup.TargetVolumeID).FirstOrDefault().IsPointInsideSegment(x.CenterPoint)).FirstOrDefault();
                                // Suche zuerst nach einer ITV Struktur
                                var itvStructure = structureSet.Structures
                                    .Where(x => x.DicomType == "GTV" &&
                                                x.Id.ToLower().StartsWith("itv") &&
                                                !x.IsEmpty &&
                                                !x.Id.ToLower().StartsWith("z") &&
                                                !x.Id.ToLower().StartsWith("h") &&
                                                x.Volume < targetVolume &&
                                                planSetup.StructureSet.Structures
                                                    .Where(y => y.Id == planSetup.TargetVolumeID)
                                                    .FirstOrDefault()
                                                    .IsPointInsideSegment(x.CenterPoint))
                                    .OrderByDescending(x=>x.Volume)
                                    .FirstOrDefault();

                                // Wenn keine ITV Struktur gefunden wurde, suche nach einer CTV Struktur
                                if (itvStructure == null)
                                {
                                    itvStructure = structureSet.Structures
                                        .Where(x => x.DicomType == "CTV" &&
                                                    !x.IsEmpty &&
                                                    !x.Id.ToLower().StartsWith("z") &&
                                                    !x.Id.ToLower().StartsWith("h") &&
                                                    x.Volume < targetVolume &&
                                                    planSetup.StructureSet.Structures
                                                        .Where(y => y.Id == planSetup.TargetVolumeID)
                                                        .FirstOrDefault()
                                                        .IsPointInsideSegment(x.CenterPoint))
                                        .OrderByDescending(x => x.Volume)
                                        .FirstOrDefault();
                                }

                                // Wenn weder ITV noch CTV gefunden wurden, suche nach einer GTV Struktur
                                if (itvStructure == null)
                                {
                                    itvStructure = structureSet.Structures
                                        .Where(x => x.DicomType == "GTV" &&
                                                    !x.IsEmpty &&
                                                    !x.Id.ToLower().StartsWith("z") &&
                                                    !x.Id.ToLower().StartsWith("h") &&
                                                    x.Volume < targetVolume &&
                                                    planSetup.StructureSet.Structures
                                                        .Where(y => y.Id == planSetup.TargetVolumeID)
                                                        .FirstOrDefault()
                                                        .IsPointInsideSegment(x.CenterPoint))
                                        .OrderByDescending(x => x.Volume)
                                        .FirstOrDefault();
                                }

                                // Die gefundene Struktur wird zugewiesen oder bleibt null, wenn keine gefunden wurde
                                evalStructure = itvStructure;
                            }

                            if ((objective.TemplateId == "Target" || objective.TemplateId == "Zielvolumen") && planSetup.TargetVolumeID != "")
                            {
                                evalStructure = structureSet.Structures.Where(x => x.Id == planSetup.TargetVolumeID).FirstOrDefault();
                            }
                            if ((objective.TemplateId == "Target" || objective.TemplateId == "Zielvolumen") && planSetup.TargetVolumeID == "")
                            {
                                if (planSetup.PrimaryReferencePoint != null)
                                {
                                    if (structureSet.Structures.Where(x => x.Id.ToLower().Replace(" ", "").StartsWith(planSetup.PrimaryReferencePoint.Id.ToLower().Replace(" ", "")) & !x.IsEmpty).Any())
                                    {

                                        evalStructure = structureSet.Structures.Where(x => x.Id.ToLower().Replace(" ", "").StartsWith(planSetup.PrimaryReferencePoint.Id.ToLower().Replace(" ", ""))).FirstOrDefault();

                                    }
                                    else
                                        evalStructure = structureSet.Structures.Where(x => x.DicomType == "PTV" & !x.IsEmpty).FirstOrDefault();
                                }
                                else
                                {
                                    evalStructure = structureSet.Structures.Where(x => x.DicomType == "PTV" & !x.IsEmpty).FirstOrDefault();
                                }
                               
                            }
                            var evalStructureVM = new StructureViewModel(evalStructure);
                            try
                            {
                                var obj = calculator.GetObjectiveProperties(objective, planningItem, structureSet, evalStructureVM);
                                PqmSummaries.Add(obj);
                            }
                            catch { continue; }
                            NotifyPropertyChanged("Structure");
                        }
                    }
                    waitWindowPQM.Close();
                }
            }
        }

        public ObservableCollection<PQMSummaryViewModel> AddPQMSummary(ObservableCollection<PQMSummaryViewModel>  PqmSummaries, ConstraintViewModel constraintPath, PlanningItemViewModel planningItem, Patient patient)
        {
            StructureSet structureSet = planningItem.PlanningItemStructureSet;
            Structure evalStructure;
            //ObservableCollection<PQMSummaryViewModel> pqmSummaries = new ObservableCollection<PQMSummaryViewModel>();
            //ObservableCollection<StructureViewModel> foundStructureList = new ObservableCollection<StructureViewModel>();
            var calculator = new PQMSummaryCalculator();
            //var numCol = PqmSummaries[0]
            //Objectives = calculator.GetObjectives(constraintPath);
            if (planningItem.PlanningItemObject is PlanSum)
            {
                var waitWindowPQM = new WaitWindowPQM();
                PlanSum plansum = (PlanSum)planningItem.PlanningItemObject;
                if (plansum.IsDoseValid() == true)
                {
                    waitWindowPQM.ShowInTaskbar = false;
                    waitWindowPQM.Show();
                    foreach (PQMSummaryViewModel pqm in PqmSummaries)
                    {
                        evalStructure = calculator.FindStructureFromAlias(structureSet, planningItem, pqm.TemplateId, pqm.TemplateAliases, pqm.TemplateCodes, pqm.TemplateType);
                        if (evalStructure != null)
                        {
                            var pqmSummary = calculator.GetObjectiveProperties(pqm, planningItem, structureSet, new StructureViewModel(evalStructure));
                            //pqm.Achieved_Comparison = pqmSummary.Achieved;
                            //pqm.AchievedColor_Comparison = pqmSummary.AchievedColor;
                            //pqm.AchievedPercentageOfGoal_Comparison = pqmSummary.AchievedPercentageOfGoal;
                            //pqm.Met_Comparison = pqmSummary.Met;
                            //pqmSummaries.Add(pqmSummary);
                            //foundStructureList.Add(new StructureViewModel(evalStructure));
                        }
                    }
                    //FoundStructureList = foundStructureList;
                    waitWindowPQM.Close();
                }
                //PqmSummaries = pqmSummaries;
            }
            else //is plansetup
            {
                var waitWindowPQM = new WaitWindowPQM();

                PlanSetup planSetup = (PlanSetup)planningItem.PlanningItemObject;
                if (planSetup.IsDoseValid() == true)
                {
                    waitWindowPQM.ShowInTaskbar = false;
                    waitWindowPQM.Show();
                    foreach (PQMSummaryViewModel pqm in PqmSummaries)
                    {
                        evalStructure = calculator.FindStructureFromAlias(structureSet, planningItem, pqm.TemplateId, pqm.TemplateAliases, pqm.TemplateCodes, pqm.TemplateType);
                        if (evalStructure != null)
                        {
                            if (evalStructure.Id.Contains("PTV") == true)
                            {
                                foreach (Structure s in structureSet.Structures)
                                {
                                    if (s.Id == planSetup.TargetVolumeID)
                                        evalStructure = s;
                                }
                            }
                            var pqmSummary = calculator.GetObjectiveProperties(pqm, planningItem, structureSet, new StructureViewModel(evalStructure));
                            //pqm.Achieved_Comparison = pqmSummary.Achieved;
                            //foundStructureList.Add(new StructureViewModel(evalStructure));
                        }
                    }
                    //FoundStructureList = foundStructureList;
                    waitWindowPQM.Close();
                }
                //PqmSummaries = pqmSummaries;
            }
            return PqmSummaries;
        }

        public ObservableCollection<PlanningItemDetailsViewModel> GetPlanningItemSummary(PlanningItemViewModel activePlanningItem, ObservableCollection<PlanningItemViewModel> planningItemList)
        {
            var calculator = new PlanningItemDetailsCalculator();
            PlanningItemSummaries = calculator.Calculate(activePlanningItem, planningItemList, PqmSummaries, ErrorGrid);
            return PlanningItemSummaries;
        }

        public List<ErrorViewModel> GetErrors(PlanningItemViewModel planningItem, Patient patient)
        {
            var calculator = new ErrorCalculator();
            ErrorGrid = calculator.Calculate(planningItem.PlanningItemObject, patient);
            ErrorGrid = ErrorGrid.OrderBy(x => x.Status).ToList();
            return ErrorGrid;
        }

        public List<RefViewModel> GetRefs(PlanningItemViewModel planningItem, Patient patient)
        {
            var calculator = new ErrorCalculator();
            RefGrid = calculator.Calculate2(planningItem.PlanningItemObject, patient);
            RefGrid = RefGrid.OrderBy(x => x.RefPointId).ToList();
            return RefGrid;
        }

        

        
        private Model3DGroup CreateModel(MeshGeometry3D bodyMesh, MeshGeometry3D couchMesh, Model3DGroup isoModelGroup, Model3DGroup collimatorModelGroup, Material collimatorMaterial)
        {
            var modelGroup = new Model3DGroup();
            AddModels(bodyMesh, couchMesh, isoModelGroup, collimatorModelGroup, modelGroup, collimatorMaterial);
            return modelGroup;
        }

        private static void AddModels(MeshGeometry3D bodyMesh, MeshGeometry3D couchMesh, Model3DGroup isoModelGroup, Model3DGroup collimatorModelGroup, Model3DGroup modelGroup, Material collimatorMaterial)
        {
            // Create some materials
            var lightblueMaterial = new DiffuseMaterial(new SolidColorBrush(Colors.LightCoral));
            var darkblueMaterial = new DiffuseMaterial(new SolidColorBrush(Colors.DarkBlue));
            var magentaMaterial = new DiffuseMaterial(new SolidColorBrush(Colors.Magenta));

            modelGroup.Children.Add(isoModelGroup);
            modelGroup.Children.Add(collimatorModelGroup);
            modelGroup.Children.Add(new GeometryModel3D { Geometry = bodyMesh, Material = lightblueMaterial, BackMaterial = darkblueMaterial });
            modelGroup.Children.Add(new GeometryModel3D { Geometry = couchMesh, Material = magentaMaterial, BackMaterial = magentaMaterial });
        }
    }
}
