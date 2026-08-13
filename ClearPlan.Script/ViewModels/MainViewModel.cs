using ClearPlan.Helpers;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Review;
using ClearPlan.Calculators;
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
        public ConstraintCatalogLoadResult ConstraintCatalogStatus { get; private set; }
        public string ConstraintSourceStatus { get; private set; }
        public bool ConstraintSelectionRequiresConfirmation { get; private set; }
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
        public ObservableCollection<FieldNamePreviewViewModel> FieldNamePreviews { get; set; }
        public OverviewViewModel Overview { get; private set; }
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
            OverviewPlotModel = CreatePlotModel();
            DvhStructures = new ObservableCollection<DvhStructureViewModel>(
                (Structures ?? Enumerable.Empty<Structure>())
                    .Select(structure => new DvhStructureViewModel(structure)));
            foreach (DvhStructureViewModel item in DvhStructures)
            {
                item.PropertyChanged += DvhStructureSelectionChanged;
            }
            
            ActivePlanningItem = planningItem;
            Patient = patient;
            User = user;
            Image = ActivePlanningItem.PlanningItemImage;
            StructureSet = ActivePlanningItem.PlanningItemStructureSet;
            //int Fx = ActivePlanningItem.
            var settings = ClearPlanSettings.Load();
            ConstraintCatalogStatus = settings.LoadConstraintCatalog();
            ConstraintSourceStatus = BuildConstraintSourceStatus(ConstraintCatalogStatus);
            ConstraintComboBoxList = ConstraintListViewModel.GetConstraintList(ConstraintCatalogStatus);
            ActiveConstraintPath = SelectConstraintTable(planningItem, ConstraintCatalogStatus, ConstraintComboBoxList);
            ConstraintSelectionRequiresConfirmation =
                ActiveConstraintPath != null && ActiveConstraintPath.RequiresConfirmation;
            PlanningItemList = planningItemList;
            StructureList = StructureSetListViewModel.GetStructureList(StructureSet);
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
            FieldNamePreviews = FieldNamingPreviewCalculator.Calculate(ActivePlanningItem);
            RefreshOverview();
            //NotifyPropertyChanged("Structure");
        }

        public void RefreshOverview()
        {
            Overview = OverviewViewModel.Create(this);
            NotifyPropertyChanged("Overview");
        }

        public void ReloadConstraintCatalog()
        {
            var settings = ClearPlanSettings.Reload();
            ConstraintCatalogStatus = settings.LoadConstraintCatalog();
            ConstraintSourceStatus = BuildConstraintSourceStatus(
                ConstraintCatalogStatus);
            ConstraintComboBoxList = ConstraintListViewModel.GetConstraintList(
                ConstraintCatalogStatus);
            ActiveConstraintPath = SelectConstraintTable(
                ActivePlanningItem,
                ConstraintCatalogStatus,
                ConstraintComboBoxList);
            ConstraintSelectionRequiresConfirmation =
                ActiveConstraintPath != null &&
                ActiveConstraintPath.RequiresConfirmation;
            NotifyPropertyChanged("ConstraintCatalogStatus");
            NotifyPropertyChanged("ConstraintSourceStatus");
            NotifyPropertyChanged("ConstraintComboBoxList");
            NotifyPropertyChanged("ActiveConstraintPath");
            NotifyPropertyChanged("ConstraintSelectionRequiresConfirmation");
        }

        private static ConstraintViewModel SelectConstraintTable(
            PlanningItemViewModel planningItem,
            ConstraintCatalogLoadResult loadResult,
            ObservableCollection<ConstraintViewModel> availableTables)
        {
            if (loadResult == null || !loadResult.IsUsable || availableTables.Count == 0)
            {
                return new ConstraintViewModel(
                    null,
                    loadResult == null ? string.Empty : loadResult.ActivePath,
                    loadResult == null ? string.Empty : loadResult.ActiveSource);
            }

            ConstraintTableSelection selection = ConstraintTableSelector.Select(
                loadResult.Catalog.Tables,
                BuildConstraintContext(planningItem));
            ConstraintTableDefinition selectedTable = selection.SelectedTable ??
                                                      selection.Candidates
                                                          .Select(candidate => candidate.Table)
                                                          .FirstOrDefault();
            ConstraintViewModel selected = availableTables.FirstOrDefault(
                item => selectedTable != null &&
                        string.Equals(
                            item.ConstraintId,
                            selectedTable.TableId,
                            StringComparison.OrdinalIgnoreCase)) ??
                                           availableTables.First();
            selected.SelectionReason = selection.Reason;
            selected.RequiresConfirmation = selection.RequiresConfirmation ||
                                            selection.SelectedTable == null;
            return selected;
        }

        private static PlanConstraintContext BuildConstraintContext(
            PlanningItemViewModel planningItem)
        {
            int fractionCount;
            var context = new PlanConstraintContext
            {
                IsPlanSum = planningItem.PlanningItemObject is PlanSum,
                FractionCount = int.TryParse(planningItem.PlanningItemFx, out fractionCount)
                    ? (int?)fractionCount
                    : null,
                DosePerFractionGy = planningItem.PlanningItemDoseFx > 0
                    ? (decimal?)Convert.ToDecimal(planningItem.PlanningItemDoseFx)
                    : null,
                TotalDoseGy = planningItem.PlanningItemDose > 0
                    ? (decimal?)Convert.ToDecimal(planningItem.PlanningItemDose)
                    : null,
                StructureIds = planningItem.PlanningItemStructureSet == null
                    ? new List<string>()
                    : planningItem.PlanningItemStructureSet.Structures
                        .Select(structure => structure.Id)
                        .ToList()
            };
            return context;
        }

        private static string BuildConstraintSourceStatus(
            ConstraintCatalogLoadResult loadResult)
        {
            if (loadResult == null || !loadResult.IsUsable)
            {
                return "Constraint-Katalog nicht verfügbar";
            }

            return string.Format(
                "{0} · {1} Tabellen · {2} Constraints",
                loadResult.ActiveSource,
                loadResult.Catalog.Statistics.ActiveTables,
                loadResult.Catalog.Statistics.TotalConstraints);
        }
        #region PlanSelector
        #endregion PlanSelector
        #region dvh
        private readonly PlanSetup _plan;
        private readonly PlanSum _psum;

        
        public IEnumerable<Structure> Structures { get; private set; }

        public ObservableCollection<DvhStructureViewModel> DvhStructures { get; private set; }

        public PlotModel PlotModel { get; private set; }

        public PlotModel OverviewPlotModel { get; private set; }

        public void AddDvhCurve(Structure structure)
        {
            if (structure == null ||
                FindSeries(PlotModel, structure.Id) != null)
            {
                return;
            }

            var dvh = CalculateDvh(structure);
            PlotModel.Series.Add(CreateDvhSeries(structure.Id, dvh));
            OverviewPlotModel.Series.Add(CreateDvhSeries(structure.Id, dvh));
            UpdatePlot();
        }

        public void RemoveDvhCurve(Structure structure)
        {
            if (structure == null)
            {
                return;
            }

            RemoveSeries(PlotModel, structure.Id);
            RemoveSeries(OverviewPlotModel, structure.Id);
            UpdatePlot();
        }

        public void ApplyDefaultDvhSelections()
        {
            var requested = (PqmSummaries ??
                             new ObservableCollection<PQMSummaryViewModel>())
                .Where(item => item != null && item.Structure != null)
                .Select(item => item.Structure.StructureName)
                .Concat((RefGrid ?? new List<RefViewModel>())
                    .Where(item => item != null)
                    .Select(item => item.RefPointId))
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (DvhStructureViewModel item in DvhStructures)
            {
                item.IsSelected = DvhSelectionPolicy.ShouldSelect(
                    item.Id,
                    requested);
            }
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
            pm.LegendBorder = OxyColor.FromRgb(203, 213, 225);
            pm.LegendBackground = OxyColor.FromAColor(225, OxyColors.White);
            pm.LegendPosition = LegendPosition.RightTop;
            pm.LegendOrientation = LegendOrientation.Vertical;
            pm.LegendPlacement = LegendPlacement.Inside;
            pm.LegendMaxHeight = 240;
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
                Tag = structureId,
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

        private void DvhStructureSelectionChanged(
            object sender,
            System.ComponentModel.PropertyChangedEventArgs args)
        {
            if (!string.Equals(
                    args.PropertyName,
                    "IsSelected",
                    StringComparison.Ordinal))
            {
                return;
            }

            var item = sender as DvhStructureViewModel;
            if (item == null)
            {
                return;
            }

            if (item.IsSelected)
            {
                AddDvhCurve(item.Structure);
            }
            else
            {
                RemoveDvhCurve(item.Structure);
            }
        }

        private static Series FindSeries(
            PlotModel plotModel,
            string structureId)
        {
            return plotModel.Series.FirstOrDefault(x =>
                string.Equals(
                    x.Tag as string,
                    structureId,
                    StringComparison.Ordinal));
        }

        private static void RemoveSeries(
            PlotModel plotModel,
            string structureId)
        {
            Series series = FindSeries(plotModel, structureId);
            if (series != null)
            {
                plotModel.Series.Remove(series);
            }
        }

        private void UpdatePlot()
        {
            PlotModel.InvalidatePlot(true);
            OverviewPlotModel.InvalidatePlot(true);
        }
        #endregion dvh

        public string GetTitle(Patient patient, string scriptVersion)
        {
            Title = patient.Name + " - " + "ClearPlan v." + scriptVersion;
            return Title;
        }

        public void GetPQMSummaries(ConstraintViewModel constraintPath, PlanningItemViewModel planningItem, Patient patient)
        {
            var nextPqmSummaries =
                new ObservableCollection<PQMSummaryViewModel>();
            StructureSet structureSet = planningItem.PlanningItemStructureSet;
            //PlanSetup plaSetup = planningItem.PlanningItemId;
            Structure evalStructure;
            var calculator = new PQMSummaryCalculator();
            var nextObjectives = calculator.GetObjectives(constraintPath);
            if (planningItem.PlanningItemObject is PlanSum)
            {
                var waitWindowPQM = new WaitWindowPQM();
                PlanSum plansum = (PlanSum)planningItem.PlanningItemObject;
                if (plansum.IsDoseValid() == true)
                {
                    waitWindowPQM.ShowInTaskbar = false;
                    waitWindowPQM.Show();
                    try
                    {
                        foreach (PQMSummaryViewModel objective in nextObjectives)
                        {
                            evalStructure = calculator.FindStructureFromAlias(structureSet, planningItem, objective.TemplateId, objective.TemplateAliases, objective.TemplateCodes, objective.TemplateType);
                            if (evalStructure != null)
                            {
                                var evalStructureVM = new StructureViewModel(evalStructure);
                                var obj = calculator.GetObjectiveProperties(objective, planningItem, structureSet, evalStructureVM);
                                nextPqmSummaries.Add(obj);
                                NotifyPropertyChanged("Structure");
                            }
                        }
                    }
                    finally
                    {
                        waitWindowPQM.Close();
                    }
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
                    try
                    {
                        //waitWindowPQM.Topmost = true;
                        foreach (PQMSummaryViewModel objective in nextObjectives)
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
                                    continue;
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
                                nextPqmSummaries.Add(obj);
                            }
                            catch { continue; }
                            NotifyPropertyChanged("Structure");
                            }
                        }
                    }
                    finally
                    {
                        waitWindowPQM.Close();
                    }
                }
            }
            Objectives = nextObjectives;
            PqmSummaries = nextPqmSummaries;
            NotifyPropertyChanged("PqmSummaries");
        }

        internal PqmReviewState CapturePqmReviewState()
        {
            var dvhSelections =
                new List<KeyValuePair<DvhStructureViewModel, bool>>();
            foreach (DvhStructureViewModel item in
                     DvhStructures ??
                     new ObservableCollection<DvhStructureViewModel>())
            {
                if (item != null)
                {
                    dvhSelections.Add(
                        new KeyValuePair<DvhStructureViewModel, bool>(
                            item,
                            item.IsSelected));
                }
            }

            return new PqmReviewState
            {
                PqmSummaries = PqmSummaries,
                Objectives = Objectives,
                PlanningItemSummaries = PlanningItemSummaries,
                Overview = Overview,
                DvhSelections = dvhSelections
            };
        }

        internal void RestorePqmReviewState(PqmReviewState state)
        {
            if (state == null)
            {
                throw new ArgumentNullException("state");
            }

            Objectives = state.Objectives;
            PqmSummaries = state.PqmSummaries;
            PlanningItemSummaries = state.PlanningItemSummaries;
            Overview = state.Overview;
            NotifyPropertyChanged("PqmSummaries");
            NotifyPropertyChanged("PlanningItemSummaries");
            NotifyPropertyChanged("Overview");

            foreach (KeyValuePair<DvhStructureViewModel, bool> selection in
                     state.DvhSelections ??
                     new List<KeyValuePair<DvhStructureViewModel, bool>>())
            {
                if (selection.Key != null)
                {
                    selection.Key.IsSelected = selection.Value;
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

    internal sealed class PqmReviewState
    {
        public ObservableCollection<PQMSummaryViewModel>
            PqmSummaries { get; set; }

        public PQMSummaryViewModel[] Objectives { get; set; }

        public ObservableCollection<PlanningItemDetailsViewModel>
            PlanningItemSummaries { get; set; }

        public OverviewViewModel Overview { get; set; }

        public IList<KeyValuePair<DvhStructureViewModel, bool>>
            DvhSelections { get; set; }
    }
}
