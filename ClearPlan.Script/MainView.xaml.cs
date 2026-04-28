using ClearPlan.Helpers;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Media3D;
using VMS.TPS.Common.Model.API;
using System.Collections.Generic;
using System.Windows.Media;
using Microsoft.Win32;
using OxyPlot.Wpf;
using System.Windows.Media.Imaging;
using System.Text;
using Path = System.IO.Path;
using VMS.TPS.Common.Model.Types;
using Beam = VMS.TPS.Common.Model.API.Beam;
using System.Drawing.Imaging;
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp;
using PdfSharp.Drawing.Layout;
using NLog;
using System.Reflection;

namespace ClearPlan
{
    /// <summary>
    /// Interaction logic for MainView.xaml
    /// </summary>
    
    public partial class MainView : UserControl
    {
        private static readonly Logger log = Log.GetLogger();
        private MainViewModel _vm;
        private readonly ClearPlanSettings _settings = ClearPlanSettings.Load();
        //private readonly DVHViewModel _vm1;
        // Create dummy PlotView to force OxyPlot.Wpf to be loaded
        private static readonly PlotView PlotView = new PlotView();
        //private readonly DVHViewModel _vm1;
        public List<string> SelectedDVHs = new List<string>();
        public List<string> CurrentpqmNameList = new List<string>();
        public string CurrentpqmNames;
        public BitmapSource dvhBitmapSource;
        public Bitmap dvhBitmap;
        public int w = 1;

        public MainView(MainViewModel mainViewModel)
        {
            
            _vm = mainViewModel;
            
            InitializeComponent();
            
            DataContext = _vm;
            
            _vm.GetPQMSummaries(_vm.ActiveConstraintPath, _vm.ActivePlanningItem, _vm.Patient);
            UpdatePqmDataGrid();

            
            var performSort = typeof(DataGrid).GetMethod("PerformSort", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[2] });
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[2] });
            

        }

        private string BuildReportFilePath(string categoryFolder, string fileName)
        {
            string fullDirectory = Path.Combine(_settings.GetReportsDirectory(), categoryFolder);
            Directory.CreateDirectory(fullDirectory);
            return Path.Combine(fullDirectory, fileName);
        }

        private string BuildLogsFilePath(params string[] pathParts)
        {
            string currentPath = _settings.GetLogsDirectory();
            foreach (string pathPart in pathParts.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                currentPath = Path.Combine(currentPath, pathPart);
            }

            string directory = Path.GetDirectoryName(currentPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            return currentPath;
        }

        private void ShowDocumentationWindow(string title, string content)
        {
            Window owner = Window.GetWindow(this);
            var viewer = new TextBox
            {
                Text = content,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(12),
                FontFamily = new System.Windows.Media.FontFamily("Consolas")
            };

            var window = new Window
            {
                Title = title,
                Width = 760,
                Height = 560,
                Content = viewer,
                WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen,
                Owner = owner
            };

            window.ShowDialog();
        }

        private void OpenDocumentation(string title, string filePath, string fallbackText, string preferredSection = null)
        {
            string content = string.IsNullOrWhiteSpace(preferredSection)
                ? DocumentationHelper.ReadTextOrDefault(filePath, fallbackText)
                : DocumentationHelper.GetLatestSection(filePath, preferredSection);

            if (string.IsNullOrWhiteSpace(content))
            {
                content = fallbackText;
            }

            ShowDocumentationWindow(title, content);
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj)
        where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

        public static childItem FindVisualChild<childItem>(DependencyObject obj)
            where childItem : DependencyObject
        {
            foreach (childItem child in FindVisualChildren<childItem>(obj))
            {
                return child;
            }

            return null;
        }

        public void UpdatePqmDataGrid()
        {
            
            pqmDataGrid.ItemsSource = null;
            pqmDataGrid.ItemsSource = _vm.PqmSummaries;
            _vm.PlanningItemSummaries = _vm.GetPlanningItemSummary(_vm.ActivePlanningItem, _vm.PlanningItemList);
            planningItemSummariesDataGrid.ItemsSource = _vm.PlanningItemSummaries;

            RefreshDVHClicked(null, null);
            var performSort = typeof(DataGrid).GetMethod("PerformSort", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[7] });
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[7] });

        }

        public static void FindChildGroup<T>(DependencyObject parent, string childName, ref List<T> list) where T : DependencyObject
        {
            
            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);

            for (int i = 0; i < childrenCount; i++)
            {
                // Get the child
                var child = VisualTreeHelper.GetChild(parent, i);

                // Compare on conformity the type
                T child_Test = child as T;

                // Not compare - go next
                if (child_Test == null)
                {
                    // Go the deep
                    FindChildGroup<T>(child, childName, ref list);
                }
                else
                {
                    // If match, then check the name of the item
                    FrameworkElement child_Element = child_Test as FrameworkElement;

                    if (child_Element.Name == childName)
                    {
                        // Found
                        list.Add(child_Test);
                    }

                    // We are looking for further, perhaps there are
                    // children with the same name
                    FindChildGroup<T>(child, childName, ref list);
                }
            }

            return;
        }

        private void ConstraintComboBoxChanged(object sender, SelectionChangedEventArgs e)
        {
            var selection = (ConstraintViewModel)ConstraintComboBox.SelectedItem;
            if (_vm.ActiveConstraintPath.ConstraintPath != selection.ConstraintPath)
            {
                _vm.ActiveConstraintPath = (ConstraintViewModel)ConstraintComboBox.SelectedItem;
                var calculator = new PQMSummaryCalculator();
                _vm.GetPQMSummaries(_vm.ActiveConstraintPath, _vm.ActivePlanningItem, _vm.Patient);
                
                planningItemSummariesDataGrid.ItemsSource = null;
                planningItemSummariesDataGrid.ItemsSource = _vm.PlanningItemSummaries;

                UpdatePqmDataGrid();
                RefreshDVHClicked(null, null);
                LogLiveMining();
            }
        }

        ///dvh stuff
        private void Structure_OnChecked(object checkBoxObject, RoutedEventArgs e)
        {            
            _vm.AddDvhCurve(GetStructure1(checkBoxObject));
        }

        private void Structure_OnUnchecked(object checkBoxObject, RoutedEventArgs e)
        {
            _vm.RemoveDvhCurve(GetStructure1(checkBoxObject));
        }

        private Structure GetStructure1(object checkBoxObject)
        {
            var checkbox = (CheckBox)checkBoxObject;
            var structure = (Structure)checkbox.DataContext;
            
            return structure;
        }

        private void ExportPlotAsPdf(object sender, RoutedEventArgs e)
        {
            var filePath = GetPdfSavePath();
            if (filePath != null)
                _vm.ExportPlotAsPdf(filePath);
        }

        private string GetPdfSavePath()
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = "Export to PDF",
                Filter = "PDF Files (*.pdf)|*.pdf"
            };

            var dialogResult = saveFileDialog.ShowDialog();

            if (dialogResult == true)
                return saveFileDialog.FileName;
            else
                return null;
        }

        ////////////
        private PlanningItemViewModel GetPlan(object sender)
        {
            var selection = (Button)sender;
            var planningItem = (PlanningItemDetailsViewModel)selection.DataContext;
            return new PlanningItemViewModel(planningItem.PlanningItemObject);
        }


        private void StructureChanged(object sender, EventArgs e)
        {       
            // Code for auto-select all comboBoxes with the same tempkateID
            ComboBox comboBox = (ComboBox)sender;            

            string selectedTemplateID = (pqmDataGrid.SelectedCells[0].Column.GetCellContent(pqmDataGrid.SelectedCells[0].Item) as TextBlock).Text;            

            StructureViewModel cb = (StructureViewModel)comboBox.SelectedItem;            
            int selectedStructureIndex = comboBox.SelectedIndex;            

            for (int i = 0; i < pqmDataGrid.Items.Count; i++)
            {
                DataGridRow row = (DataGridRow)pqmDataGrid.ItemContainerGenerator.ContainerFromIndex(i);                
                if ((pqmDataGrid.Columns[0].GetCellContent(row) as TextBlock).Text == selectedTemplateID)
                { 
                    // here is the tricky part because the comboBox is nested
                    ComboBox cbi = FindChild<ComboBox>(pqmDataGrid.Columns[1].GetCellContent(row), "StructureComboBox");
                    cbi.SelectedIndex = selectedStructureIndex;                    
                }                
            }

            // normal update
            UpdatePqmDataGrid();
        }

        
        
        
        private void RefreshDVHClicked(object sender, RoutedEventArgs e)
        {
            
            try
            {
                CurrentpqmNameList.Clear();
                foreach (var pqm in _vm.PqmSummaries)
                {
                    CurrentpqmNameList.Add(pqm.Structure.StructureNameWithCode);
                }
                foreach (var _ref in _vm.RefGrid)
                {
                    CurrentpqmNameList.Add(_ref.RefPointId);
                }
                CurrentpqmNames = string.Join(",", CurrentpqmNameList);

                for (int i = 0; i < DVHstrucListBox.Items.Count; i++)
                {
                    var item = DVHstrucListBox.ItemContainerGenerator.ContainerFromItem(DVHstrucListBox.Items[i]) as ListBoxItem;
                    var template = item.ContentTemplate as DataTemplate;

                    ContentPresenter myContentPresenter = FindVisualChild<ContentPresenter>(item);
                    string itemContent = item.Content.ToString().Contains(":") ? item.Content.ToString().Split(':')[0] : item.Content.ToString();
                    CheckBox myCheckBox = (CheckBox)template.FindName("DVHstrucCBs", myContentPresenter);
                    if (CurrentpqmNames.Contains(itemContent) && itemContent.ToUpper() != "BODY" 
                        && itemContent.ToUpper() != "OUTER CONTOUR" && itemContent.ToUpper() != "EXTERNAL" 
                        && itemContent.ToUpper() != "KÖRPER" && itemContent.ToUpper() != "KÖRPER"
                        &! ((itemContent.Contains("TL") || itemContent.Contains("TB") || itemContent.Contains("TR") || itemContent.Contains("TL") || itemContent.ToLower().Contains("teil"))
                        & !(itemContent.ToLower().Contains("ptv") || itemContent.ToLower().Contains("ctv") || itemContent.ToLower().Contains("gtv") || itemContent.ToLower().Contains("itv")))
                        )
                        myCheckBox.IsChecked = true;
                    else
                    {
                        myCheckBox.IsChecked = false;
                    }
                }
            }
            catch { }
            
            

        }

        public static BitmapSource ConvertBitmap(Bitmap source)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                          source.GetHbitmap(),
                          IntPtr.Zero,
                          Int32Rect.Empty,
                          BitmapSizeOptions.FromEmptyOptions());
        }

        public static Bitmap BitmapFromSource(BitmapSource bitmapsource)
        {
            Bitmap bitmap;
            using (var outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapsource));
                enc.Save(outStream);
                
                bitmap = new Bitmap(outStream);
                
            }
            //using(var gr = Graphics.)
            return bitmap;
        }

        private void LogLiveMining()
        {
            var liveMiningLogger = new CustomLog("LiveMining.csv");

            try
            {
                StringBuilder csvContent = new StringBuilder();
                string logsDir = _settings.GetLogsDirectory();
                string logFilePath = Path.Combine(logsDir, "LiveMining.csv");

                // Sicherstellen, dass Logs-Ordner existiert
                if (!Directory.Exists(logsDir))
                {
                    Directory.CreateDirectory(logsDir);
                }

                bool fileExists = File.Exists(logFilePath);

                // Header nur schreiben, falls Datei noch nicht existiert
                if (!fileExists)
                {
                    List<string> headers = new List<string>
            {
                "User", "PC", "Script", "DateNow", "TimeNow", "Patient", "PatientId",
                "Course", "PlanId", "PlanUID", "PlanApprover", "PlanApproveDate",
                "TreatApprover", "TreatApproveDate", "TotalDose", "FxDose",
                "Constraint-Table", "PQMs-Fail", "PQMs-Total", "Error-ID", "Status", "Description"
            };
                    csvContent.AppendLine(string.Join(";", headers));
                }

                int miningPqmsFail = _vm.PqmSummaries.Count(x => x.Met == "Not met");
                int miningPqmsTotal = _vm.PqmSummaries.Count();

                string userId = Environment.UserName.ToString();
                string pc = Environment.MachineName;
                string scriptVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
                string scriptId = "ClearPlan - " + scriptVersion;
                string date = DateTime.Now.ToString("yyyy-MM-dd");
                string time = DateTime.Now.ToString("HH:mm");

                foreach (var check in _vm.ErrorGrid.Where(x => !x.Status.Contains("OK")))
                {
                    List<string> row = new List<string>
            {
                userId, pc, scriptId, date, time,
                _vm.Patient.Name.Replace(";", "^"),
                _vm.Patient.Id.Replace(";", "^"),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemCourse),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemId),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemUID) == ""
                    ? ("Sum - " + _vm.ActivePlanningItem.PlanningItemSumApproved.ToString())
                    : MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemUID),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningApprover ?? "noPlanningApr").Replace("_ Planner_", ""),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemApproveDate ?? "noPlanAprDate"),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName),
                MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemTreatmentApproveDate ?? "noTreatAprDate"),
                Math.Round(_vm.ActivePlanningItem.PlanningItemDose, 2).ToString(),
                Math.Round(_vm.ActivePlanningItem.PlanningItemDoseFx, 2).ToString(),
                MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName),
                miningPqmsFail.ToString(),
                miningPqmsTotal.ToString(),
                MakeFilenameValid(check.Severity).Replace(";", "."),
                MakeFilenameValid(check.Status).Replace(";", "."),
                MakeFilenameValid(check.Description).Replace(";", ".")
            };

                    csvContent.AppendLine(string.Join(";", row));
                }

                
                    csvContent.AppendLine("end");
                    File.AppendAllText(logFilePath, csvContent.ToString());
                

                //liveMiningLogger.Info("Live mining data logged.");
            }
            catch (Exception ex)
            {
                liveMiningLogger.Error(ex, "Error during live data mining.");
            }
        }

        
        private void PrintButtonClicked(object sender, RoutedEventArgs e)
        {

            for (int i = 0; i < DVHstrucListBox.Items.Count; i++)
            {
                var item = DVHstrucListBox.ItemContainerGenerator.ContainerFromItem(DVHstrucListBox.Items[i]) as ListBoxItem;
                var template = item.ContentTemplate as DataTemplate;                
                ContentPresenter myContentPresenter = FindVisualChild<ContentPresenter>(item);
                string itemContent = item.Content.ToString().Contains(":") ? item.Content.ToString().Split(':')[0] : item.Content.ToString();
                CheckBox myCheckBox = (CheckBox)template.FindName("DVHstrucCBs", myContentPresenter);
                if (myCheckBox.IsChecked == true)
                {
                    if (!SelectedDVHs.Contains(item.Content.ToString().Contains(":") ? item.Content.ToString().Split(':')[0] : item.Content.ToString()))
                    {
                        SelectedDVHs.Add(item.Content.ToString().Contains(":") ? item.Content.ToString().Split(':')[0] : item.Content.ToString());
                    }
                }
            }
            
            //if (noDVHlines == true)
                //RefreshDVHClicked(null, null);
            int scale = 4;
            var pngExporter = new PngExporter {Width=scale*650, Height=scale*330, Resolution = scale*96 };
            dvhBitmapSource = pngExporter.ExportToBitmap(_vm.PlotModel);
            dvhBitmap = BitmapFromSource(dvhBitmapSource);

            //CalculateDVHstat(_vm.ActivePlanningItem.PlanningItemObject, _vm.Patient);

            Window window = Window.GetWindow(this);
            window.Topmost = false;

            List<CheckBox> checkBoxlist = new List<CheckBox>();
            List<CheckBox> checkBoxlist2 = new List<CheckBox>();

            // Find all elements
            FindChildGroup<CheckBox>(pqmDataGrid, "checkboxinstance", ref checkBoxlist);
            FindChildGroup<CheckBox>(pqmDataGrid, "checkboxinstance2", ref checkBoxlist2);
            //int w = 0;
            int j = 0;
            int numPqms = _vm.PqmSummaries.Count();
            string msg = "";
            ReportPQM[] reportPQMList = new ReportPQM[numPqms];

            int f = 0;
            foreach (var PlanningItemObject in _vm.PlanningItemSummaries)
            {
                if (_vm.ActivePlanningItem.PlanningItemIdWithCourse == _vm.PlanningItemSummaries[f].PlanningItemIdWithCourse)
                {                 
                    break;
                }
                f++;
            }
            
            string pathReports = _settings.GetReportsDirectory();
            string constraintReportsPath = Path.Combine(pathReports, "ConstraintReports");

            var path = "";
            var reportService = new ReportPdf();
            var reportData = CreateReportData();
            string patientId = _vm.Patient.Id;
            string courseId = _vm.ActivePlanningItem.PlanningItemCourse;
            string LastName = _vm.Patient.LastName;
            string planId = _vm.ActivePlanningItem.PlanningItemId.Replace(":", "_");
            var path2 = Path.Combine(pathReports, "DRR", MakeFilenameValid(patientId), MakeFilenameValid(courseId), MakeFilenameValid(planId)) + Path.DirectorySeparatorChar;
            
            string printUser = "print-" + _vm.User.Id.Replace(@"oncology\", "").Replace(@"\", "") +"_";
            printUser = "";
            bool skipReport = false;

            log.Info(string.Format("reportPath: '{0}' ", constraintReportsPath));
            if (_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName == "" && _vm.ActivePlanningItem.PlanningItemType == "Plan")
            {

                string msg2 = string.Format("The associated plan '{0}' is not approved for treatment.\nThe report will be generated, but should not be used for official patient documentation.", _vm.ActivePlanningItem.PlanningItemId);
                MessageBox.Show(msg2, "Attention", MessageBoxButton.OK, MessageBoxImage.Information);
                //return;
                
                path = BuildReportFilePath(
                    Path.Combine("ConstraintReports", "NotTreatmentApproved"),
                    "NotTreatmentApproved_" + printUser + patientId + "_" + LastName + "_" + MakeFilenameValid(planId.Replace(":", "_")) + ".pdf");
            }
            else if (_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName != "" && _vm.ActivePlanningItem.PlanningItemType == "Plan")
            {
                foreach (var pqm in _vm.PqmSummaries)
                {
                    if (checkBoxlist[j].IsChecked == false && pqm.Met.ToString().ToUpper() == "NOT MET" && pqm.Ignore.ToString() != "True")
                    {

                        msg += string.Format("Objective '{0}' for {1} is not met.\n", pqm.DVHObjective.ToString(), pqm.Structure.StructureNameWithCode.ToString());
                        checkBoxlist[j].IsChecked = true;
                    }
                    if (pqm.Met.ToString().ToUpper() == "NOT EVALUATED" && pqm.Ignore.ToString() != "True")
                    {
                        checkBoxlist2[j].IsChecked = true;
                    }
                    j++;
                }
                if (msg != "")
                {
                    MessageBox.Show(msg +" \n" + "'Not met'-objectives will be checked for you. You will accept this by pressing the Print-Button again.", "PDF-Generation failed", MessageBoxButton.OK, MessageBoxImage.Information);
                    skipReport = true;
                    return;
                }
                else
                {
                    
                    path = BuildReportFilePath(
                        "ConstraintReports",
                        "TreatmentApproved_" + printUser + patientId + "_" + LastName + "_" + MakeFilenameValid(planId.Replace(":", "_")) + ".pdf");
                    
                }
            }
            else if (_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName != "" &! _vm.PlanningItemSummaries[f].TreatmentApprover.ToString().ToUpper().Contains("NOT APPROVED") && _vm.ActivePlanningItem.PlanningItemType == "PlanSum")
            {
                foreach (var pqm in _vm.PqmSummaries)
                {
                    if (checkBoxlist[j].IsChecked == false && pqm.Met.ToString().ToUpper() == "NOT MET" && pqm.Ignore.ToString()!="True")
                    {

                        msg += string.Format("Objective '{0}' for {1} is not met.\n", pqm.DVHObjective.ToString(), pqm.Structure.StructureNameWithCode.ToString());
                        checkBoxlist[j].IsChecked = true;
                    }
                    if (pqm.Met.ToString().ToUpper() == "NOT EVALUATED" && pqm.Ignore.ToString() != "True")
                    {
                        checkBoxlist2[j].IsChecked = true;
                    }
                    j++;
                }
                if (msg != "")
                {
                    MessageBox.Show(msg + " \n" + "'Not met'-objectives will be checked for you. You will accept this by pressing the Print-Button again.", "PDF-Generation failed", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                else
                {
                    
                    path = BuildReportFilePath(
                        "ConstraintReports",
                        "PlanSum_" + printUser + patientId + "_" + LastName + "_" + MakeFilenameValid(planId.Replace(":", "_")) + ".pdf");
                }
            }
            else if (_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName != "" && _vm.PlanningItemSummaries[f].TreatmentApprover.ToString().ToUpper().Contains("NOT APPROVED") && _vm.ActivePlanningItem.PlanningItemType == "PlanSum")
            {
                string msg2 = string.Format("The associated PlanSum '{0}' contains plans that are not treatment approved.\nThe report will be generated, but should not be used for official patient documentation.", _vm.ActivePlanningItem.PlanningItemId);
                MessageBox.Show(msg2, "Attention", MessageBoxButton.OK, MessageBoxImage.Information);

                    path = BuildReportFilePath(
                        Path.Combine("ConstraintReports", "NotTreatmentApproved"),
                        "PlanSum_" + printUser + patientId + "_" + LastName + "_" + MakeFilenameValid(planId.Replace(":", "_")) + ".pdf");
                }
            else
            {

                path = BuildReportFilePath(
                    Path.Combine("ConstraintReports", "Error"),
                    "ERROR_" + printUser + patientId + "_" + LastName + "_" + MakeFilenameValid(planId.Replace(":", "_")) + ".pdf");
            }

            if (skipReport!=true)
            {
                reportService.Export(path, reportData);

                if (DRR_CB.IsChecked == true)
                {
                    try
                    {
                        #region creating DRR
                        List<Beam> beams = _vm.ActivePlanningItem.PlanningItemBeams;


                        if (!Directory.Exists(path2))
                            Directory.CreateDirectory(path2);
                        PlanSetup plan = _vm.ActivePlanningItem.PlanningItemPlanSetup;

                        foreach (Beam beam in beams.Where(x => x.ReferenceImage != null && (x.MLCPlanType.ToString() == "NotDefined" || x.MLCPlanType.ToString() == "Static")).OrderBy(x=>x.Id))
                        {
                            VMS.TPS.Common.Model.API.Image Drr = beam.ReferenceImage;
                            double maxX1 = 0;
                            double maxY1 = 0;
                            double maxX2 = 0;
                            double maxY2 = 0;
                            double maxX = 0;
                            double maxY = 0;
                            foreach (ControlPoint cp in beam.ControlPoints)
                            {
                                if (Math.Abs(cp.JawPositions.X1) / 10 > Math.Abs(maxX1))
                                {
                                    //jaw size returned in mm. div by 10 for cm.
                                    maxX1 = cp.JawPositions.X1 / 10;
                                }
                                if (Math.Abs(cp.JawPositions.X2) / 10 > Math.Abs(maxX2))
                                {
                                    //jaw size returned in mm. div by 10 for cm.
                                    maxX2 = cp.JawPositions.X2 / 10;
                                }
                                if (Math.Abs(cp.JawPositions.Y1) / 10 > Math.Abs(maxY1))
                                {
                                    //jaw size returned in mm. div by 10 for cm.
                                    maxY1 = cp.JawPositions.Y1 / 10;
                                }
                                if (Math.Abs(cp.JawPositions.Y2) / 10 > Math.Abs(maxY2))
                                {
                                    //jaw size returned in mm. div by 10 for cm.
                                    maxY2 = cp.JawPositions.Y2 / 10;
                                }
                            }
                            maxX = Math.Abs(maxX2 - maxX1);
                            maxY = Math.Abs(maxY2 - maxY1);
                            WriteableBitmap drr = BuildDRRImage(beam, Drr,maxX,maxY);
                            SourcetoPng(drr, path2);
                            System.Drawing.Image png = System.Drawing.Image.FromFile(path2 + "Drr.png");
                            Bitmap drr_fieldLines = FieldLines(beam, png, Drr, plan, _vm.Patient, (Anonymize_CheckBox.IsChecked == true ? NewID_TextBox.Text : _vm.Patient.Id));
                            png.Dispose();
                            string filenameDRR = MakeFilenameValid("Drr_" + (Anonymize_CheckBox.IsChecked == true ? NewID_TextBox.Text : _vm.Patient.Id) + "_" + plan.Id + "_" + beam.Id + ".png");

                            drr_fieldLines.Save(path2 + filenameDRR, ImageFormat.Png);
                        }
                        try
                        { File.Delete(path2 + "Drr.png"); }
                        catch { }
                        #endregion creating DRR
                        #region merge DRR with PQM
                        // Path to the input PDF file
                        string inputFile = path;

                        // Path to the folder containing the PNG files to be added
                        string pngFolder = path2;

                        // Load the input PDF file
                        // Load existing PDF document
                        PdfDocument inputDocument = PdfReader.Open(path, PdfDocumentOpenMode.Modify);
                        int checkPages = inputDocument.PageCount;
                        // Loop through PNG files in a folder and add each one as a new page to the PDF
                        string[] pngFiles = Directory.GetFiles(path2, "*.png");
                        double scale2 = 1;
                        int pngFileIndex = 0;
                        foreach (string pngFile in pngFiles)
                        {
                            Beam PNGbeam = beams.Where(x => x.ReferenceImage != null && (x.MLCPlanType.ToString() == "NotDefined" || x.MLCPlanType.ToString() == "Static")).OrderBy(x => x.Id).ElementAtOrDefault(pngFileIndex);
                            pngFileIndex++;
                            // Create new PDF page with size of PNG image
                            using (Bitmap bitmap = new Bitmap(pngFile))
                            {
                                PdfPage page = inputDocument.AddPage();
                                XGraphics gfx = XGraphics.FromPdfPage(page);
                                using (XImage image = XImage.FromFile(pngFile))
                                {
                                    // Calculate the scale factor to fit the image to the page
                                    scale2 = Math.Min(page.Width / image.PixelWidth, page.Height / image.PixelHeight);

                                    // Calculate the new image size with the aspect ratio preserved
                                    double newWidth = image.PixelWidth * scale2;
                                    double newHeight = image.PixelHeight * scale2;

                                    // Calculate the offset to center the image on the page with a little bit of space from the top and left
                                    double offsetX = (page.Width - newWidth) / 2 + page.Width / 16;
                                    double offsetY = (page.Height - newHeight) / 2 + page.Height / 24;

                                    // Draw the image with the calculated parameters
                                    gfx.DrawImage(image, offsetX, offsetY, newWidth, newHeight);

                                    #region print fieldInfo table
                                    // Create Table
                                    XStringFormat format = new XStringFormat();
                                    format.LineAlignment = XLineAlignment.Near;
                                    format.Alignment = XStringAlignment.Near;
                                    var tf = new XTextFormatter(gfx);

                                    XFont fontParagraph = new XFont("Verdana", 8, XFontStyle.Regular);
                                    XFont fontParagraph2 = new XFont("Verdana", 16, XFontStyle.Bold);

                                    // Row elements
                                    double cell_width = newWidth/9.5;

                                    // page structure options
                                    double lineHeight = 20;
                                    double marginLeft = offsetX;
                                    double marginTop = offsetY/1.5;

                                    int el_height = 30;
                                    int rect_height = 17;

                                    int interLine_X_1 = 4;
                                    int interLine_Y_1 = 3;

                                    double offSetX_1 = cell_width;

                                    XSolidBrush rect_style1 = new XSolidBrush(XColors.Lavender);
                                    XSolidBrush rect_style2 = new XSolidBrush(XColor.FromArgb(0, 81, 154)); 

                                    double dist_Y = lineHeight;
                                    double dist_Y2 = dist_Y - 2;

                                    //Title
                                    tf.DrawString("Beam's-Eye-View", fontParagraph2, XBrushes.Black,
                                                    new XRect(marginLeft, offsetY / 4, cell_width * 5, el_height), format);
                                    tf.DrawString(_vm.Patient.Name, fontParagraph, XBrushes.Black,
                                            new XRect(marginLeft, offsetY / 2.66, cell_width * 10, el_height), format);
                                    //print stemp
                                    DateTime newDateTime = DateTime.Now - new TimeSpan(0,0,7);
                                    string datetime = string.Format(" {0}-{1}-{2} {3}:{4}", DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"), newDateTime.ToString("HH"), newDateTime.ToString("mm"));
                                    //string time = string.Format("{0}:{1}", DateTime.Now.ToString("HH"), DateTime.Now.ToString("mm"));
                                    tf.DrawString(string.Format("Print from {1}:{0}",datetime, _vm.User.Name.Replace(@"oncology\", "")), fontParagraph, XBrushes.Black,
                                                    new XRect(marginLeft, offsetY + newHeight + interLine_Y_1, cell_width*5, el_height), format);
                                    
                                    //Columns
                                    gfx.DrawRectangle(rect_style2, marginLeft, marginTop, newWidth, rect_height);

                                    tf.DrawString("Feld-ID", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);

                                    tf.DrawString("Energie", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);

                                    tf.DrawString("Maschine", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*2 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);
                                    tf.DrawString("Gantry", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*3 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);
                                    tf.DrawString("Collimator", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*4 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);
                                    tf.DrawString("Couch", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*5 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);
                                    tf.DrawString("SSD[cm]", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*6 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);
                                    tf.DrawString("Keil", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*7 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);
                                    tf.DrawString("MU", fontParagraph, XBrushes.White,
                                                    new XRect(marginLeft + offSetX_1*8 + interLine_X_1, marginTop+ interLine_Y_1, cell_width, el_height), format);

                                    // Row
                                    gfx.DrawRectangle(rect_style1, marginLeft, dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(PNGbeam.Id, fontParagraph, XBrushes.Black,
                                                    new XRect(marginLeft + interLine_X_1, dist_Y + marginTop, cell_width, el_height), format);

                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        PNGbeam.EnergyModeDisplayName,
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1 + interLine_X_1, dist_Y + marginTop, cell_width, el_height),
                                        format);

                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*2 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        PNGbeam.TreatmentUnit.Id,
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*2 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*3 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        Math.Round(PNGbeam.ControlPoints.First().GantryAngle,1).ToString(),
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*3 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*4 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        Math.Round(PNGbeam.ControlPoints.First().CollimatorAngle,1).ToString(),
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*4 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*5 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        Math.Round(PNGbeam.ControlPoints.First().PatientSupportAngle,1).ToString(),
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*5 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*6 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        Math.Round(PNGbeam.SSD / 10.0, 2).ToString(),
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*6 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*7 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        PNGbeam.Wedges.FirstOrDefault()==null?"": PNGbeam.Wedges.FirstOrDefault().Id,
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*7 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    gfx.DrawRectangle(rect_style1, marginLeft + offSetX_1*8 , dist_Y2 + marginTop, cell_width-2, rect_height);
                                    tf.DrawString(
                                        Math.Round(PNGbeam.Meterset.Value, 2).ToString(),
                                        fontParagraph,
                                        XBrushes.Black,
                                        new XRect(marginLeft + offSetX_1*8 + interLine_X_1, dist_Y + marginTop, cell_width-2, el_height),
                                        format);
                                    #endregion print fieldInfo table                                     

                                }
                            }
                        }

                        // Save modified PDF document
                        inputDocument.Save(path);
                        #endregion merge DRR with PQM

                        #region rotate and combine two pages
                        // Create the output document
                        PdfDocument outputDocument = new PdfDocument();

                        // Show single pages
                        // (Note: one page contains two pages from the source document)
                        outputDocument.PageLayout = PdfPageLayout.SinglePage;

                        XFont font = new XFont("Verdana", 8, XFontStyle.Bold);
                        XStringFormat format2 = new XStringFormat();
                        format2.Alignment = XStringAlignment.Center;
                        format2.LineAlignment = XLineAlignment.Far;
                        XGraphics gfx2;
                        XRect box;

                        // Open the external document as XPdfForm object
                        XPdfForm form = XPdfForm.FromFile(path);

                        for (int idx = 0; idx < form.PageCount; idx += 1)
                        {
                            if (idx < checkPages)
                            {
                                // Add a new page to the output document
                                PdfPage page = outputDocument.AddPage();
                                //page.Orientation = PageOrientation.Landscape;
                                double width = page.Width;
                                double height = page.Height;

                                //int rotate = page.Elements.GetInteger("/Rotate");

                                gfx2 = XGraphics.FromPdfPage(page);

                                // Set page number (which is one-based)
                                form.PageNumber = idx + 1;

                                box = new XRect(0, 0, width, height);
                                // Draw the page identified by the page number like an image
                                gfx2.DrawImage(form, box);
                            }
                            
                            else
                            {
                                // Add a new page to the output document
                                PdfPage page = outputDocument.AddPage();
                                page.Orientation = PageOrientation.Landscape;
                                double width = page.Width;
                                double height = page.Height;

                                int rotate = page.Elements.GetInteger("/Rotate");

                                gfx2 = XGraphics.FromPdfPage(page);

                                // Set page number (which is one-based)
                                form.PageNumber = idx + 1;

                                box = new XRect(0, 0, width / 2, height);
                                // Draw the page identified by the page number like an image
                                gfx2.DrawImage(form, box);


                                if (idx + 1 < form.PageCount)
                                {
                                    // Set page number (which is one-based)
                                    form.PageNumber = idx + 2;

                                    box = new XRect(width / 2, 0, width / 2, height);
                                    // Draw the page identified by the page number like an image
                                    gfx2.DrawImage(form, box);

                                }
                                idx += 1;
                            }
                        }

                        // Save the document...
                        outputDocument.Save(path);
                        #endregion rotate and combine two pages

                        Directory.Delete(path2, true);

                    }
                    catch { }
                }

                Process.Start(path);

                if (eDoc_CB.IsChecked == true && _vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName != "" && eDoc_CB.IsEnabled)
                {
                    //Log.Initialize(_vm.User, _vm.Patient, System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());

                   
                    //DocumentType _documentType = new DocumentType
                    //{
                    //    DocumentTypeDescription = "PQM Report"
                    //};

                    //string user = @"oncology\" + _vm.User.Id;

                    //System.Net.ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;
                    //string docKey = "xxx";
                    //var documentPushRequest = new CustomInsertDocumentsParameter
                    //{
                    //    PatientId = new PatientIdentifier { ID1 = _vm.Patient.Id },
                    //    DateOfService = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/",
                    //    DateEntered = $"/Date({Math.Floor((DateTime.Now - new DateTime(1970, 1, 1)).TotalMilliseconds)})/",
                    //    //BinaryContent = Convert.ToBase64String(binaryContent),
                    //    BinaryContent = Convert.ToBase64String(File.ReadAllBytes(path)),
                    //    FileFormat = FileFormat.PDF,
                    //    AuthoredByUser = new DocumentUser
                    //    {
                    //        SingleUserId = user
                    //    },
                    //    SupervisedByUser = new DocumentUser
                    //    {
                    //        SingleUserId = user
                    //    },
                    //    EnteredByUser = new DocumentUser
                    //    {
                    //        SingleUserId = user
                    //    },
                    //    TemplateName = $"PQM_{_vm.ActivePlanningItem.PlanningItemIdWithCourse}",
                    //    DocumentType = _documentType
                    //};
                    //var request_base = "{\"__type\":\"";
                    //var request_document = $"{request_base}InsertDocumentRequest:http://services.varian.com/Patient/Documents\",{JsonConvert.SerializeObject(documentPushRequest).TrimStart('{')}}}";
                    //string response_document = SendData(request_document, true, docKey, hostNameFromBox);

                    //if (!response_document.Contains("GatewayError"))
                    //{
                    //    //VMS.OIS.ARIAExternal.WebServices.Documents.Contracts.DocumentResponse documentResponse = JsonConvert.DeserializeObject<VMS.OIS.ARIAExternal.WebServices.Documents.Contracts.DocumentResponse>(response_document);
                    //    log.Info(string.Format("Upload for '{0}' was succesfull.", _vm.Patient.Id+"_"+_vm.ActivePlanningItem.PlanningItemIdWithCourse));
                    //    //ShowLogMsg(response_document + "\n");
                    //}
                    //else
                    //{
                    //    log.Warn(string.Format("Upload for '{0}' was not succesfull. Here is the GateWayError-Message:\n", _vm.Patient.Id + "_" + _vm.ActivePlanningItem.PlanningItemIdWithCourse));
                    //    log.Warn(response_document + "\n");
                    //}

                    //LogManager.Shutdown();
                }
                
            }

        }
        public static string SendData(string request, bool bIsJson, string apiKey, string hostnameBox)
        {
            var sMediaTYpe = bIsJson ? "application/json" :
            "application/xml";
            var sResponse = System.String.Empty;
            using (var c = new HttpClient(new
            HttpClientHandler()
            { UseDefaultCredentials = true, PreAuthenticate = true }))
            {
                if (c.DefaultRequestHeaders.Contains("ApiKey"))
                {
                    c.DefaultRequestHeaders.Remove("ApiKey");
                }
                c.DefaultRequestHeaders.Add("ApiKey", apiKey);
                var hostName = hostnameBox;
                var port = "55051";
                var gatewayURL = $"https://{hostName}:{port}/Gateway/service.svc/interop/rest/Process";
                var task =
                c.PostAsync(gatewayURL, new StringContent(request, Encoding.UTF8, sMediaTYpe));
                Task.WaitAll(task);
                var responseTask = task.Result.Content.ReadAsStringAsync();
                Task.WaitAll(responseTask);
                sResponse = responseTask.Result;
            }
            return sResponse;
        }
        string MakeFilenameValid(string s)
        {
            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (char ch in invalidChars)
            {
                s = s.Replace(ch, '_');
            }
            return s;
        }
        static void PrintPDF(string filePath, string printerName)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Die angegebene Datei wurde nicht gefunden.", filePath);

            using (System.Diagnostics.Process process = new System.Diagnostics.Process())
            {
                process.StartInfo.FileName = filePath;
                process.StartInfo.Verb = "print";
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                process.StartInfo.Arguments = $"/t \"{filePath}\" \"{printerName}\"";
                process.Start();
                process.WaitForExit();
            }
        }
        public static void SourcetoPng(BitmapSource bmp, string Path)
        {
            using (var fileStream = new FileStream(Path + @"\Drr.png", FileMode.Create))
            {
                BitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bmp));
                encoder.Save(fileStream);
                fileStream.Close();
            }
        }

        private static WriteableBitmap BuildDRRImage(Beam beam, VMS.TPS.Common.Model.API.Image Drr, double xMax, double yMax)
        {
            if (beam.ReferenceImage == null) { return null; }

            int whiteRand = 100;
            if (xMax > 30 || yMax > 30)
                whiteRand = 50;

            int[,] pixels = new int[Drr.YSize, Drr.XSize];
            Drr.GetVoxels(0, pixels);
            int[] flat_pixels = new int[Drr.YSize * Drr.XSize];

            for (int i = whiteRand; i < Drr.YSize - whiteRand; i++)
            {
                for (int j = whiteRand; j < Drr.XSize - whiteRand; j++)
                {
                    flat_pixels[j + Drr.XSize * i] = pixels[i, j];
                }
            }

            var Drr_max = flat_pixels.Max();
            var Drr_min = flat_pixels.Min();

            System.Windows.Media.PixelFormat format = PixelFormats.Gray8;
            int stride = (Drr.XSize * format.BitsPerPixel + 7) / 8;
            byte[] image_bytes = new byte[stride * Drr.YSize];
            // Alle Bytes auf 255 (Weiß) setzen
            // Set all bytes to 255 (white)
            Array.Copy(Enumerable.Repeat((byte)255, image_bytes.Length).ToArray(), image_bytes, image_bytes.Length);


            for (int i = whiteRand; i < Drr.YSize - whiteRand; i++)
            {
                for (int j = whiteRand; j < Drr.XSize - whiteRand; j++)
                {
                    double value = flat_pixels[j + Drr.XSize * i];

                    // Überprüfen, ob der Pixel sich im Randbereich befindet
                    if (i > Drr.YSize - whiteRand || j > Drr.XSize - whiteRand || j < whiteRand || i < whiteRand)
                    {
                        // Wenn ja, setzen Sie den Pixelwert auf 255 (Weiß)
                        image_bytes[j * stride + i] = 255;
                    }
                    else
                    {
                        // Andernfalls, verwenden Sie die vorhandene Logik zur Graustufenkonvertierung
                        image_bytes[j * stride + i] = Convert.ToByte(255 * ((value - Drr_min) / (Drr_max - Drr_min)));
                    }
                }
            }


            BitmapSource source = BitmapSource.Create(Drr.XSize, Drr.YSize, 25.4 / Drr.XRes, 25.4 / Drr.YRes, format, null, image_bytes, stride);

            WriteableBitmap writeableBitmap = new WriteableBitmap(source);

            return writeableBitmap;

        }


        private static Bitmap FieldLines(Beam beam, System.Drawing.Image source, VMS.TPS.Common.Model.API.Image Drr, PlanSetup plan, Patient p, string pID)
        {
            Bitmap bmp = new Bitmap(source);

            Bitmap bitmap = new Bitmap(bmp.Width * 4, bmp.Height * 4, bmp.PixelFormat);
            bitmap.SetResolution(bmp.HorizontalResolution, bmp.VerticalResolution);
            using (var gr = Graphics.FromImage(bitmap))
            {
                gr.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                gr.DrawImage(bmp, new System.Drawing.Rectangle(0, 0, bmp.Width * 4, bmp.Height * 4));
            }

            // Gantry Angle
            double GantryAngle = (beam.ControlPoints.FirstOrDefault().GantryAngle) * (Math.PI / 180.00);

            //Isocentre = center of image.
            double Centre_X = bitmap.Width / 2;
            double Centre_Y = bitmap.Height / 2;

            // Shifts to user orgin.
            double X_offset = (Math.Cos(GantryAngle) * (plan.StructureSet.Image.UserOrigin.x - beam.IsocenterPosition.x) + Math.Sin(GantryAngle) * (plan.StructureSet.Image.UserOrigin.y - beam.IsocenterPosition.y)) * 4;
            double Y_offset = (plan.StructureSet.Image.UserOrigin.z - beam.IsocenterPosition.z) * 4;

            // User origin.
            double user_origin_X = Centre_X + X_offset;
            double user_origin_Y = Centre_Y - Y_offset;


            // Shifts to 3D dose max 
            double DoseMax_Xoffset = 0;
            double DoseMax_Yoffset = 0;
            double DoseMax_X = 0;
            double DoseMax_Y = 0;
            if (plan.Dose != null)
            {
                plan.DoseValuePresentation = DoseValuePresentation.Absolute;
                DoseMax_Xoffset = (Math.Cos(GantryAngle) * (plan.Dose.DoseMax3DLocation.x - beam.IsocenterPosition.x) + Math.Sin(GantryAngle) * (plan.Dose.DoseMax3DLocation.y - beam.IsocenterPosition.y)) * 4;
                DoseMax_Yoffset = (plan.Dose.DoseMax3DLocation.z - beam.IsocenterPosition.z) * 4;

                // Dose Max
                DoseMax_X = Centre_X + DoseMax_Xoffset;
                DoseMax_Y = Centre_Y - DoseMax_Yoffset;
            }

            // MLC properties
            // Leaf offset
            List<double> LeafOffset = new List<double> { -195, -185, -175, -165, -155, -145, -135, -125, -115, -105, -97.5, -92.5, -87.5, -82.5, -77.5, -72.5, -67.5, -62.5, -57.5, -52.5, -47.5, -42.5, -37.5, -32.5, -27.5, -22.5, -17.5, -12.5, -7.5, -2.5, 2.5, 7.5, 12.5, 17.5, 22.5, 27.5, 32.5, 37.5, 42.5, 47.5, 52.5, 57.5, 62.5, 67.5, 72.5, 77.5, 82.5, 87.5, 92.5, 97.5, 105, 115, 125, 135, 145, 155, 165, 175, 185, 195 };
            // Leaf Widths
            List<double> LeafWidths = new List<double> { 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10 };
            //override for HD-MLC
            if (beam.MLC != null)
            {
                if (beam.MLC.Model == "Varian High Definition 120")
                {
                    LeafWidths = new List<double> { 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 2.5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 };
                    LeafOffset = new List<double> { -107.5, -102.5, -97.5, -92.5, -87.5, -82.5, -77.5, -72.5, -67.5, -62.5, -57.5, -52.5, -47.5, -42.5, -38.75, -36.25, -33.75, -31.25, -28.75, -26.25, -23.75, -21.25, -18.75, -16.25, -13.75, -11.25, -8.75, -6.25, -3.75, -1.25, 1.25, 3.75, 6.25, 8.75, 11.25, 13.75, 16.25, 18.75, 21.25, 23.75, 26.25, 28.75, 31.25, 33.75, 36.25, 38.75, 42.5, 47.5, 52.5, 57.5, 62.5, 67.5, 72.5, 77.5, 82.5, 87.5, 92.5, 97.5, 102.5, 107.5 };
                }
            }
            // Leaf positions
            List<double> Bank0Positions = new List<double>();
            List<double> Bank1Positions = new List<double>();
            
            if (beam.MLC != null)
            {
                var LeafPostions = beam.ControlPoints.ElementAtOrDefault(0).LeafPositions;
               
                for (var i = 0; i < 60; i++)
                {
                    Bank0Positions.Add(LeafPostions[0, i]);
                    Bank1Positions.Add(LeafPostions[1, i]);
                }
            }


            double collimatorAngle = (beam.ControlPoints.FirstOrDefault().CollimatorAngle) * (Math.PI / 180.00);

            VRect<double> jawPositions = beam.ControlPoints.FirstOrDefault().JawPositions;

            double Y2_CenterX = Centre_X - (jawPositions.Y2 * Math.Sin(collimatorAngle)) * 4;
            double Y2_CenterY = Centre_Y - (jawPositions.Y2 * Math.Cos(collimatorAngle)) * 4;

            double Y1_CenterX = Centre_X - (jawPositions.Y1 * Math.Sin(collimatorAngle)) * 4;
            double Y1_CenterY = Centre_Y - (jawPositions.Y1 * Math.Cos(collimatorAngle)) * 4;


            double Y2_UpperX = Y2_CenterX + jawPositions.X2 * Math.Cos(collimatorAngle) * 4;
            double Y2_UpperY = Y2_CenterY - jawPositions.X2 * Math.Sin(collimatorAngle) * 4;

            double Y2_LowerX = Y2_CenterX + jawPositions.X1 * Math.Cos(collimatorAngle) * 4;
            double Y2_LowerY = Y2_CenterY - jawPositions.X1 * Math.Sin(collimatorAngle) * 4;

            double Y1_UpperX = Y1_CenterX + jawPositions.X2 * Math.Cos(collimatorAngle) * 4;
            double Y1_UpperY = Y1_CenterY - jawPositions.X2 * Math.Sin(collimatorAngle) * 4;

            double Y1_LowerX = Y1_CenterX + jawPositions.X1 * Math.Cos(collimatorAngle) * 4;
            double Y1_LowerY = Y1_CenterY - jawPositions.X1 * Math.Sin(collimatorAngle) * 4;

            int Graticule_Length = (int)(Math.Round(Drr.XSize * 2 / 100d, 0) * 100) * 4;

            int Graticule_PositiveX = (int)(Graticule_Length * Math.Sin(collimatorAngle)) * 4;
            int Graticule_PositiveY = (int)(Graticule_Length * Math.Cos(collimatorAngle)) * 4;

            System.Drawing.Pen fieldPen = new System.Drawing.Pen(System.Drawing.Color.Yellow, 3);
            if (beam.IsSetupField)
                fieldPen = new System.Drawing.Pen(System.Drawing.Color.Cyan, 3);
            System.Drawing.Pen isoPen = new System.Drawing.Pen(System.Drawing.Color.Red, 3);
            System.Drawing.Pen graticulePen = new System.Drawing.Pen(System.Drawing.Color.Yellow, (float)0.5);
            System.Drawing.Pen referncePointPen = new System.Drawing.Pen(System.Drawing.Color.Cyan, 3);


            System.Drawing.Font FieldFont = new Font("Arial", 24, System.Drawing.FontStyle.Bold);
            Font TextFont = new Font("Arial", 16, System.Drawing.FontStyle.Bold);
            System.Drawing.SolidBrush FieldText = new System.Drawing.SolidBrush(System.Drawing.Color.Red);
            System.Drawing.SolidBrush ReferencePointText = new System.Drawing.SolidBrush(System.Drawing.Color.Cyan);



            using (var graphics = Graphics.FromImage(bitmap))
            {
                //adding field-Info to the bottom of the DRR
                string text = beam.Id +"  ("+pID + "_C: " + plan.Course.Id + "_P: " + plan.Id + ")";
                Font font = new Font("Arial", 20, System.Drawing.FontStyle.Bold);
                SizeF textSize = graphics.MeasureString(text, font);
                PointF textPosition = new PointF(0, 8);
                graphics.DrawString(text, font, FieldText, textPosition);


                graphics.DrawLine(isoPen, Convert.ToInt32(user_origin_X), Convert.ToInt32(user_origin_Y) - 10, Convert.ToInt32(user_origin_X), Convert.ToInt32(user_origin_Y) + 10);
                graphics.DrawLine(fieldPen, Convert.ToInt32(user_origin_X - 10), Convert.ToInt32(user_origin_Y), Convert.ToInt32(user_origin_X + 10), Convert.ToInt32(user_origin_Y));

                graphics.DrawEllipse(isoPen, Convert.ToInt32(Centre_X - 5), Convert.ToInt32(Centre_Y - 5), 10, 10);

                graphics.DrawLine(graticulePen, Convert.ToInt32(Centre_X + Graticule_PositiveX), Convert.ToInt32(Centre_Y + Graticule_PositiveY), Convert.ToInt32(Centre_X - Graticule_PositiveX), Convert.ToInt32(Centre_Y - Graticule_PositiveY));
                graphics.DrawLine(graticulePen, Convert.ToInt32(Centre_X + Graticule_PositiveY), Convert.ToInt32(Centre_Y - Graticule_PositiveX), Convert.ToInt32(Centre_X - Graticule_PositiveY), Convert.ToInt32(Centre_Y + Graticule_PositiveX));

                for (int A = -Graticule_Length; A < Graticule_Length; A += 10)
                {
                    double Marker_CenterX1 = Centre_X + (A * Math.Sin(collimatorAngle)) * 4;
                    double Marker_CenterY1 = Centre_Y + (A * Math.Cos(collimatorAngle)) * 4;
                    double Marker_CenterX2 = Centre_X + (A * Math.Cos(collimatorAngle)) * 4;
                    double Marker_CenterY2 = Centre_Y - (A * Math.Sin(collimatorAngle)) * 4;


                    if (A % 50 == 0)
                    {
                        double Marker_Cosine = (10 * Math.Cos(collimatorAngle)) * 4;
                        double Marker_Sine = (10 * Math.Sin(collimatorAngle)) * 4;
                        graphics.DrawLine(graticulePen, Convert.ToInt32(Marker_CenterX1 + Marker_Cosine), Convert.ToInt32(Marker_CenterY1 - Marker_Sine), Convert.ToInt32(Marker_CenterX1 - Marker_Cosine), Convert.ToInt32(Marker_CenterY1 + Marker_Sine));
                        graphics.DrawLine(graticulePen, Convert.ToInt32(Marker_CenterX2 + Marker_Sine), Convert.ToInt32(Marker_CenterY2 + Marker_Cosine), Convert.ToInt32(Marker_CenterX2 - Marker_Sine), Convert.ToInt32(Marker_CenterY2 - Marker_Cosine));

                    }
                    else
                    {
                        double Marker_Cosine = (5 * Math.Cos(collimatorAngle)) * 4;
                        double Marker_Sine = (5 * Math.Sin(collimatorAngle)) * 4;
                        graphics.DrawLine(graticulePen, Convert.ToInt32(Marker_CenterX1 + Marker_Cosine), Convert.ToInt32(Marker_CenterY1 - Marker_Sine), Convert.ToInt32(Marker_CenterX1 - Marker_Cosine), Convert.ToInt32(Marker_CenterY1 + Marker_Sine));
                        graphics.DrawLine(graticulePen, Convert.ToInt32(Marker_CenterX2 + Marker_Sine), Convert.ToInt32(Marker_CenterY2 + Marker_Cosine), Convert.ToInt32(Marker_CenterX2 - Marker_Sine), Convert.ToInt32(Marker_CenterY2 - Marker_Cosine));

                    }


                }

                // MLCs
                if (beam.MLC != null)
                {
                    for (int l = 0; l < 60; l++)
                    {
                        try
                        {
                            if (LeafOffset.ElementAt(l) + 10 > jawPositions.Y1 && LeafOffset.ElementAt(l) - 10 < jawPositions.Y2)
                            {
                                double MLC_Edge1X_atCenter = Centre_X - ((LeafOffset.ElementAt(l) - LeafWidths.ElementAt(l) / 2) * Math.Sin(collimatorAngle)) * 4;
                                double MlC_Edge1Y_atCenter = Centre_Y - ((LeafOffset.ElementAt(l) - LeafWidths.ElementAt(l) / 2) * Math.Cos(collimatorAngle)) * 4;

                                double MLC_Edge2X_atCenter = Centre_X - ((LeafOffset.ElementAt(l) + LeafWidths.ElementAt(l) / 2) * Math.Sin(collimatorAngle)) * 4;
                                double MlC_Edge2Y_atCenter = Centre_Y - ((LeafOffset.ElementAt(l) + LeafWidths.ElementAt(l) / 2) * Math.Cos(collimatorAngle)) * 4;

                                //if (Math.Abs(Bank0Positions.ElementAt(l)) < Math.Abs(jawPositions.X1))
                                //{
                                    double MLC0_Edge1X = MLC_Edge1X_atCenter + (Bank0Positions.ElementAt(l) * Math.Cos(collimatorAngle)) * 4;
                                    double MLC0_Edge1Y = MlC_Edge1Y_atCenter - (Bank0Positions.ElementAt(l) * Math.Sin(collimatorAngle)) * 4;

                                    if (Bank0Positions.ElementAt(l - 1) != Bank0Positions.ElementAt(l) && LeafOffset.ElementAt(l - 1)+10  > jawPositions.Y1)
                                    {
                                        if (Math.Abs(Bank0Positions.ElementAt(l - 1)) < Math.Abs(jawPositions.X1))
                                        {
                                            double MLC0_Side1X = MLC_Edge1X_atCenter + (Bank0Positions.ElementAt(l - 1) * Math.Cos(collimatorAngle)) * 4;
                                            double MLC0_Side1Y = MlC_Edge1Y_atCenter - (Bank0Positions.ElementAt(l - 1) * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC0_Edge1X), Convert.ToInt32(MLC0_Edge1Y), Convert.ToInt32(MLC0_Side1X), Convert.ToInt32(MLC0_Side1Y));
                                        }
                                        else
                                        {
                                            double MLC0_Side1X = MLC_Edge1X_atCenter + (jawPositions.X1 * Math.Cos(collimatorAngle)) * 4;
                                            double MLC0_Side1Y = MlC_Edge1Y_atCenter - (jawPositions.X1 * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC0_Edge1X), Convert.ToInt32(MLC0_Edge1Y), Convert.ToInt32(MLC0_Side1X), Convert.ToInt32(MLC0_Side1Y));
                                        }

                                    }

                                    double MLC0_Edge2X = MLC_Edge2X_atCenter + (Bank0Positions.ElementAt(l) * Math.Cos(collimatorAngle)) * 4;
                                    double MlC0_Edge2Y = MlC_Edge2Y_atCenter - (Bank0Positions.ElementAt(l) * Math.Sin(collimatorAngle)) * 4;

                                    if (Bank0Positions.ElementAt(l + 1) != Bank0Positions.ElementAt(l) && LeafOffset.ElementAt(l)-10  < jawPositions.Y2)
                                    {
                                        if (Math.Abs(Bank0Positions.ElementAt(l + 1)) < Math.Abs(jawPositions.X1))
                                        {
                                            double MLC0_Side2X = MLC_Edge2X_atCenter + (Bank0Positions.ElementAt(l + 1) * Math.Cos(collimatorAngle)) * 4;
                                            double MLC0_Side2Y = MlC_Edge2Y_atCenter - (Bank0Positions.ElementAt(l + 1) * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC0_Edge2X), Convert.ToInt32(MlC0_Edge2Y), Convert.ToInt32(MLC0_Side2X), Convert.ToInt32(MLC0_Side2Y));
                                        }
                                        else
                                        {
                                            double MLC0_Side2X = MLC_Edge2X_atCenter + (jawPositions.X1 * Math.Cos(collimatorAngle)) * 4;
                                            double MLC0_Side2Y = MlC_Edge2Y_atCenter - (jawPositions.X1 * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC0_Edge2X), Convert.ToInt32(MlC0_Edge2Y), Convert.ToInt32(MLC0_Side2X), Convert.ToInt32(MLC0_Side2Y));
                                        }

                                    }

                                    graphics.DrawLine(fieldPen, Convert.ToInt32(MLC0_Edge1X), Convert.ToInt32(MLC0_Edge1Y), Convert.ToInt32(MLC0_Edge2X), Convert.ToInt32(MlC0_Edge2Y));
                               
                                    double MLC1_Edge1X = MLC_Edge1X_atCenter + (Bank1Positions.ElementAt(l) * Math.Cos(collimatorAngle)) * 4;
                                    double MLC1_Edge1Y = MlC_Edge1Y_atCenter - (Bank1Positions.ElementAt(l) * Math.Sin(collimatorAngle)) * 4;

                                    if (Bank1Positions.ElementAt(l - 1) != Bank1Positions.ElementAt(l) && LeafOffset.ElementAt(l)+10  > jawPositions.Y1)
                                    {
                                        if (Math.Abs(Bank1Positions.ElementAt(l - 1)) < Math.Abs(jawPositions.X2))
                                        {
                                            double MLC1_Side1X = MLC_Edge1X_atCenter + (Bank1Positions.ElementAt(l - 1) * Math.Cos(collimatorAngle)) * 4;
                                            double MLC1_Side1Y = MlC_Edge1Y_atCenter - (Bank1Positions.ElementAt(l - 1) * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC1_Edge1X), Convert.ToInt32(MLC1_Edge1Y), Convert.ToInt32(MLC1_Side1X), Convert.ToInt32(MLC1_Side1Y));
                                        }
                                        else
                                        {
                                            double MLC1_Side1X = MLC_Edge1X_atCenter + (jawPositions.X2 * Math.Cos(collimatorAngle)) * 4;
                                            double MLC1_Side1Y = MlC_Edge1Y_atCenter - (jawPositions.X2 * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC1_Edge1X), Convert.ToInt32(MLC1_Edge1Y), Convert.ToInt32(MLC1_Side1X), Convert.ToInt32(MLC1_Side1Y));
                                        }

                                    }

                                    double MLC1_Edge2X = MLC_Edge2X_atCenter + (Bank1Positions.ElementAt(l) * Math.Cos(collimatorAngle)) * 4;
                                    double MlC1_Edge2Y = MlC_Edge2Y_atCenter - (Bank1Positions.ElementAt(l) * Math.Sin(collimatorAngle)) * 4;

                                    if (Bank1Positions.ElementAt(l + 1) != Bank1Positions.ElementAt(l) && LeafOffset.ElementAt(l)-10  < jawPositions.Y2)
                                    {
                                        if (Math.Abs(Bank1Positions.ElementAt(l + 1)) < Math.Abs(jawPositions.X2))
                                        {
                                            double MLC1_Side2X = MLC_Edge2X_atCenter + (Bank1Positions.ElementAt(l + 1) * Math.Cos(collimatorAngle)) * 4;
                                            double MLC1_Side2Y = MlC_Edge2Y_atCenter - (Bank1Positions.ElementAt(l + 1) * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC1_Edge2X), Convert.ToInt32(MlC1_Edge2Y), Convert.ToInt32(MLC1_Side2X), Convert.ToInt32(MLC1_Side2Y));
                                        }
                                        else
                                        {
                                            double MLC1_Side2X = MLC_Edge2X_atCenter + (jawPositions.X2 * Math.Cos(collimatorAngle)) * 4;
                                            double MLC1_Side2Y = MlC_Edge2Y_atCenter - (jawPositions.X2 * Math.Sin(collimatorAngle)) * 4;

                                            graphics.DrawLine(fieldPen, Convert.ToInt32(MLC1_Edge2X), Convert.ToInt32(MlC1_Edge2Y), Convert.ToInt32(MLC1_Side2X), Convert.ToInt32(MLC1_Side2Y));
                                        }

                                    }

                                    graphics.DrawLine(fieldPen, Convert.ToInt32(MLC1_Edge1X), Convert.ToInt32(MLC1_Edge1Y), Convert.ToInt32(MLC1_Edge2X), Convert.ToInt32(MlC1_Edge2Y));
                                //}
                            }
                        }
                        catch
                        { 
                        }
                    }
                }


                StringFormat sf = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Alignment = StringAlignment.Center
                };


                graphics.DrawLine(fieldPen, Convert.ToInt32(Y2_LowerX), Convert.ToInt32(Y2_LowerY), Convert.ToInt32(Y2_UpperX), Convert.ToInt32(Y2_UpperY));
                
                graphics.DrawLine(fieldPen, Convert.ToInt32(Y1_LowerX), Convert.ToInt32(Y1_LowerY), Convert.ToInt32(Y1_UpperX), Convert.ToInt32(Y1_UpperY));
                
                graphics.DrawLine(fieldPen, Convert.ToInt32(Y2_LowerX), Convert.ToInt32(Y2_LowerY), Convert.ToInt32(Y1_LowerX), Convert.ToInt32(Y1_LowerY));
                
                graphics.DrawLine(fieldPen, Convert.ToInt32(Y2_UpperX), Convert.ToInt32(Y2_UpperY), Convert.ToInt32(Y1_UpperX), Convert.ToInt32(Y1_UpperY));
                
                if (beam.ControlPoints.FirstOrDefault().CollimatorAngle > 90)
                {
                    graphics.DrawString("Y2", FieldFont, FieldText, Convert.ToInt32((Y2_LowerX + Y2_UpperX) / 2), Convert.ToInt32((Y2_LowerY + Y2_UpperY) / 2));
                    graphics.DrawString("Y1", FieldFont, FieldText, Convert.ToInt32((Y1_LowerX + Y1_UpperX) / 2), Convert.ToInt32((Y1_LowerY + Y1_UpperY) / 2));
                    graphics.DrawString("X1", FieldFont, FieldText, Convert.ToInt32((Y2_LowerX + Y1_LowerX) / 2), Convert.ToInt32((Y2_LowerY + Y1_LowerY) / 2));
                    graphics.DrawString("X2", FieldFont, FieldText, Convert.ToInt32((Y2_UpperX + Y1_UpperX) / 2), Convert.ToInt32((Y2_UpperY + Y1_UpperY) / 2));
                }
                // ansonten der default
                else
                {
                    graphics.DrawString("Y2", FieldFont, FieldText, new Rectangle(Math.Min(Convert.ToInt32(Y2_LowerX), Convert.ToInt32(Y2_UpperX)) - 20, Math.Min(Convert.ToInt32(Y2_LowerY), Convert.ToInt32(Y2_UpperY)) - 20, (Convert.ToInt32(Y2_UpperX) - Convert.ToInt32(Y2_LowerX)), (Convert.ToInt32(Y2_LowerY) - Convert.ToInt32(Y2_UpperY))), sf);
                    graphics.DrawString("Y1", FieldFont, FieldText, new Rectangle(Math.Min(Convert.ToInt32(Y1_LowerX), Convert.ToInt32(Y1_UpperX)) + 20, Math.Min(Convert.ToInt32(Y1_LowerY), Convert.ToInt32(Y1_UpperY)) + 20, (Convert.ToInt32(Y1_UpperX) - Convert.ToInt32(Y1_LowerX)), (Convert.ToInt32(Y1_LowerY) - Convert.ToInt32(Y1_UpperY))), sf);
                    graphics.DrawString("X1", FieldFont, FieldText, new Rectangle(Math.Min(Convert.ToInt32(Y1_LowerX), Convert.ToInt32(Y2_LowerX)) - 20, Math.Min(Convert.ToInt32(Y1_LowerY), Convert.ToInt32(Y2_LowerY)) + 20, (Convert.ToInt32(Y1_LowerX) - Convert.ToInt32(Y2_LowerX)), (Convert.ToInt32(Y1_LowerY) - Convert.ToInt32(Y2_LowerY))), sf);
                    graphics.DrawString("X2", FieldFont, FieldText, new Rectangle(Math.Min(Convert.ToInt32(Y1_UpperX), Convert.ToInt32(Y2_UpperX)) + 20, Math.Min(Convert.ToInt32(Y1_UpperY), Convert.ToInt32(Y2_UpperY)) - 20, (Convert.ToInt32(Y1_UpperX) - Convert.ToInt32(Y2_UpperX)), (Convert.ToInt32(Y1_UpperY) - Convert.ToInt32(Y2_UpperY))), sf);

                }

                if (plan.Dose != null)
                {
                    graphics.FillEllipse(FieldText, Convert.ToInt32(DoseMax_X), Convert.ToInt32(DoseMax_Y), 5, 5);
                    graphics.DrawString(plan.Dose.DoseMax3D.ValueAsString + " Gy", TextFont, FieldText, Convert.ToInt32(DoseMax_X) + 10, Convert.ToInt32(DoseMax_Y) - 5);
                }
            }




            return bitmap;

        }

        private Sex GetPatientSex(string sex)
        {
            switch (sex)
            {
                case "Male": return Sex.Male;
                case "Männlich": return Sex.Male;
                case "Female": return Sex.Female;
                case "Weiblich": return Sex.Female;
                case "Other": return Sex.Other;
                case "Unbekannt": return Sex.Other;
                case "Unknown": return Sex.Other;
                default: throw new ArgumentOutOfRangeException();
            }

        }
        private ReportData CreateReportData()
        {
            var reportPQMs = new ReportPQMs();
            int numPqms = _vm.PqmSummaries.Count();
            ReportPQM[] reportPQMList = new ReportPQM[numPqms];
            var reportPQM = new ReportPQM();

            // for right PlanSum approver selection
            int k = 0;
            foreach (var PlanningItemObject in _vm.PlanningItemSummaries)
            {
                if (_vm.ActivePlanningItem.PlanningItemIdWithCourse == _vm.PlanningItemSummaries[k].PlanningItemIdWithCourse)
                {
                    break;
                }
                k++;
            }
            //
            int i = 0;
            foreach (var pqm in _vm.PqmSummaries)
            {
                reportPQMList[i] = new ReportPQM();
                reportPQMList[i].Achieved = pqm.Achieved;
                reportPQMList[i].DVHObjective = pqm.DVHObjective;
                reportPQMList[i].Goal = pqm.Goal;
                reportPQMList[i].StructureNameWithCode = pqm.Structure.StructureNameWithCode.Length > 16 ? pqm.Structure.StructureNameWithCode.Substring(0, 14) + ".." : pqm.Structure.StructureNameWithCode;
                reportPQMList[i].StructVolume = pqm.StructVolume;
                reportPQMList[i].TemplateId = pqm.TemplateId.Length>12? pqm.TemplateId.Substring(0,11)+"..": pqm.TemplateId;
                reportPQMList[i].Variation = pqm.Variation;
                reportPQMList[i].Met = pqm.Met;
                reportPQMList[i].Ignore = pqm.Ignore;
                i++;
            }
            reportPQMs.PQMs = reportPQMList;
            

            var reportPCs = new ReportPCs();
            int numPCs = _vm.ErrorGrid.Count();
            ReportPC[] reportPCList = new ReportPC[numPCs];
            var reportPC = new ReportPC();

            int j = 0;
            foreach (var pc in _vm.ErrorGrid)
            {
                reportPCList[j] = new ReportPC();
                reportPCList[j].Description = pc.Description;
                reportPCList[j].Status = pc.Status;
                j++;
            }
            reportPCs.PCs = reportPCList;
            //////////////////////////////////////////
            var reportRefs = new ReportRefs();
            int numRefs = _vm.RefGrid.Count();
            ReportRef[] reportRefList = new ReportRef[numRefs];
            var reportRef = new ReportRef();

            int h = 0;
            foreach (var re in _vm.RefGrid)
            {
                reportRefList[h] = new ReportRef();
                reportRefList[h].RefPointId = re.RefPointId;
                reportRefList[h].Prescription = re.Prescription;
                reportRefList[h].Session = re.Session;
                reportRefList[h].Fractions = re.Fractions;
                reportRefList[h].D50= re.D50;
                reportRefList[h].D95 = re.D95;
                reportRefList[h].D2 = re.D2;
                reportRefList[h].D98 = re.D98;
                h++;
            }
            reportRefs.Refs = reportRefList;
            ///////////////////////////////////////////
            var reportDVHstats = new ReportDVHstats();
            int numDVHstats = SelectedDVHs.Count();
            ReportDVHstat[] reportDVHstatList = new ReportDVHstat[numDVHstats];
            var reportDVHstat = new ReportDVHstat();
            

            h = 0;
            foreach (var DVHselectStructure in SelectedDVHs)
            {
                try
                {
                    Structure s = _vm.ActivePlanningItem.PlanningItemObject.StructureSet.Structures.Single(st => st.Id == DVHselectStructure);
                    DVHData dvhData = _vm.ActivePlanningItem.PlanningItemObject.GetDVHCumulativeData(s, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.01);
                    double median = Math.Round(_vm.ActivePlanningItem.PlanningItemObject.GetDoseAtVolume(s, 50, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 2);
                    double D1 = Math.Round(_vm.ActivePlanningItem.PlanningItemObject.GetDoseAtVolume(s, 1, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 2);
                    double D99 = Math.Round(_vm.ActivePlanningItem.PlanningItemObject.GetDoseAtVolume(s, 99, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 2);
                    double mean = Math.Round(dvhData.MeanDose.Dose, 2);
                    double max = Math.Round(dvhData.MaxDose.Dose, 2);
                    double min = Math.Round(dvhData.MinDose.Dose, 2);
                    reportDVHstatList[h] = new ReportDVHstat();
                    reportDVHstatList[h].dsStructureID = s.Id;
                    reportDVHstatList[h].dsVolume = Math.Round(s.Volume, 2);
                    reportDVHstatList[h].dsMaxNear = max.ToString("0.00") + "_" + D1.ToString("0.00");
                    reportDVHstatList[h].dsMinNear = min.ToString("0.00") + "_" + D99.ToString("0.00");
                    reportDVHstatList[h].dsMeanMedian = mean.ToString("0.00") + "_" + median.ToString("0.00");
                    h++;
                }
                catch { }
            }
            reportDVHstats.DVHstats = reportDVHstatList;

                //////////////////////////////////////////////
            return new ReportData
            {
                ReportPatient = new ReportPatient
                {

                    Id = Anonymize_CheckBox.IsChecked == true? NewID_TextBox.Text : _vm.Patient.Id,
                    planId = _vm.ActivePlanningItem.PlanningItemId,
                    planType = _vm.ActivePlanningItem.PlanningItemType,
                    FirstName = Anonymize_CheckBox.IsChecked == true ? "anonym" : _vm.Patient.FirstName,
                    userId = _vm.User.Name.Replace(@"oncology\", ""),
                    LastName = Anonymize_CheckBox.IsChecked == true ? NewID_TextBox.Text : _vm.Patient.LastName,
                    Sex = GetPatientSex(_vm.Patient.Sex),
                    Birthdate = _vm.Patient.DateOfBirth.ToString() == "" ? _vm.Patient.HistoryDateTime : (Anonymize_CheckBox.IsChecked == true ? DateTime.Now : (DateTime)_vm.Patient.DateOfBirth),
                    Doctor = new Doctor
                    {
                        Name = _vm.Patient.PrimaryOncologistId
                    }           

                },
                ReportPlanningItem = new ReportPlanningItem
                {
                    Id = _vm.ActivePlanningItem.PlanningItemIdWithCourse,
                    Type = _vm.ActivePlanningItem.PlanningItemType,
                    Created = _vm.ActivePlanningItem.Creation,
                    TreatmentApproverDisplayName = _vm.ActivePlanningItem.PlanningItemType == "Plan" ?
                        _vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName : _vm.PlanningItemSummaries[k].TreatmentApprover,
                    PlansInPlansum = _vm.PlanningItemSummaries[k].PlanCount,
                    PrintPC_CheckboxState = PrintPC_CB.IsChecked.Value,
                    ConstraintPath = _vm.ActiveConstraintPath.ConstraintName,
                    prinTtotaldose = _vm.ActivePlanningItem.PlanningItemType == "Plan" ? Math.Round(_vm.ActivePlanningItem.PlanningItemDose, 2) : 0,
                    prinTfractiondose = _vm.ActivePlanningItem.PlanningItemType == "Plan" ? Math.Round(_vm.ActivePlanningItem.PlanningItemDoseFx, 2) : 0,
                    prinTfractions = _vm.ActivePlanningItem.PlanningItemFx,
                    gatingExtenedString = _vm.ActivePlanningItem.PlanningItemType == "Plan" ? _vm.ActivePlanningItem.PlanningGatingExtenedString : "",
                    algorithmus = _vm.ActivePlanningItem.Algorithmus,
                    technique = _vm.ActivePlanningItem.Technique,
                    orientation = _vm.ActivePlanningItem.Orientation,
                    planningApprover = _vm.ActivePlanningItem.PlanningApprover == null ? "noPlanningApr" : _vm.ActivePlanningItem.PlanningApprover,
                    treatmentapproveDate = _vm.ActivePlanningItem.PlanningItemTreatmentApproveDate == null ? "noTreatApprovalDate" : _vm.ActivePlanningItem.PlanningItemTreatmentApproveDate,
                },
                ReportStructureSet = new ReportStructureSet
                {
                    Id = _vm.StructureSet.Id,
                    Image = new ReportImage
                    {
                        Id = _vm.Image.Id,
                        CreationTime = (DateTime)_vm.Image.Series.Images.First().CreationDateTime
                    }
                },
                ReportDVH = new ReportDVH
                {
                    _dvhBitmap = dvhBitmap
                },
                ReportPQMs = reportPQMs,
                ReportPCs = reportPCs,
                ReportRefs = reportRefs,
                ReportDVHstats = reportDVHstats,
            };
        }

        private static string GetTempPdfPath()
        {
           
            return Path.GetTempFileName();
        }

       

        

       

        private void UserControl_MouseMove(object sender, MouseEventArgs e)
        {
            Window window = Window.GetWindow(this);
            window.Topmost = false;
            if (w==1)
            {
                
                RefreshDVHClicked(null, null); 
                w=2;
                if (_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName != "" && _vm.ActivePlanningItem.PlanningItemType == "Plan")
                    DRR_CB.IsChecked = true;
                if (_vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName != "" )
                    eDoc_CB.IsChecked = true;

                LogLiveMining();
                List<CheckBox> checkBoxlist2 = new List<CheckBox>();

                // Find all elements
                FindChildGroup<CheckBox>(pqmDataGrid, "checkboxinstance2", ref checkBoxlist2);
                int j = 0;
                foreach (var pqm in _vm.PqmSummaries)
                {
                    if ((pqm.StructureName.ToString().Contains("TL") || pqm.StructureName.ToString().Contains("TB") || pqm.StructureName.ToString().Contains("TR") || pqm.StructureName.ToString().Contains("TL") || pqm.StructureName.ToString().ToLower().Contains("teil")) 
                        && pqm.Ignore.ToString() != "True" 
                        || pqm.Achieved.ToString().StartsWith("0.0")
                        || pqm.Achieved == "Not evaluated" || pqm.Achieved == "Volume too small" || pqm.Met == "Not evaluated"
                        & ! (pqm.StructureName.ToString().ToLower().Contains("ptv") || pqm.StructureName.ToString().ToLower().Contains("ctv") || pqm.StructureName.ToString().ToLower().Contains("gtv") || pqm.StructureName.ToString().ToLower().Contains("itv")))
                    {
                        checkBoxlist2[j].IsChecked = true;
                    }
                    j++;
                }
            }
        }

        

        private void PlanOpen_Button_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.GetWindow(this);
            _vm.ActivePlanningItem = GetPlan(sender);
            System.Reflection.Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            string scriptVersion = fvi.FileVersion;
            var mainViewModel = new MainViewModel(_vm.User, _vm.Patient, scriptVersion, _vm.PlanningItemList, _vm.ActivePlanningItem, GetPlan(sender).PlanningItemObject);
            //var viewModel = new DVHViewModel(selectedItem.PlanningItemObject);

            window.Title = mainViewModel.Title;
            var mainView = new MainView(mainViewModel);
            window.Content = mainView;
            //_psvm.MainWindow.Activate();
            window.Height = System.Windows.SystemParameters.PrimaryScreenHeight * 0.85;
            window.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;

            var performSort = typeof(DataGrid).GetMethod("PerformSort", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[7] });
            performSort?.Invoke(planningItemSummariesDataGrid, new[] { planningItemSummariesDataGrid.Columns[7] });
        }
        private void Compare_Click(object sender, RoutedEventArgs e)
        {
            int p = 0;
            foreach (var pi in _vm.PlanningItemList.Where(x => x.PlanningItemId != _vm.ActivePlanningItem.PlanningItemId && x.PlanningItemObject.IsDoseValid() && x.PlanningItemCourse == _vm.ActivePlanningItem.PlanningItemCourse))
            {
                p++;
            }
            
            if (p == 0)
            {
                MessageBox.Show("No plans for Comparison found.\n\nYou need more than 1 plan with valid dose in the active course.");
            }
            else
            {
                var comparisonLogger = new CustomLog("Comparison.csv");
                StringBuilder csvContent = new StringBuilder();
                string logsDir = _settings.GetLogsDirectory();
                string logFilePath = Path.Combine(logsDir, "Comparison.csv");

                // Sicherstellen, dass Logs-Ordner existiert
                if (!Directory.Exists(logsDir))
                {
                    Directory.CreateDirectory(logsDir);
                }

                bool fileExists = File.Exists(logFilePath);

                // Header nur schreiben, falls Datei noch nicht existiert
                if (!fileExists)
                {
                    List<string> headers = new List<string>
                    {
                        "User", "PC", "Script", "Date", "Time", "Patient",
                        "Compare-Course", "ActivePlan", "TotalDose", "FxDose", "Constraint-Table"
                    };
                            csvContent.AppendLine(string.Join(",", headers));
                        }

                        string userId = Environment.UserName.ToString();
                        string pc = Environment.MachineName;
                        string scriptId = "ClearPlan";
                        string date = DateTime.Now.ToString("yyyy-MM-dd");
                        string time = DateTime.Now.ToString("HH:mm");

                        List<string> row = new List<string>
                        {
                            userId, pc, scriptId, date, time,
                            _vm.Patient.Name.Replace(",", "^"),
                            MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemCourse),
                            MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemId),
                            Math.Round(_vm.ActivePlanningItem.PlanningItemDose, 2).ToString(),
                            Math.Round(_vm.ActivePlanningItem.PlanningItemDoseFx, 2).ToString(),
                            MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName)
                        };

                        csvContent.AppendLine(string.Join(",", row));

                        // PQM-Header für die detaillierte Vergleichsliste
                        List<string> pqmHeaders = new List<string>
                        {
                            "Template ID", "Structure", "Vol [cc]", "Objective", "Goal", MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemId) + "*"
                        };

                        foreach (var pi in _vm.PlanningItemList.Where(x =>
                            x.PlanningItemId != _vm.ActivePlanningItem.PlanningItemId &&
                            x.PlanningItemObject.IsDoseValid() &&
                            x.PlanningItemCourse == _vm.ActivePlanningItem.PlanningItemCourse))
                        {
                            pqmHeaders.Add(MakeFilenameValid(pi.PlanningItemId));
                        }

                        csvContent.AppendLine(string.Join(",", pqmHeaders));

                        // Vergleich mit anderen Plänen
                        DataTable otherPlans = new DataTable();
                        otherPlans.Columns.Add("Plan-Id", typeof(string));
                        otherPlans.Columns.Add("Structure-Id", typeof(string));
                        otherPlans.Columns.Add("Objective", typeof(string));
                        otherPlans.Columns.Add("Achieved", typeof(string));

                        int k = 0;
                        foreach (var pi in _vm.PlanningItemList.Where(x =>
                            x.PlanningItemId != _vm.ActivePlanningItem.PlanningItemId &&
                            x.PlanningItemObject.IsDoseValid() &&
                            x.PlanningItemCourse == _vm.ActivePlanningItem.PlanningItemCourse))
                        {
                            _vm.GetPQMSummaries(_vm.ActiveConstraintPath, pi, _vm.Patient);

                            foreach (var pqm in _vm.PqmSummaries)
                            {
                                DataRow rowEntry = otherPlans.NewRow();
                                rowEntry["Plan-Id"] = pi.PlanningItemId;
                                rowEntry["Structure-Id"] = pqm.StructureName;
                                rowEntry["Objective"] = pqm.DVHObjective;
                                rowEntry["Achieved"] = pqm.Achieved;
                                otherPlans.Rows.Add(rowEntry);
                                k++;
                            }
                        }

                        _vm.GetPQMSummaries(_vm.ActiveConstraintPath, _vm.ActivePlanningItem, _vm.Patient);

                        foreach (var pqm in _vm.PqmSummaries)
                        {
                            List<string> pqmRow = new List<string>
                            {
                                pqm.TemplateId, pqm.StructureName, pqm.StructVolume.ToString(),
                                pqm.DVHObjective, pqm.Goal, pqm.Achieved
                            };

                            foreach (var pi in _vm.PlanningItemList.Where(x =>
                                x.PlanningItemId != _vm.ActivePlanningItem.PlanningItemId &&
                                x.PlanningItemObject.IsDoseValid() &&
                                x.PlanningItemCourse == _vm.ActivePlanningItem.PlanningItemCourse))
                            {
                                bool foundConstraint = false;
                                for (int i = 0; i < k; i++)
                                {
                                    if (otherPlans.Rows[i]["Structure-Id"].ToString() == pqm.StructureName &&
                                        otherPlans.Rows[i]["Objective"].ToString() == pqm.DVHObjective &&
                                        pi.PlanningItemId == otherPlans.Rows[i]["Plan-Id"].ToString())
                                    {
                                        pqmRow.Add(otherPlans.Rows[i]["Achieved"].ToString());
                                        foundConstraint = true;
                                        break;
                                    }
                                }
                                if (!foundConstraint)
                                    pqmRow.Add("not found");
                            }

                            csvContent.AppendLine(string.Join(",", pqmRow));
                        }

                        File.AppendAllText(logFilePath, csvContent.ToString());
                        //comparisonLogger.Info("Comparison data logged.");


                //ile.AppendAllText(userLogPath, userLogCsvContent2.ToString());
                string filename = string.Format("{0}_{1}_{2}.csv", MakeFilenameValid(_vm.Patient.Id), MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemStructureSet.Id), MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName));
                string filenameHTML = string.Format("{0}_{1}_{2}.html", MakeFilenameValid(_vm.Patient.Id), MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemStructureSet.Id), MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName));

                DataTable csvPQMdata = csvToDataTable(logFilePath, true);
                DataTable csvPQMintro = csvToDataTable(logFilePath, false);

                string HtmlBody = ExportDatatableToHtml(csvPQMdata, csvPQMintro);
                string htmlPath = BuildLogsFilePath("Compare", filenameHTML);
                System.IO.File.WriteAllText(htmlPath, HtmlBody);
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = htmlPath;
                //info.Arguments = "chrome";
                info.UseShellExecute = true;
                File.Delete(logFilePath);
                try
                {
                    Process.Start("chrome", string.Format("\"{0}\"", htmlPath));
                }
                catch { Process.Start(info); }
            }
        }

        public static DataTable csvToDataTable(string file, bool pqm)
        {

            DataTable csvDataTable = new DataTable();

            //no try/catch - add these in yourselfs or let exception happen
            String[] csvData = File.ReadAllLines(file);

            //if no data in file ‘manually’ throw an exception
            if (csvData.Length == 0)
            {
                throw new Exception("CSV File Appears to be Empty");
            }

            String[] headings;
            //int index = 0;

            if (pqm) 
            {
                csvDataTable.Clear();
                headings = csvData[2].Split(',');
                //for each heading
                for (int i = 0; i < headings.Length; i++)
                {
                    //add a column for each heading
                    csvDataTable.Columns.Add(headings[i], typeof(string));
                }
                //populate the DataTable
                for (int i = 3; i < csvData.Length; i++)
                {
                    //create new rows
                    DataRow row = csvDataTable.NewRow();
                    String[] headingsPuffer = csvData[i].Split(',');

                    for (int j = 0; j < headingsPuffer.Length; j++)
                    {
                        //fill them
                        row[j] = csvData[i].Split(',')[j];
                    }

                    //add rows to over DataTable
                    csvDataTable.Rows.Add(row);
                }

                //return the CSV DataTable
                return csvDataTable;
            }
            else
            {
                csvDataTable.Clear();
                headings = null;
                headings = csvData[0].Split(',');
                //for each heading
                for (int i = 0; i < headings.Length; i++)
                {
                    //add a column for each heading
                    csvDataTable.Columns.Add(headings[i], typeof(string));
                }
                //populate the DataTable
                
                    //create new rows
                    DataRow row = csvDataTable.NewRow();

                    for (int j = 0; j < headings.Length; j++)
                    {
                        //fill them
                        row[j] = csvData[1].Split(',')[j];
                    }

                    //add rows to over DataTable
                    csvDataTable.Rows.Add(row);
                

                //return the CSV DataTable
                return csvDataTable;
            } 
        }
        string ExportDatatableToHtml(DataTable dt, DataTable dtIntro)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<link rel='stylesheet' href='dist/sortable-tables.min.css'>");
            strHTMLBuilder.Append("<script src='dist/sortable-tables.min.js'></script>");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("<title >");
            strHTMLBuilder.Append(_vm.Patient.Id+"_PQM-Comparison-Table");
            strHTMLBuilder.Append("</title>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='1' width='100%' style='border-collapse:collapse; font-family:'Lucida Sans', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif; font-size:17px'>");

            strHTMLBuilder.Append("<tr style='background-color: lightskyblue;'>");
            foreach (DataColumn myColumn in dtIntro.Columns)
            {
                strHTMLBuilder.Append("<th >");
                strHTMLBuilder.Append(myColumn.ColumnName);
                strHTMLBuilder.Append("</th>");

            }
            strHTMLBuilder.Append("</tr>");


            foreach (DataRow myRow in dtIntro.Rows)
            {

                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dtIntro.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");

                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</table>");

           strHTMLBuilder.Append(" <table id='table' class='sortable-table' border='1' cellpadding='1' cellspacing='1' width='100%' style='border-collapse:collapse; font-family:'FontAwesome', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif; font-size:17px'>");
            strHTMLBuilder.Append("<thead>");
            strHTMLBuilder.Append("<tr style='background-color: lightskyblue;'>");
            foreach (DataColumn myColumn in dt.Columns)
            {
                
                if (myColumn.ColumnName == "Template ID" || myColumn.ColumnName == "Structure" || myColumn.ColumnName == "Objective" || myColumn.ColumnName == "Goal")
                {
                    strHTMLBuilder.Append("<th>");
                    strHTMLBuilder.Append(myColumn.ColumnName);
                    strHTMLBuilder.Append("</th>");
                }
                else
                {
                    strHTMLBuilder.Append("<th class='numeric-sort'>");
                    strHTMLBuilder.Append(myColumn.ColumnName);
                    strHTMLBuilder.Append("</th>");
                }

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</thead>");
            strHTMLBuilder.Append("<tbody>");
            foreach (DataRow myRow in dt.Rows)
            {

                strHTMLBuilder.Append("<tr >");
                foreach (DataColumn myColumn in dt.Columns)
                {
                    
                        strHTMLBuilder.Append("<td>");
                        strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                        strHTMLBuilder.Append("</td>");
                    
                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</tbody>");
            //Close tags.
            strHTMLBuilder.Append("</table>");
            strHTMLBuilder.Append("<table border='0' cellpadding='1' cellspacing='1' width='100%' style='border-collapse:collapse; font-family:'Lucida Sans', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif; font-size:17px'>");
            strHTMLBuilder.Append("<tr>");
            strHTMLBuilder.Append("<td style='font-weight:bold'>");

            strHTMLBuilder.Append("Legend: ");


            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td>");

            strHTMLBuilder.Append("a) Which Plans will be compared?: All plans with valid dose in the active course.");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr>");

            strHTMLBuilder.Append("<tr >");
            strHTMLBuilder.Append("<td>");
            strHTMLBuilder.Append("b) Background colors in supported Browsers? green: best plan per row; red: worst plan per row.");
            strHTMLBuilder.Append("</td>");
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;

        }

        private void AddConstraintButtonClicked(object sender, RoutedEventArgs e)
        {
            var calculator = new PQMSummaryCalculator();
            PQMSummaryViewModel objective = new PQMSummaryViewModel();
            // Structure ID                
            objective.TemplateId =AddConstraintComboBox.Text.ToString();
            // Structure Code
            //string codes = "1";
            objective.TemplateCodes = new string[] { objective.TemplateId };
            // Aliases : extract individual aliases using "|" as separator.  Ignore whitespaces.  If blank, use the ID.
            //string aliases = line[2];
            objective.TemplateAliases =  new string[] { objective.TemplateId };
            // DVH Objective
            //string types = line[3];
            objective.TemplateType = new string[] { objective.TemplateId };
            // DVH Objective
            //objective.DVHObjective = "Mean[Gy]";
            if (DVHObjective1ComboBox.Text.Contains("_"))
                objective.DVHObjective = DVHObjective1ComboBox.Text.Replace("_", DVHObjective2TextBox.Text);
            else
                objective.DVHObjective = DVHObjective1ComboBox.Text;
            // Evaluator
            //objective.Goal = "<=35";
            objective.Goal = Goal1ComboBox.Text + Goal2TextBox.Text;
            //Variation
            //objective.Variation = "36";
            if (VariationTextBox.Text!= "")
                objective.Variation = VariationTextBox.Text;
            else
                objective.Variation = Goal2TextBox.Text;
            // Priority
            objective.Priority = "1";
            // Met (calculate this later, check if meeting - Goal, Variation, Not met)
            objective.Met = "";
            // Achieved (calculate this later)
            objective.Achieved = "";

            if (AddConstraintComboBox.Text.ToString() == "" || (DVHObjective1ComboBox.Text.Contains("_") && DVHObjective2TextBox.Text.ToString() == "") || Goal2TextBox.Text.ToString() == "")
                MessageBox.Show("Sorry, but not all requiered fields are populated. Please check your entries.", "Ooopss...", MessageBoxButton.OK, MessageBoxImage.Warning);
            else
            {
                Structure evalStructure = calculator.FindStructureFromAlias(_vm.StructureSet, _vm.ActivePlanningItem, objective.TemplateId, objective.TemplateAliases, objective.TemplateCodes, objective.TemplateType);
                if (evalStructure != null)
                {
                    var evalStructureVM = new StructureViewModel(evalStructure);
                    var obj = calculator.GetObjectiveProperties(objective, _vm.ActivePlanningItem, _vm.StructureSet, evalStructureVM);
                    //_vm.PqmSummaries.Add(obj);
                    _vm.PqmSummaries.Insert(0, obj);
                    _vm.NotifyPropertyChanged("Structure");
                }
                //UpdatePqmDataGrid();
                RefreshDVHClicked(null, null);
            }
        }

        private void Anonymize_Click(object sender, RoutedEventArgs e)
        {
            if (Anonymize_CheckBox.IsChecked == true)
            {
                NewID_TextBox.Visibility = Visibility.Visible;
                NewID_TextBlock.Visibility = Visibility.Visible;
            }
            else
            {
                NewID_TextBox.Visibility = Visibility.Collapsed;
                NewID_TextBlock.Visibility = Visibility.Collapsed;
            }
        }

        private void ChangeLog_Click(object sender, RoutedEventArgs e)
        {
            string version = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
            OpenDocumentation(
                "ClearPlan Change Log",
                _settings.GetChangeLogFile(),
                "No changelog file was found. Add CHANGELOG.md next to the built application or configure paths.changeLogFile in settings.json.",
                version);
        }

        private void Feedback_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_settings.Links.FeedbackUrl))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(_settings.Links.FeedbackUrl) { UseShellExecute = true });
                    return;
                }
                catch
                {
                }
            }

            OpenDocumentation(
                "ClearPlan Feedback",
                _settings.GetFeedbackFile(),
                "No feedback page was found. Add FEEDBACK.md next to the built application or configure paths.feedbackFile in settings.json.");
        }

        public static T FindChild<T>(DependencyObject parent, string childName)
          where T : DependencyObject
        {
            // Confirm parent and childName are valid.   
            if (parent == null) return null;

            T foundChild = null;

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                // If the child is not of the request child type child  
                T childType = child as T;
                if (childType == null)
                {
                    // recursively drill down the tree  
                    foundChild = FindChild<T>(child, childName);

                    // If the child is found, break so we do not overwrite the found child.   
                    if (foundChild != null) break;
                }
                else if (!string.IsNullOrEmpty(childName))
                {
                    var frameworkElement = child as FrameworkElement;
                    // If the child's name is set for search  
                    if (frameworkElement != null && frameworkElement.Name == childName)
                    {
                        // if the child's name is of the request name  
                        foundChild = (T)child;
                        break;
                    }
                }
                else
                {
                    // child element found.  
                    foundChild = (T)child;
                    break;
                }
            }

            return foundChild;
        }

        private void CSV_Button_Click(object sender, RoutedEventArgs e)
        {           
            #region CSV-Export Part1
            string MakeFilenameValid(string s)
            {
                char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                foreach (char ch in invalidChars)
                {
                    s = s.Replace(ch, '_');
                }
                return s;
            }

            string userLogPath;
            string filename;

                filename = string.Format("{0}_{1}_{4}_{2}_{3}.csv", MakeFilenameValid(_vm.Patient.Name), MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemStructureSet.Id), MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName), System.DateTime.Now.ToLocalTime().ToString("yyMMddHHmmss"), MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemId));
                //string filenameHTML = string.Format("{0}_{1}_{2}.html", MakeFilenameValid(_vm.Patient.Id), MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemStructureSet.Id), MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName));

                StringBuilder userLogCsvContent = new StringBuilder();
                userLogPath = Path.Combine(_settings.GetExportsDirectory(), filename);


                // add headers if the file doesn't exist
                // list of target headers for desired dose stats
                // in this case I want to display the headers every time so i can verify which target the distance is being measured for
                // this is due to the inconsistency in target naming (PTV1/2 vs ptv45/79.2) -- these can be removed later when cleaning up the data
                if (File.Exists(userLogPath))
                {
                    File.Delete(userLogPath);
                }
                if (!File.Exists(userLogPath))
                {
                    List<string> dataHeaderList = new List<string>();
                    dataHeaderList.Add("User");
                    // dataHeaderList.Add("Domain");
                    dataHeaderList.Add("PC");
                    dataHeaderList.Add("Script");
                    //dataHeaderList.Add("Version");
                    dataHeaderList.Add("Date");
                    dataHeaderList.Add("Time");
                    dataHeaderList.Add("ImageCreationDate");
                    dataHeaderList.Add("Patient");
                    dataHeaderList.Add("Course");
                    dataHeaderList.Add("PlanId");
                    dataHeaderList.Add("TotalDose");
                    dataHeaderList.Add("FxDose");
                    dataHeaderList.Add("Constraint-Table");



                    string concatDataHeader = string.Join(";", dataHeaderList.ToArray());

                    userLogCsvContent.AppendLine(concatDataHeader);
                }


                List<object> userStatsList = new List<object>();

                var culture = new System.Globalization.CultureInfo("de-DE");
                var day2 = culture.DateTimeFormat.GetDayName(System.DateTime.Today.DayOfWeek);

                string pc = Environment.MachineName.ToString();
                //string domain = Environment.UserDomainName.ToString();
                string userId = Environment.UserName.ToString();
                string scriptId = "ClearPlan";
                string date = System.DateTime.Now.ToString("yyyy-MM-dd");
                // string dayOfWeek = day2;
                string time = string.Format("{0}:{1}", System.DateTime.Now.ToLocalTime().ToString("HH"), System.DateTime.Now.ToLocalTime().ToString("mm"));

                userStatsList.Add(userId);
                // userStatsList.Add(domain);
                userStatsList.Add(pc);
                userStatsList.Add(scriptId);
                // userStatsList.Add(version);
                userStatsList.Add(date);
                //userStatsList.Add(dayOfWeek);
                userStatsList.Add(time);
                userStatsList.Add(_vm.ActivePlanningItem.PlanningItemImage.CreationDateTime);
                    //ps.TreatmentSessions.Where(x => x.Status == TreatmentSessionStatus.Completed).Max(t => t.HistoryDateTime).ToString("yyyyMMdd"));
                userStatsList.Add(_vm.Patient.Name.Replace(",", "^"));
                userStatsList.Add(MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemCourse));
                userStatsList.Add(MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemId));
                userStatsList.Add(Math.Round(_vm.ActivePlanningItem.PlanningItemDose, 2));
                userStatsList.Add(Math.Round(_vm.ActivePlanningItem.PlanningItemDoseFx, 2));
                userStatsList.Add(MakeFilenameValid(_vm.ActiveConstraintPath.ConstraintName));


                string concatUserStats = string.Join(";", userStatsList.ToArray());

                userLogCsvContent.AppendLine(concatUserStats);


                List<string> dataHeaderList2 = new List<string>();
                //dataHeaderList2.Add("Plan");
                dataHeaderList2.Add("Template ID");
                dataHeaderList2.Add("Structure");
                dataHeaderList2.Add("Vol [cc]");
                dataHeaderList2.Add("Objective");
                dataHeaderList2.Add("Goal");
                dataHeaderList2.Add("Achieved");
                dataHeaderList2.Add(MakeFilenameValid(_vm.ActivePlanningItem.PlanningItemId) + "*");
                /*foreach (var pi in _vm.PlanningItemList.Where(x => x.PlanningItemId != _vm.ActivePlanningItem.PlanningItemId && x.PlanningItemObject.IsDoseValid() && x.PlanningItemCourse == _vm.ActivePlanningItem.PlanningItemCourse))
                {
                    dataHeaderList2.Add(MakeFilenameValid(pi.PlanningItemId));
                }*/

                string concatDataHeader2 = string.Join(";", dataHeaderList2.ToArray());

                userLogCsvContent.AppendLine(concatDataHeader2);

                File.AppendAllText(userLogPath, userLogCsvContent.ToString());
                userLogCsvContent.Clear();
                string concatUserStats2 = "";
                StringBuilder userLogCsvContent2 = new StringBuilder();
                List<object> userStatsList2 = new List<object>();

                #endregion CSV-Export Part1

                _vm.GetPQMSummaries(_vm.ActiveConstraintPath, _vm.ActivePlanningItem, _vm.Patient);

                List<CheckBox> checkBoxlist = new List<CheckBox>();
                List<CheckBox> checkBoxlist2 = new List<CheckBox>();

                // Find all elements
                //FindChildGroup<CheckBox>(pqmDataGrid, "checkboxinstance", ref checkBoxlist);
                FindChildGroup<CheckBox>(pqmDataGrid, "checkboxinstance2", ref checkBoxlist2);
                int j = 0;

                foreach (var pqm in _vm.PqmSummaries.Where(x => x.Met.ToString().ToUpper() != "NOT EVALUATED"))
                {

                    #region CSV-Report Part 2
                    if (checkBoxlist2[j].IsChecked != true)
                    {
                        //PQMs
                        userStatsList2.Clear();
                        //userStatsList2.Add(pi.PlanningItemId);
                        userStatsList2.Add(pqm.TemplateId);
                        userStatsList2.Add(pqm.StructureName);
                        userStatsList2.Add(pqm.StructVolume);
                        userStatsList2.Add(pqm.DVHObjective);
                        userStatsList2.Add(pqm.Goal);
                        // userStatsList2.Add(pqm.Variation);                
                        userStatsList2.Add(pqm.Achieved.Replace(" ","").Replace("%","").Replace("Gy","").Replace("cc",""));


                        concatUserStats2 = string.Join(";", userStatsList2.ToArray());

                        userLogCsvContent2.AppendLine(concatUserStats2);
                    }
                    j++;
                    #endregion
                }


                File.AppendAllText(userLogPath, userLogCsvContent2.ToString());
                MessageBox.Show("Metrics saved here:\n\n" + userLogPath);
            
            
        }

        private void InfoTafel_Click(object sender, RoutedEventArgs e)
        {
            /////
            ///// Neues Fenster erstellen
            ///
            int width = _vm.ActivePlanningItem.PlanningItemObject is PlanSetup ? 900 : 1400;
            int height = _vm.ActivePlanningItem.PlanningItemObject is PlanSetup ? 800 : 700;
            int fontsize = 20;
            Window infoWindow = new Window
            {
                Title = "Frühbesprechungs-InfoTafel",
                Width = width,
                Height = height,
                ResizeMode = ResizeMode.NoResize,
                Background = System.Windows.Media.Brushes.White,
                WindowStartupLocation = WindowStartupLocation.CenterScreen // Fenster in der Mitte des Bildschirms öffnen
            };

            // ScrollViewer erstellen
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            // Grid für tabellenähnliche Struktur erstellen
            Grid grid = new Grid
            {
                Margin = new Thickness(10)
            };

            // Dynamische Breitenberechnung
            int totalColumns = _vm.ActivePlanningItem.PlanningItemPlans.Any() ? 2 + _vm.ActivePlanningItem.PlanningItemPlans.Count() : 2; // Parameter-Spalte + Plan1 + evtl. zusätzliche Plan-Spalten
            int totalWidth = width;
            int columnWidth = totalWidth / (totalColumns); // Breite pro Spalte
            int charPerLine = columnWidth / 10; // Zeichen pro Zeile (ca. 10 Pixel pro Zeichen)

            // Spalten für Parameter und Werte hinzufügen
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) }); // Parameter-Spalte
           

            if (_vm.ActivePlanningItem.PlanningItemPlans.Any())
            {
                foreach (var plan in _vm.ActivePlanningItem.PlanningItemPlans)
                {
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Eine Spalte pro Plan
                }
            }
            else
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Werte-Spalte
            }

            // Sammlung von Daten und Übergabe an Display-Funktion
            List<Dictionary<string, string>> collectedData = new List<Dictionary<string, string>>();
            
            // Wenn `_vm.PlanningItemObject` ein `PlanSetup` ist
            if (_vm.ActivePlanningItem.PlanningItemObject is PlanSetup)
            {
                int failedPQM = 0;
                //int mining_pqmsTotal = _vm.PqmSummaries.Count();
                foreach (var pqm in _vm.PqmSummaries.Where(x => x.Met == "Not met"))
                {
                    //if(pqm.Met!= "Goal" && pqm.Met != "Variation")
                    failedPQM++;
                }
                
                string patient = _vm.Patient.Name.Replace(",", "^");
                string ssID = _vm.ActivePlanningItem.PlanningItemPlanSetup.StructureSet.Id;
                string planID = _vm.ActivePlanningItem.PlanningItemId;
                string courseID = _vm.ActivePlanningItem.PlanningItemPlanSetup.Course.Id;
                string dose = _vm.ActivePlanningItem.PlanningItemFx + "*" + Math.Round(_vm.ActivePlanningItem.PlanningItemDoseFx, 2) + "Gy=" + Math.Round(_vm.ActivePlanningItem.PlanningItemDose, 2) + "Gy";
                string target = _vm.ActivePlanningItem.PlanningItemTargetId;
                string refPoints = string.Join("; ", _vm.ActivePlanningItem.PlanningItemPlanSetup.ReferencePoints.Select(rp=>$"{rp.Id} ({Math.Round(rp.TotalDoseLimit.Dose,2)}Gy)"));
                int beamCount = _vm.ActivePlanningItem.PlanningItemPlanSetup.Beams.Where(x => !x.IsSetupField).Count();
                string technik = _vm.ActivePlanningItem.Technique;
                string linac = _vm.ActivePlanningItem.PlanningItemPlanSetup.Beams.FirstOrDefault() != null ? _vm.ActivePlanningItem.PlanningItemPlanSetup.Beams.First().TreatmentUnit.Id.ToString().Substring(0, 6) : "NaN";
                //int failedPQM = 3; // Dummy-Wert
                string neDate = "Not configured";
                string planningApproval = _vm.ActivePlanningItem.PlanningItemPlanSetup.PlanningApproverDisplayName + "/ " + _vm.ActivePlanningItem.PlanningItemTreatmentApproverDisplayName;
                string algorithmus = _vm.ActivePlanningItem.Algorithmus;
                string gated = _vm.ActivePlanningItem.PlanningItemPlanSetup.UseGating == true ? "Ja" : "Nein";
                string bolus = _vm.ActivePlanningItem.PlanningItemPlanSetup.Beams.Where(x => x.Boluses.Count() > 0).Any() == true ? "Ja" : "Nein";
                string treatedFx = _vm.ActivePlanningItem.PlanningItemPlanSetup.TreatmentSessions.Where(x => x.Status == TreatmentSessionStatus.Completed).Count() + "/" + _vm.ActivePlanningItem.PlanningItemPlanSetup.NumberOfFractions.ToString();

                // Daten in Dictionary speichern
                Dictionary<string, string> setupData = new Dictionary<string, string>
                {
                { "Patient", patient },
                { "Kurs", courseID },
                { "Plan", planID },
                { "StructureSet", ssID },
                { "Dosis", dose },
                { "Target", target },
                { "RefPoints", refPoints },
                { "TreatedFx",  treatedFx },
                { "Beams#", beamCount.ToString() },
                { "Technik", technik },
                { "Linac", linac },
                { "Failed-PQMs", failedPQM.ToString() },
                { "Next NE", neDate },
                { "Approver", planningApproval },
                { "Algorithmus", algorithmus },
                { "Bolus", bolus },
                { "Gating", gated }
                };
                collectedData.Add(setupData);
            }

            // Wenn `_vm.PlanningItemObject` ein `PlanSum` ist
            else if (_vm.ActivePlanningItem.PlanningItemObject is PlanSum)
            {
                fontsize = 18;
                string patient = _vm.Patient.LastName + ", " + _vm.Patient.FirstName;
                string psumID = _vm.ActivePlanningItem.PlanningItemId;
                int h = 0;
                // Für `PlanSum` iterieren wir durch die Pläne und sammeln deren Daten
                foreach (PlanSetup ps in _vm.ActivePlanningItem.PlanningItemPlans.Reverse())
                {
                    if (h>0)
                    {
                        patient = "";
                        psumID = "";
                    }
                    h++;
                    string ssID = ps.StructureSet.Id;
                    string planID = ps.Id;
                    string courseID = ps.Course.Id;
                    string dose = ps.NumberOfFractions + "*" + Math.Round(ps.DosePerFraction.Dose, 2) + "Gy=" + Math.Round(ps.TotalDose.Dose, 2) + "Gy";
                    string target = ps.TargetVolumeID;
                    string refPoints = string.Join("; ", ps.ReferencePoints.Select(rp => $"{rp.Id} ({Math.Round(rp.TotalDoseLimit.Dose, 2)}Gy)"));
                    int beamCount = ps.Beams.Where(x=>!x.IsSetupField).Count();
                    string technik = ps.Beams.Where(x => !x.IsSetupField).FirstOrDefault() == null ? "" : (ps.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "static") && ps.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "vmat")) ?
                        "Hybrid" : (ps.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "static") ? "Static"
                        : (ps.Beams.Where(x => !x.IsSetupField).Any(b => b.MLCPlanType.ToString().ToLower() == "vmat") ? "VMAT"
                        : ps.Beams.Where(x => !x.IsSetupField).FirstOrDefault().MLCPlanType.ToString()));
                    string linac = ps.Beams.FirstOrDefault() != null ? ps.Beams.First().TreatmentUnit.Id.ToString().Substring(0, 6) : "NaN";
                    //int failedPQM = 3; // Dummy-Wert
                    //string neDate = "+ 5 Tage"; // Dummy-Wert
                    string planningApproval = ps.PlanningApproverDisplayName + "/ " + ps.TreatmentApproverDisplayName;
                    string algorithmus = ps.PhotonCalculationModel.ToString() + (ps.PhotonCalculationOptions.Where(x => x.Key == "CalculationGridSizeInCM").Count() > 0 ? " (" + ps.PhotonCalculationOptions.Where(x => x.Key == "CalculationGridSizeInCM").FirstOrDefault().Value.ToString() + "cm)" : "");
                    string gated = ps.UseGating == true ? "Ja" : "Nein";
                    string bolus = ps.Beams.Where(x => x.Boluses.Count() > 0).Any() == true ? "Ja" : "Nein";
                    string treatedFx = ps.TreatmentSessions.Where(x => x.Status == TreatmentSessionStatus.Completed).Count()+"/"+ps.NumberOfFractions.ToString();

                    // Daten in Dictionary speichern
                    Dictionary<string, string> planSumData = new Dictionary<string, string>
                    {
                    { "Patient", patient },
                    { "Summe", psumID },
                    { "Kurs", courseID },
                    { "Plan", planID },
                    { "StructureSet", ssID },
                    { "Dosis", dose },
                    { "Target", target },
                    { "RefPoints", refPoints },
                    { "TreatedFx",  treatedFx },
                    { "Beams#", beamCount.ToString() },
                    { "Technik", technik },
                    { "Linac", linac },
                    { "Approver", planningApproval },
                    { "Algorithmus", algorithmus },
                    { "Bolus", bolus },
                    { "Gating", gated }
                    };
                    collectedData.Add(planSumData);
                }
            }

            // Display-Funktion aufrufen, um gesammelte Daten anzuzeigen
            DisplayCollectedData(grid, collectedData, charPerLine, fontsize);

            // Das Grid in den ScrollViewer setzen
            scrollViewer.Content = grid;
            infoWindow.Content = scrollViewer;

            // Fenster anzeigen
            infoWindow.ShowDialog();
            /////
        }

        // Funktion, um Daten anzuzeigen (verwendet für beide: PlanSetup und PlanSum)
        // Funktion zum Anzeigen der gesammelten Daten
        void DisplayCollectedData(Grid grid, List<Dictionary<string, string>> collectedData, int charPerLine, int fontsize)
        {
            System.Windows.Media.Brush[] rowColors = { System.Windows.Media.Brushes.White, System.Windows.Media.Brushes.LightGray };

            // Datenzeilen hinzufügen
            for (int i = 0; i < collectedData[0].Count; i++) // Jede Zeile für Parameter (Patient, Plan, usw.)
            {
                RowDefinition rowDef = new RowDefinition { Height = GridLength.Auto };
                grid.RowDefinitions.Add(rowDef);

                // Parameter-Label
                string parameter = collectedData[0].Keys.ElementAt(i);
                Label parameterLabel = new Label
                {
                    Content = parameter,
                    FontSize = fontsize,
                    FontWeight = FontWeights.Bold,
                    //HorizontalAlignment = HorizontalAlignment.Left,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    Foreground = System.Windows.Media.Brushes.Black,
                    Background = rowColors[i % 2], // Alternierende Farben
                    Padding = new Thickness(5)
                };
                Grid.SetRow(parameterLabel, i);
                Grid.SetColumn(parameterLabel, 0);
                grid.Children.Add(parameterLabel);

                // Für jeden Plan die zugehörigen Werte anzeigen
                for (int j = 0; j < collectedData.Count; j++) // Durch alle Pläne iterieren
                {
                    string value = InsertLineBreaks(collectedData[j].Values.ElementAt(i), charPerLine);

                    System.Windows.Media.Brush foregroundColor = System.Windows.Media.Brushes.Black;
                    if(parameter== "Failed-PQMs" && int.Parse(collectedData[j]["Failed-PQMs"]) > 0)
                        foregroundColor = System.Windows.Media.Brushes.Red;

                    Label valueLabel = new Label
                    {
                        Content = value,
                        FontSize = fontsize,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        Foreground = foregroundColor,
                        Background = rowColors[i % 2], // Alternierende Farben
                        Padding = new Thickness(5)
                    };
                    Grid.SetRow(valueLabel, i);
                    Grid.SetColumn(valueLabel, j +1); // Spalten für Werte (ab 1, da 0 Parameter)
                    grid.Children.Add(valueLabel);
                }
            }
        }
       
        // Funktion für den dynamischen Zeilenumbruch
        private string InsertLineBreaks(string input, int columnWidth)
        {
            for (int i = columnWidth; i < input.Length; i += columnWidth)
            {
                input = input.Insert(i, "\n");
                i++; // Ausgleich für den Zeilenumbruch
            }
            return input;
        }
}
}
