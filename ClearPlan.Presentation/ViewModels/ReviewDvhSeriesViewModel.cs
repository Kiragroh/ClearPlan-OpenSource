using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using ClearPlan.Core.Review;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed class ReviewDvhSeriesViewModel : INotifyPropertyChanged
    {
        private bool isSelected;

        public ReviewDvhSeriesViewModel(ReviewDvhSeries series) : this(series, false) { }

        public ReviewDvhSeriesViewModel(ReviewDvhSeries series, bool synthetic)
        {
            if (series == null)
            {
                throw new ArgumentNullException("series");
            }

            StableId = series.StableId;
            StructureId = series.StructureId;
            DisplayName = series.DisplayName;
            Role = series.Role;
            ColorHex = series.ColorHex;
            LineStyle = series.LineStyle;
            VolumeCc = series.VolumeCc;
            VolumeText = series.VolumeCc.HasValue
                ? series.VolumeCc.Value.ToString(
                    "0.0",
                    CultureInfo.InvariantCulture) + " cm³"
                : "Volumen n/a";
            Points = new List<ReviewDvhPoint>(
                series.Points ?? new List<ReviewDvhPoint>());
            RequiredForTargetReview = series.RequiredForTargetReview;
            // Display-only summary; same calculator and rule as the PDF/HTML report.
            Statistics = ReviewDvhStatisticsCalculator.ForSeries(series, synthetic);
            MinimumText = Dose(Statistics.MinimumDoseGy, Statistics.MinimumDoseEstimated);
            D98Text = Dose(Statistics.D98Gy, false);
            D50Text = Dose(Statistics.MedianDoseGy, false);
            MeanText = Dose(Statistics.MeanDoseGy, Statistics.MeanDoseEstimated);
            D2Text = Dose(Statistics.D2Gy, false);
            MaximumText = Dose(Statistics.MaximumDoseGy, Statistics.MaximumDoseEstimated);
            VolumeCellText = series.VolumeCc.HasValue
                ? series.VolumeCc.Value.ToString("0.0", CultureInfo.InvariantCulture) : "—";
            HoverStatisticsText = "V " + (series.VolumeCc.HasValue ? VolumeCellText + " cm³" : "—") +
                " · Dmin " + MinimumText + " · D98 " + D98Text + " · D50 " + D50Text +
                "\nDmean " + MeanText + " · D2 " + D2Text + " · Dmax " + MaximumText + " Gy";
            isSelected = series.Selected || RequiredForTargetReview;
            IsInitiallySelected = isSelected;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public string StableId { get; private set; }

        public string StructureId { get; private set; }

        public string DisplayName { get; private set; }

        public string Role { get; private set; }

        public string ColorHex { get; private set; }

        public string LineStyle { get; private set; }

        public double? VolumeCc { get; private set; }

        public string VolumeText { get; private set; }

        public ReviewDvhStatistics Statistics { get; private set; }
        public string VolumeCellText { get; private set; }
        public string MinimumText { get; private set; }
        public string D98Text { get; private set; }
        public string D50Text { get; private set; }
        public string MeanText { get; private set; }
        public string D2Text { get; private set; }
        public string MaximumText { get; private set; }
        /// <summary>Appended to the DVH hover text; no format placeholders.</summary>
        public string HoverStatisticsText { get; private set; }

        private static string Dose(double? value, bool estimated)
        {
            if (!value.HasValue) return "—";
            return (estimated ? "≈" : "") + value.Value.ToString("0.00", CultureInfo.InvariantCulture);
        }

        public IList<ReviewDvhPoint> Points { get; private set; }

        public bool IsInitiallySelected { get; private set; }

        public bool RequiredForTargetReview { get; private set; }

        public bool IsSelected
        {
            get { return isSelected; }
            set
            {
                // Targets remain in the captured data and default selection, but an
                // explicit display choice may hide them without altering evaluation.
                if (isSelected == value)
                {
                    return;
                }

                isSelected = value;
                OnPropertyChanged("IsSelected");
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
