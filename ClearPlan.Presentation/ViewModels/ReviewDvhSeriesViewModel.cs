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

        public ReviewDvhSeriesViewModel(ReviewDvhSeries series)
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
