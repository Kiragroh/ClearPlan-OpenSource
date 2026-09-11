using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Presentation.ViewModels;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;

namespace ClearPlan.Presentation.Plot
{
    public static class ReviewPlotFactory
    {
        public static PlotController CreateHoverController()
        {
            var controller = new PlotController();
            controller.BindMouseEnter(PlotCommands.HoverSnapTrack);
            return controller;
        }

        public static PlotModel Create(
            IEnumerable<ReviewDvhSeriesViewModel> source)
        {
            IList<ReviewDvhSeriesViewModel> seriesRows =
                (source ?? Enumerable.Empty<ReviewDvhSeriesViewModel>())
                    .ToList();
            double maximumDose = seriesRows
                .SelectMany(row => row.Points)
                .Select(point => point.DoseGy)
                .DefaultIfEmpty(1.0)
                .Max();

            var model = new PlotModel
            {
                Background = OxyColors.White,
                PlotAreaBackground = OxyColor.Parse("#FBFCFE"),
                PlotAreaBorderColor = OxyColor.Parse("#CBD5E1"),
                PlotAreaBorderThickness = new OxyThickness(1),
                IsLegendVisible = true,
                LegendPlacement = LegendPlacement.Outside,
                LegendPosition = LegendPosition.RightTop,
                LegendOrientation = LegendOrientation.Vertical,
                LegendMaxWidth = 220,
                LegendBackground = OxyColor.FromAColor(
                    224,
                    OxyColors.White),
                LegendBorder = OxyColor.Parse("#CBD5E1"),
                LegendTextColor = OxyColor.Parse("#334155")
            };

            model.Axes.Add(
                new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Dosis [Gy]",
                    Minimum = 0,
                    Maximum = Math.Max(1.0, maximumDose * 1.05),
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    MajorGridlineColor = OxyColor.Parse("#D8E1EB"),
                    MinorGridlineColor = OxyColor.Parse("#EEF2F7")
                });
            model.Axes.Add(
                new LinearAxis
                {
                    Position = AxisPosition.Left,
                    Title = "Relatives Volumen [%]",
                    Minimum = 0,
                    Maximum = 100,
                    AbsoluteMinimum = 0,
                    AbsoluteMaximum = 100,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    MajorGridlineColor = OxyColor.Parse("#D8E1EB"),
                    MinorGridlineColor = OxyColor.Parse("#EEF2F7")
                });

            foreach (ReviewDvhSeriesViewModel row in seriesRows)
            {
                var line = new LineSeries
                {
                    Tag = row.StableId,
                    Title = row.DisplayName,
                    Color = ParseColor(row.ColorHex),
                    LineStyle = ParseLineStyle(row.LineStyle),
                    StrokeThickness = 2.2,
                    IsVisible = row.IsSelected,
                    CanTrackerInterpolatePoints = true,
                    TrackerFormatString =
                        "{0}\nDosis: {2:0.00} Gy\nVolumen: {4:0.00} %"
                };
                foreach (Core.Review.ReviewDvhPoint point in row.Points)
                {
                    line.Points.Add(
                        new DataPoint(point.DoseGy, point.VolumePercent));
                }

                model.Series.Add(line);
            }

            return model;
        }

        public static void SetSeriesVisibility(
            PlotModel model,
            string stableId,
            bool isVisible)
        {
            if (model == null)
            {
                return;
            }

            Series series = model.Series.FirstOrDefault(
                candidate =>
                    string.Equals(
                        Convert.ToString(candidate.Tag),
                        stableId,
                        StringComparison.Ordinal));
            if (series != null)
            {
                series.IsVisible = isVisible;
                model.InvalidatePlot(false);
            }
        }

        public static void SetSeriesVisibilities(
            PlotModel model,
            IEnumerable<ReviewDvhSeriesViewModel> source,
            bool resetAxes)
        {
            var selected = source.ToDictionary(row => row.StableId, row => row.IsSelected,
                StringComparer.Ordinal);
            foreach (Series series in model.Series)
            {
                bool isVisible;
                if (selected.TryGetValue(Convert.ToString(series.Tag), out isVisible))
                    series.IsVisible = isVisible;
            }
            if (resetAxes)
                model.ResetAllAxes();
            // Only visibility/viewport changed; the detached DVH points stay intact.
            model.InvalidatePlot(false);
        }

        private static OxyColor ParseColor(string colorHex)
        {
            if (string.IsNullOrWhiteSpace(colorHex))
            {
                return OxyColor.Parse("#0F766E");
            }

            try
            {
                return OxyColor.Parse(colorHex);
            }
            catch (FormatException)
            {
                return OxyColor.Parse("#0F766E");
            }
        }

        private static LineStyle ParseLineStyle(string lineStyle)
        {
            switch (lineStyle)
            {
                case Core.Review.ReviewLineStyleCodes.Dash:
                    return LineStyle.Dash;
                case Core.Review.ReviewLineStyleCodes.Dot:
                    return LineStyle.Dot;
                default:
                    return LineStyle.Solid;
            }
        }
    }
}
