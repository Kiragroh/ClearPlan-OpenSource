using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using ClearPlan.Core.Review;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>Read-only overview capture on the ESAPI owning STA. Only detached pixels and overlays leave this boundary.</summary>
    public static class EsapiPlanImageBuilder
    {
        private static readonly string[] Kinds = { "transversal", "coronal", "sagittal" };
        private static readonly string[] Titles = { "Transversal", "Coronal (orthogonal)", "Sagittal" };
        private const int MaxStructures = 24;
        private const int MaxDoseSamples = 192;
        private const int MaximumCaptureMilliseconds = 45000;
        private const int NativeRowChunkSize = 4;

        public static List<ReviewPlanImage> Build(PlanningItem planningItem, string planKey)
        {
            // Every await completes inline when yielding is disabled, including headless STA exports.
            return BuildCoreAsync(planningItem, planKey, new CaptureContext(false, CancellationToken.None, () => true)).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Start on the owning WPF dispatcher while the patient remains open. The host callback must
        /// inspect detached host state only and become false before closing or replacing the plan.
        /// Cancellation and context replacement propagate; never publish the partially captured list.
        /// </summary>
        public static Task<List<ReviewPlanImage>> BuildAsync(PlanningItem planningItem, string planKey,
            CancellationToken cancellation, Func<bool> isCurrent)
        {
            if (isCurrent == null) throw new ArgumentNullException("isCurrent");
            return BuildCoreAsync(planningItem, planKey, new CaptureContext(true, cancellation, isCurrent));
        }

        private static async Task<List<ReviewPlanImage>> BuildCoreAsync(PlanningItem planningItem, string planKey, CaptureContext context)
        {
            context.RequireCurrent();
            var plan = planningItem as ExternalPlanSetup;
            if (plan == null) return Unavailable(planKey, "CT overview requires a single external-beam plan, not a plan sum.");
            try
            {
                var captureClock = Stopwatch.StartNew();
                var image = plan.StructureSet == null ? null : plan.StructureSet.Image;
                if (image == null) return Unavailable(planKey, "The active plan has no planning image.");
                if (!string.Equals(image.DisplayUnit, "HU", StringComparison.OrdinalIgnoreCase))
                    return Unavailable(planKey, "The planning image does not expose HU display values; no CT image was substituted.");
                var beams = plan.Beams.Where(beam => !beam.IsSetupField).ToList();
                if (beams.Count == 0) return Unavailable(planKey, "The plan has no treatment-beam isocenter.");
                VVector iso = beams[0].IsocenterPosition;
                if (beams.Any(beam => Distance(beam.IsocenterPosition, iso) > 0.1))
                    return Unavailable(planKey, "Multiple treatment isocenters: a single CT overview center is ambiguous.");

                int sx = OrthogonalImageGeometry.AxisSign(image.XDirection.x, image.XDirection.y, image.XDirection.z, 0);
                int sy = OrthogonalImageGeometry.AxisSign(image.YDirection.x, image.YDirection.y, image.YDirection.z, 1);
                int sz = OrthogonalImageGeometry.AxisSign(image.ZDirection.x, image.ZDirection.y, image.ZDirection.z, 2);
                int nx = image.XSize, ny = image.YSize, nz = image.ZSize;
                if (nx < 2 || ny < 2 || nz < 2 || nx > 2048 || ny > 2048 || nz > 4096)
                    return Unavailable(planKey, "Image grid dimensions are outside the supported volumetric overview limits.");
                int ix = OrthogonalImageGeometry.NearestIndex(iso.x, image.Origin.x, image.XRes, sx, nx);
                int iy = OrthogonalImageGeometry.NearestIndex(iso.y, image.Origin.y, image.YRes, sy, ny);
                int iz = OrthogonalImageGeometry.NearestIndex(iso.z, image.Origin.z, image.ZRes, sz, nz);
                int wx = Math.Min(nx, 512), wy = Math.Min(ny, 512), wz = Math.Min(nz, 512);
                double dx = image.XRes * (nx - 1) / (wx - 1), dy = image.YRes * (ny - 1) / (wy - 1), dz = image.ZRes * (nz - 1) / (wz - 1);
                double px = (sx > 0 ? ix : nx - 1 - ix) * (wx - 1.0) / (nx - 1);
                double py = (sy > 0 ? iy : ny - 1 - iy) * (wy - 1.0) / (ny - 1);
                double pz = (sz < 0 ? iz : nz - 1 - iz) * (wz - 1.0) / (nz - 1);
                string caption = string.Format(CultureInfo.InvariantCulture,
                    "Planning CT; nearest planes to treatment isocenter. W400 / L40 HU; nearest-neighbor overview. Crosshair is the projected treatment isocenter. Plane offsets x/y/z: {0:0.##}/{1:0.##}/{2:0.##} mm.",
                    image.Origin.x + ix * image.XRes * sx - iso.x,
                    image.Origin.y + iy * image.YRes * sy - iso.y,
                    image.Origin.z + iz * image.ZRes * sz - iso.z);
                var result = new List<ReviewPlanImage>
                {
                    Make(planKey, 0, wx, wy, dx, dy, px, py, caption),
                    Make(planKey, 1, wx, wz, dx, dz, px, pz, caption),
                    Make(planKey, 2, wy, wz, dy, dz, py, pz, caption)
                };
                // Vendor conversion is applied for each distinct raw voxel value; no assumed HU slope/intercept.
                var lookup = new Dictionary<int, byte>();
                Func<int, byte> gray = raw =>
                {
                    byte value;
                    if (!lookup.TryGetValue(raw, out value))
                    {
                        if (lookup.Count >= 65536) throw new ArgumentException("Native HU conversion limit exceeded.");
                        value = OrthogonalImageGeometry.WindowHu(image.VoxelToDisplayValue(raw), 40, 400);
                        lookup.Add(raw, value);
                    }
                    return value;
                };
                var buffer = new int[nx, ny];
                await context.YieldAsync();
                context.RequireCurrent();
                image.GetVoxels(iz, buffer);
                for (int row = 0; row < wy; row++)
                {
                    if (row % NativeRowChunkSize == 0)
                    {
                        await context.YieldAsync();
                        context.RequireCurrent();
                    }
                    context.RequireCurrent();
                    RequireTime(captureClock, MaximumCaptureMilliseconds);
                    for (int column = 0; column < wx; column++)
                        result[0].GrayscalePixels[row * wx + column] = gray(buffer[Index(column, wx, nx, sx), Index(row, wy, ny, sy)]);
                }
                for (int row = 0; row < wz; row++)
                {
                    if (row % NativeRowChunkSize == 0)
                    {
                        await context.YieldAsync();
                        context.RequireCurrent();
                    }
                    context.RequireCurrent();
                    RequireTime(captureClock, MaximumCaptureMilliseconds);
                    image.GetVoxels(Index(row, wz, nz, -sz), buffer);
                    for (int column = 0; column < wx; column++)
                        result[1].GrayscalePixels[row * wx + column] = gray(buffer[Index(column, wx, nx, sx), iy]);
                    for (int column = 0; column < wy; column++)
                        result[2].GrayscalePixels[row * wy + column] = gray(buffer[ix, Index(column, wy, ny, sy)]);
                }
                double minX = Math.Min(image.Origin.x, image.Origin.x + (nx - 1) * image.XRes * sx);
                double minY = Math.Min(image.Origin.y, image.Origin.y + (ny - 1) * image.YRes * sy);
                double maxZ = Math.Max(image.Origin.z, image.Origin.z + (nz - 1) * image.ZRes * sz);
                var planes = new[] {
                    new PlanImagePlane(Kinds[0], wx, wy, minX, minY, image.Origin.z + iz * image.ZRes * sz, dx, dy),
                    new PlanImagePlane(Kinds[1], wx, wz, minX, image.Origin.y + iy * image.YRes * sy, maxZ, dx, dz),
                    new PlanImagePlane(Kinds[2], wy, wz, image.Origin.x + ix * image.XRes * sx, minY, maxZ, dy, dz) };
                for (int view = 0; view < result.Count; view++)
                {
                    var projectedIso = planes[view].ToPixel(new[] { iso.x, iso.y, iso.z });
                    result[view].IsocenterPixelX = projectedIso.X; result[view].IsocenterPixelY = projectedIso.Y;
                }
                await CaptureStructuresAsync(plan.StructureSet, iz, result, planes, context);
                context.RequireCurrent();
                await CaptureDoseAsync(plan, result, planes, context);
                context.RequireCurrent();
                return result;
            }
            catch (TimeoutException)
            {
                return Unavailable(planKey, "Native CT capture exceeded its bounded overview time budget; no partial or substitute anatomy was retained.");
            }
            catch (ArgumentException)
            {
                return Unavailable(planKey, "CT geometry is unsupported, invalid, or the treatment isocenter lies outside the image grid.");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception)
            {
                // Do not expose vendor exception text, which may contain patient identifiers or paths.
                return Unavailable(planKey, "The ESAPI planning-image read failed; no substitute anatomy was generated.");
            }
        }

        private static async Task CaptureStructuresAsync(StructureSet structureSet, int transverseIndex,
            List<ReviewPlanImage> images, PlanImagePlane[] planes, CaptureContext context)
        {
            var clock = Stopwatch.StartNew();
            try
            {
                context.RequireCurrent();
                // A display subset, never an inferred clinical target/constraint assignment. All getters stay on this thread.
                var candidates = structureSet.Structures.Take(257).ToList();
                bool inventoryLimited = candidates.Count > 256;
                var structures = candidates.Take(256).Where(s => !s.IsEmpty && s.HasSegment)
                    .OrderBy(s => StructureRank(s.Id, s.DicomType)).ThenBy(s => s.Id, StringComparer.OrdinalIgnoreCase).Take(MaxStructures + 1).ToList();
                bool subset = structures.Count > MaxStructures || inventoryLimited;
                foreach (var structure in structures.Take(MaxStructures))
                {
                    await context.YieldAsync();
                    context.RequireCurrent();
                    RequireTime(clock, 20000);
                    string label = structure.Id;
                    var color = structure.Color;
                    string colorHex = string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
                    try
                    {
                        RequireTime(clock, 20000);
                        var native = structure.GetContoursOnImagePlane(transverseIndex);
                        if (native == null) throw new ArgumentException("Native contours unavailable.");
                        if (native.Sum(c => c == null ? 0L : c.LongLength) > PlanImageOverlayGeometry.MaxMeshVertices)
                            throw new ArgumentException("Native contour capture limit exceeded.");
                        var detached = native.Where(c => c != null).Select(c => c.Select(p => new[] { p.x, p.y, p.z }).ToArray()).ToArray();
                        images[0].Overlays.Add(AvailableOverlay("structure", label, colorHex, "ESAPI Structure.GetContoursOnImagePlane; native CT plane", PlanImageOverlayGeometry.ProjectContours(planes[0], detached)));
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception)
                    {
                        images[0].Overlays.Add(UnavailableOverlay("structure", label, "Native transverse contour unavailable, invalid, or exceeds capture limits.", colorHex));
                    }
                    try
                    {
                        await context.YieldAsync();
                        context.RequireCurrent();
                        RequireTime(clock, 20000);
                        var mesh = structure.MeshGeometry;
                        if (mesh == null || mesh.Positions == null || mesh.TriangleIndices == null || mesh.Positions.Count == 0 || mesh.TriangleIndices.Count == 0 ||
                            mesh.Positions.Count > PlanImageOverlayGeometry.MaxMeshVertices || mesh.TriangleIndices.Count / 3 > PlanImageOverlayGeometry.MaxMeshTriangles)
                            throw new ArgumentException("Native mesh unavailable or exceeds capture limits.");
                        // Copy WPF mesh primitives here; no vendor/WPF object survives the capture boundary.
                        var positions = mesh.Positions.Select(p => new[] { p.X, p.Y, p.Z }).ToArray();
                        var indices = mesh.TriangleIndices.ToArray();
                        for (int view = 1; view < 3; view++)
                        {
                            await context.YieldAsync();
                            context.RequireCurrent();
                            RequireTime(clock, 20000);
                            var paths = PlanImageOverlayGeometry.IntersectMesh(planes[view], positions, indices);
                            images[view].Overlays.Add(AvailableOverlay("structure", label, colorHex, "ESAPI Structure.MeshGeometry; triangle-plane intersection (native mesh approximation)", paths));
                        }
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception)
                    {
                        for (int view = 1; view < 3; view++)
                            if (!images[view].Overlays.Any(o => o.Kind == "structure" && o.Label == label))
                                images[view].Overlays.Add(UnavailableOverlay("structure", label, "Native mesh cross-section unavailable, invalid, or exceeds capture limits.", colorHex));
                    }
                }
                foreach (var image in images)
                    image.OverlaySummary = (subset ? "Limited structure display subset; " : "Structure display; ") +
                        "up to 24 nonempty segmented structures, targets first. Empty paths mean no intersection with this plane; unavailable sources are listed explicitly.";
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception)
            {
                foreach (var image in images)
                {
                    image.OverlaySummary = "Structure capture incomplete: source read failed or the 20 s cooperative budget was reached. Previously completed structures are retained.";
                    image.Overlays.Add(UnavailableOverlay("structure", "Additional structures", "Native structure capture incomplete; no replacement contours generated.", "#AAB4C0"));
                }
            }
        }

        private static async Task CaptureDoseAsync(ExternalPlanSetup plan, List<ReviewPlanImage> images,
            PlanImagePlane[] planes, CaptureContext context)
        {
            var clock = Stopwatch.StartNew();
            try
            {
                await context.YieldAsync();
                context.RequireCurrent();
                if (!plan.IsDoseValid) { AddDoseUnavailable(images, "Native IsDoseValid is false; cached or invalid dose is withheld."); return; }
                PlanImageDoseGrid doseGrid = null;
                double prescriptionGy = 0;
                var dose = WithAbsoluteDosePresentation(plan, context, () =>
                {
                    var nativeDose = plan.Dose;
                    if (nativeDose == null) return null;
                    var doseOrigin = nativeDose.Origin; var doseX = nativeDose.XDirection; var doseY = nativeDose.YDirection; var doseZ = nativeDose.ZDirection;
                    doseGrid = new PlanImageDoseGrid(new[] { doseOrigin.x, doseOrigin.y, doseOrigin.z },
                        new[] { new[] { doseX.x, doseX.y, doseX.z }, new[] { doseY.x, doseY.y, doseY.z }, new[] { doseZ.x, doseZ.y, doseZ.z } },
                        new[] { nativeDose.XRes, nativeDose.YRes, nativeDose.ZRes }, new[] { nativeDose.XSize, nativeDose.YSize, nativeDose.ZSize });
                    prescriptionGy = PlanImageOverlayGeometry.ToGy(plan.TotalDose.Dose, plan.TotalDose.Unit.ToString());
                    return nativeDose;
                });
                if (dose == null) { AddDoseUnavailable(images, "No native total plan dose is available."); return; }
                if (prescriptionGy <= 0) { AddDoseUnavailable(images, "No positive absolute prescription is available for isodose display levels."); return; }
                for (int view = 0; view < images.Count; view++)
                {
                    await context.YieldAsync();
                    context.RequireCurrent();
                    var image = images[view]; var plane = planes[view];
                    string captureStep = "prepare plane";
                    try
                    {
                        int columns = Math.Min(MaxDoseSamples, image.WidthPixels), rows = Math.Min(MaxDoseSamples, image.HeightPixels);
                        double[] samples = new double[columns * rows];
                        var profileBuffer = new double[columns];
                        for (int firstRow = 0; firstRow < rows; firstRow += NativeRowChunkSize)
                        {
                            await context.YieldAsync();
                            context.RequireCurrent();
                            WithAbsoluteDosePresentation(plan, context, () =>
                            {
                                for (int row = firstRow; row < Math.Min(rows, firstRow + NativeRowChunkSize); row++)
                                {
                                    context.RequireCurrent();
                                    RequireTime(clock, 20000);
                                    captureStep = "native row " + row;
                                    double pixelY = row * (image.HeightPixels - 1.0) / (rows - 1);
                                    double[] start = plane.ToWorld(0, pixelY), end = plane.ToWorld(image.WidthPixels - 1.0, pixelY);
                                    // Native interpolation: one profile per row, at most 192 x 3 calls, no per-pixel ESAPI calls.
                                    var profile = dose.GetDoseProfile(new VVector(start[0], start[1], start[2]), new VVector(end[0], end[1], end[2]), profileBuffer);
                                    if (profile == null || profile.Count != columns) throw new ArgumentException("Unexpected native profile length.");
                                    string unit = profile.Unit.ToString();
                                    captureStep = "dose unit " + unit;
                                    double unitScale = PlanImageOverlayGeometry.ToGy(1, unit);
                                    for (int column = 0; column < columns; column++)
                                    {
                                        var sample = profile[column];
                                        captureStep = "profile position row " + row + " column " + column + " unit " + unit;
                                        double[] expected = plane.ToWorld(column * (image.WidthPixels - 1.0) / (columns - 1), pixelY);
                                        if (Distance(sample.Position, new VVector(expected[0], expected[1], expected[2])) > 0.01)
                                            throw new ArgumentException("Native profile sample coordinates do not match the CT plane.");
                                        double value = sample.Value;
                                        samples[row * columns + column] = !doseGrid.Contains(expected) || double.IsNaN(value) || double.IsInfinity(value) || value < 0 ? double.NaN : value * unitScale;
                                    }
                                }
                                return true;
                            });
                        }
                        RequireTime(clock, 20000);
                        captureStep = "finite dose coverage";
                        if (!samples.Any(v => !double.IsNaN(v))) throw new ArgumentException("No finite native dose in this CT plane.");
                        image.DoseFocusRegion = PlanImageViewport.CaptureDoseRegion(samples, columns, rows,
                            prescriptionGy * 0.02, image.WidthPixels, image.HeightPixels);
                        captureStep = "isodose tracing";
                        image.DosePlane = new ReviewImageDosePlane { Columns = columns, Rows = rows, PrescriptionGy = prescriptionGy,
                            SamplesGy = samples, Source = "ESAPI total-plan Dose.GetDoseProfile; native unit explicitly converted to Gy; " + columns + " x " + rows +
                            " native interpolated samples; marching squares; display levels relative to prescribed total dose, not clinical constraints." };
                        await context.YieldAsync(); context.RequireCurrent();
                        image.Overlays = IsodoseDisplayConfiguration.CreateDefault().Apply(image).Overlays;
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception)
                    {
                        image.Overlays.Add(UnavailableOverlay("isodose", "Isodoses", "Native dose capture unavailable at " + captureStep + ". No substitute dose was rendered.", "#AAB4C0"));
                    }
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception)
            {
                AddDoseUnavailable(images, "Native dose validity, absolute prescription, or dose source could not be verified; no substitute dose was rendered.");
            }
        }

        private static T WithAbsoluteDosePresentation<T>(ExternalPlanSetup plan, CaptureContext context, Func<T> read)
        {
            context.RequireCurrent();
            DoseValuePresentation? previousPresentation = null;
            try
            {
                // Presentation is only a read-context setting. The complete scope is synchronous:
                // restore before the dispatcher can service a different plan/patient selection.
                previousPresentation = plan.DoseValuePresentation;
                plan.DoseValuePresentation = DoseValuePresentation.Absolute;
                return read();
            }
            finally
            {
                context.VerifyOwner();
                // Cancellation alone still restores presentation; a closed/replaced native context must never be touched.
                if (previousPresentation.HasValue && context.IsCurrent)
                    plan.DoseValuePresentation = previousPresentation.Value;
            }
        }

        private static int StructureRank(string structureId, string dicomType)
        {
            string target = DvhSelectionPolicy.ClassifyTarget(structureId, dicomType);
            return target == "PTV" ? 0 : target == "CTV" ? 1 : target == "GTV" || target == "ITV" ? 2 :
                string.Equals(dicomType, "EXTERNAL", StringComparison.OrdinalIgnoreCase) ? 4 : 3;
        }

        private sealed class CaptureContext
        {
            private readonly bool yieldToDispatcher;
            private readonly CancellationToken cancellation;
            private readonly Func<bool> isCurrent;
            private readonly int ownerThreadId;
            private readonly Dispatcher owner;

            internal CaptureContext(bool yieldToDispatcher, CancellationToken cancellation, Func<bool> isCurrent)
            {
                this.yieldToDispatcher = yieldToDispatcher;
                this.cancellation = cancellation;
                this.isCurrent = isCurrent;
                ownerThreadId = Thread.CurrentThread.ManagedThreadId;
                owner = Dispatcher.FromThread(Thread.CurrentThread);
                VerifyOwner();
                if (yieldToDispatcher && (owner == null || !(SynchronizationContext.Current is DispatcherSynchronizationContext)))
                    throw new InvalidOperationException("Responsive native CT capture requires the ESAPI owning WPF dispatcher context.");
                RequireCurrent();
            }

            internal bool IsCurrent { get { return isCurrent(); } }

            internal void VerifyOwner()
            {
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || Thread.CurrentThread.ManagedThreadId != ownerThreadId)
                    throw new InvalidOperationException("Native CT capture requires the ESAPI owning STA thread.");
                if (owner != null) owner.VerifyAccess();
                if (System.Windows.Application.Current != null) System.Windows.Application.Current.Dispatcher.VerifyAccess();
            }

            internal void RequireCurrent()
            {
                VerifyOwner();
                cancellation.ThrowIfCancellationRequested();
                if (!isCurrent()) throw new OperationCanceledException("Native CT capture context was replaced.", cancellation);
            }

            internal async Task YieldAsync()
            {
                RequireCurrent();
                if (!yieldToDispatcher) return;
                await Dispatcher.Yield(DispatcherPriority.Background);
                RequireCurrent();
            }
        }

        private static ReviewImageOverlay AvailableOverlay(string kind, string label, string color, string source, List<ReviewImagePath> paths)
        {
            return new ReviewImageOverlay { Kind = kind, Label = label, ColorHex = color, Source = source,
                SourceStatus = ReviewStatusCodes.Available, Paths = paths };
        }

        private static ReviewImageOverlay UnavailableOverlay(string kind, string label, string reason, string color)
        {
            return new ReviewImageOverlay { Kind = kind, Label = label, ColorHex = color, Source = "Native ESAPI planning data",
                SourceStatus = ReviewStatusCodes.Unavailable, UnavailableReason = reason };
        }

        private static void AddDoseUnavailable(IEnumerable<ReviewPlanImage> images, string reason)
        {
            foreach (var image in images) image.Overlays.Add(UnavailableOverlay("isodose", "Isodoses", reason, "#AAB4C0"));
        }

        private static void RequireTime(Stopwatch clock, int milliseconds)
        {
            // ESAPI calls cannot safely be interrupted: this is a cooperative bound between native calls, not a hard timeout.
            if (clock.ElapsedMilliseconds > milliseconds) throw new TimeoutException("Native overview capture time budget exceeded.");
        }

        private static int Index(int output, int outputSize, int inputSize, int direction)
        {
            int index = (int)Math.Round(output * (inputSize - 1.0) / (outputSize - 1));
            return direction > 0 ? index : inputSize - 1 - index;
        }

        private static double Distance(VVector a, VVector b)
        {
            return Math.Sqrt((a.x - b.x) * (a.x - b.x) + (a.y - b.y) * (a.y - b.y) + (a.z - b.z) * (a.z - b.z));
        }

        private static ReviewPlanImage Make(string key, int view, int width, int height, double dx, double dy, double px, double py, string caption)
        {
            return new ReviewPlanImage
            {
                PlanKey = key, Kind = Kinds[view], Title = Titles[view], Caption = caption,
                SourceStatus = ReviewStatusCodes.Available, WidthPixels = width, HeightPixels = height,
                PixelSpacingXMillimeters = dx, PixelSpacingYMillimeters = dy,
                GrayscalePixels = new byte[width * height], IsocenterPixelX = px, IsocenterPixelY = py,
                LeftOrientation = view == 2 ? "A" : "R", RightOrientation = view == 2 ? "P" : "L",
                TopOrientation = view == 0 ? "A" : "S", BottomOrientation = view == 0 ? "P" : "I"
            };
        }

        private static List<ReviewPlanImage> Unavailable(string key, string reason)
        {
            return Enumerable.Range(0, 3).Select(index => new ReviewPlanImage
            {
                PlanKey = key, Kind = Kinds[index], Title = Titles[index],
                SourceStatus = ReviewStatusCodes.Unavailable, UnavailableReason = reason
            }).ToList();
        }
    }
}
