using System;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace ClearPlan.Presentation.Views
{
    // Optional local display asset. Deliberately never exposed to collision evaluators.
    internal static class DioramaDisplayModel
    {
        private static Model3DGroup fixedPart, movingPart;
        private static bool attempted;
        internal static bool Available { get { return fixedPart != null && movingPart != null; } }
        internal static bool IsExplicitRingDisplay(string machineId)
        {
            if(string.IsNullOrWhiteSpace(machineId))return false;
            string path=Path.Combine(Path.GetDirectoryName(typeof(DioramaDisplayModel).Assembly.Location),"VisualModels","machine-display.ini");
            try
            {
                if(!File.Exists(path))return false;
                foreach(string line in File.ReadLines(path))
                {
                    string text=line.Trim();if(text.StartsWith("#")||text.StartsWith(";"))continue;
                    int split=text.IndexOf('=');if(split<1)continue;
                    if(string.Equals(text.Substring(0,split).Trim(),machineId,StringComparison.OrdinalIgnoreCase))
                        return text.Substring(split+1).Trim()=="Ring";
                }
            }
            catch(IOException) { }
            catch(UnauthorizedAccessException) { }
            return false;
        }
        internal static Model3DGroup Create(double gantryDegrees)
        {
            if (!attempted)
            {
                attempted = true;
                string file = Path.Combine(Path.GetDirectoryName(typeof(DioramaDisplayModel).Assembly.Location), "VisualModels", "truebeam-display.bin");
                if (File.Exists(file))
                {
                    try { Load(file); }
                    catch (IOException) { fixedPart = movingPart = null; }
                    catch (ArgumentException) { fixedPart = movingPart = null; }
                    catch (UnauthorizedAccessException) { fixedPart = movingPart = null; }
                }
            }
            if (fixedPart == null || movingPart == null) return null;
            var group = new Model3DGroup(); group.Children.Add(fixedPart);
            var moving = movingPart.Clone();
            moving.Transform = new RotateTransform3D(new AxisAngleRotation3D(new Vector3D(0,0,1),gantryDegrees));
            group.Children.Add(moving); return group;
        }
        private static void Load(string path)
        {
            using (var reader = new BinaryReader(File.OpenRead(path)))
            {
                if (reader.ReadInt32() != 0x43504431) throw new InvalidDataException("Display model format");
                int count = reader.ReadInt32(); if (count < 1 || count > 100) throw new InvalidDataException("Display mesh count");
                var stand = new Model3DGroup(); var gantry = new Model3DGroup();
                for (int m=0;m<count;m++)
                {
                    bool moving=reader.ReadInt32()==1;
                    int r=reader.ReadInt32(),g=reader.ReadInt32(),b=reader.ReadInt32();
                    if(r<0||r>255||g<0||g>255||b<0||b>255)throw new InvalidDataException("Display color");
                    var mesh=new MeshGeometry3D(); int n=reader.ReadInt32();
                    if(n<3||n>200000)throw new InvalidDataException("Display vertex count");
                    for(int k=0;k<n;k++) {double x=reader.ReadDouble(),y=reader.ReadDouble(),z=reader.ReadDouble();
                        if(double.IsNaN(x+y+z)||double.IsInfinity(x+y+z)||Math.Max(Math.Abs(x),Math.Max(Math.Abs(y),Math.Abs(z)))>10000)throw new InvalidDataException("Display vertex");
                        mesh.Positions.Add(new Point3D(x,y,z));}
                    int indices=reader.ReadInt32(); if(indices<3||indices>1000000||indices%3!=0)throw new InvalidDataException("Display index count");
                    for(int k=0;k<indices;k++){int index=reader.ReadInt32();if(index<0||index>=n)throw new InvalidDataException("Display index");mesh.TriangleIndices.Add(index);}
                    var material=new DiffuseMaterial(new SolidColorBrush(Color.FromRgb((byte)r,(byte)g,(byte)b)));
                    (moving?gantry:stand).Children.Add(new GeometryModel3D(mesh,material){BackMaterial=material});
                }
                if(reader.BaseStream.Position!=reader.BaseStream.Length)throw new InvalidDataException("Trailing display data");
                stand.Freeze();gantry.Freeze();fixedPart=stand;movingPart=gantry;
            }
        }
    }
}
