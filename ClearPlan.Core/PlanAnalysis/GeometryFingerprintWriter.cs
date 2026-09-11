using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ClearPlan.Core.PlanAnalysis
{
    /// <summary>Bounded, typed in-memory encoding for native-state equality; no geometry interpretation or file IO.</summary>
    public sealed class GeometryFingerprintWriter : IDisposable
    {
        private const int MaximumBytes=16*1024*1024;
        private readonly MemoryStream buffer=new MemoryStream();
        private readonly BinaryWriter writer;
        private bool finished;

        public GeometryFingerprintWriter()
        {
            writer=new BinaryWriter(buffer,Encoding.UTF8);
            writer.Write("ClearPlan.NativeGeometry.v1");
        }

        public void Add(int value) { Guard(5);writer.Write((byte)1);writer.Write(value); }
        public void Add(double value)
        {
            RequireFinite(value);Guard(9);writer.Write((byte)2);writer.Write(value==0 ? 0.0 : value);
        }
        public void Add(string value)
        {
            if(value!=null && value.Length>256) throw new ArgumentException("Fingerprint label exceeds its bound.");
            Guard(1030);writer.Write((byte)3);writer.Write(value!=null);
            if(value!=null) writer.Write(value);
        }
        public void Add(float[,] value)
        {
            if(value!=null && (value.GetLength(0)>8 || value.GetLength(1)>512))
                throw new ArgumentException("Native array dimensions exceed the fingerprint bound.");
            Guard(value==null ? 2 : checked(10+value.Length*4));
            writer.Write((byte)4);writer.Write(value!=null);
            if(value==null) return;
            writer.Write(value.GetLength(0));writer.Write(value.GetLength(1));
            foreach(float sample in value)
            {
                RequireFinite(sample);
                writer.Write(sample==0 ? 0.0f : sample);
            }
        }
        public void AddOptional(double value)
        {
            if(double.IsInfinity(value)) throw new ArgumentException("Invalid optional fingerprint coordinate.");
            Guard(10);writer.Write((byte)5);writer.Write(double.IsNaN(value));
            if(!double.IsNaN(value)) writer.Write(value==0 ? 0.0 : value);
        }
        public string Complete()
        {
            Guard(0);finished=true;writer.Flush();buffer.Position=0;
            using(var sha=SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(buffer)).Replace("-","");
        }
        public void Dispose() { finished=true;writer.Dispose(); }
        private void Guard(int additional)
        {
            if(finished) throw new InvalidOperationException("Fingerprint writer is already closed.");
            if(buffer.Length+additional>MaximumBytes) throw new ArgumentException("Native fingerprint exceeds its byte budget.");
        }
        private static void RequireFinite(double value)
        { if(double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("Fingerprint value must be finite."); }
    }
}
