using System.IO;

namespace Smartmill.Reporting.OpenXml.Utils
{
    internal class DocumentLoader : IDocumentLoader
    {
        protected Stream _stream;

        public void Load(string path)
        {
            _stream = new FileStream(path, FileMode.OpenOrCreate);
        }

        public void Load(Stream stream)
        {
            _stream = stream;
        }
    }
}