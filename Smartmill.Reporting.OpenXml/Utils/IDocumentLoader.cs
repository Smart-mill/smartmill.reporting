using System.IO;

namespace Smartmill.Reporting.OpenXml.Utils
{
    public interface IDocumentLoader
    {
        void Load(string path);

        void Load(Stream stream);
    }
}