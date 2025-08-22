using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksAddIn
{
    internal static class SwFileSearch
    {
        public static IEnumerable<string> FindDrawingsSibling(string modelFullPath)
        {
            string dir = Path.GetDirectoryName(modelFullPath);
            string baseName = Path.GetFileNameWithoutExtension(modelFullPath);

            // Ищем чертежи на одном уровне, которые относятся к модели
            foreach (var f in Directory.EnumerateFiles(dir, "*.slddrw", SearchOption.TopDirectoryOnly))
            {
                var name = Path.GetFileNameWithoutExtension(f);
                if (string.Equals(name, baseName, StringComparison.InvariantCultureIgnoreCase))
                {
                    yield return f;
                }
            }
        }

        public static string ComposeRenamedPath(string oldFullPath, string newBaseName)
        {
            string dir = Path.GetDirectoryName(oldFullPath);
            string ext = Path.GetExtension(oldFullPath);
            return Path.Combine(dir, newBaseName + ext);
        }
    }
}
