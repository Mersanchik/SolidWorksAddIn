using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

namespace SolidWorksAddIn
{
    internal static class SwHelpers
    {
        public static swDocumentTypes_e GetDocType(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            switch (ext)
            {
                case ".sldprt": return swDocumentTypes_e.swDocPART;
                case ".sldasm": return swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw": return swDocumentTypes_e.swDocDRAWING;
                default: return swDocumentTypes_e.swDocNONE;
            }
        }

        public static string GetPathSafe(IModelDoc2 doc)
        {
            try
            {
                return doc?.GetPathName() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static void ForEachDrawingView(DrawingDoc drw, Action<IView> action)
        {
            if (drw == null) return;

            IView view = drw.GetFirstView();
            if (view != null)
            {
                view = view.GetNextView();
            }

            while (view != null)
            {
                action(view);
                view = view.GetNextView();
            }
        }
    }
}
