using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace ISTools
{
    internal class ObjRvtLink
    {
        public string linkName;
        public int linkCount;
        public Document currentDoc;
        public int GetLinkCount(List<Element> links)
        {
            int n = 0;
            if (linkName == currentDoc.Title.Replace(".rvt", ""))
                n = 1;
            else
            {
                foreach (Element el in links)
                {
                    if (el.Name.Split(':')[0].Replace(".rvt", "").TrimEnd(' ') == linkName)
                    {
                        n++;
                    }
                }
            }
            return n;
        }
    }
}
