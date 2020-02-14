using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordPrueba
{
    public static class ExtensionOpenXml
    {
        /*** EXTENSTION METHODS START ***/

        /// <summary>
        /// Obtiene el primer elemento hijo del tipo Especificado (OpenXml)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static T GetFirstDescendant<T>(this OpenXmlElement parent) where T : OpenXmlElement
        {
            var descendants = parent.Descendants<T>();
            if (descendants != null)
                return descendants.FirstOrDefault();
            else
                return null;
        }

        /// <summary>
        /// Obtiene el padre del objeto tipo Openxml
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="child"></param>
        /// <returns></returns>
        public static T GetParent<T>(this OpenXmlElement child) where T : OpenXmlElement
        {
            while (child != null)
            {
                child = child.Parent;
                if (child is T)
                    return (T)child;
            }
            return null;
        }

        /// <summary>
        /// Evalua si es un marcador de fin
        /// </summary>
        /// <param name="element"></param>
        /// <param name="startBookmark"></param>
        /// <returns></returns>
        public static bool IsEndBookmark(this OpenXmlElement element, BookmarkStart startBookmark)
        {
            return IsEndBookmark(element as BookmarkEnd, startBookmark);
        }

        public static bool IsEndBookmark(this BookmarkEnd endBookmark, BookmarkStart startBookmark)
        {
            return endBookmark == null ? false : endBookmark.Id == startBookmark.Id;
        }

        /*** EXTENSTION METHODS  END ***/
    }
}
