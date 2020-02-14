using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordPrueba
{
    public class Word : IDisposable
    {
        private WordprocessingDocument cWordprocessing;
        private Document cDocument;
        private IEnumerable<BookmarkStart> cBookmarks;

        public static void CreateDocument(string pvStrPath, string pvStrName)
        {
            var vStrFullPath = Path.Combine(pvStrPath, pvStrName);
            // Create Document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(vStrFullPath, WordprocessingDocumentType.Document, true))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                // Create the document structure and add some text.
                mainPart.Document = new Document();
                mainPart.Document.AppendChild(new Body());
                // Close the handle explicitly.
                wordDocument.Close();
            }
        }

        public static TableProperties GetTableProperties(uint pvSize = 10, BorderValues borderValues = BorderValues.Single)
        {
            return new TableProperties(
                new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(borderValues), Size = pvSize },
                new BottomBorder { Val = new EnumValue<BorderValues>(borderValues), Size = pvSize },
                new LeftBorder { Val = new EnumValue<BorderValues>(borderValues), Size = pvSize },
                new RightBorder { Val = new EnumValue<BorderValues>(borderValues), Size = pvSize },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(borderValues), Size = pvSize },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(borderValues), Size = pvSize }
           ));
        }

        public void OpenDoc(string pvStrPath, string pvStrName)
        {
            var vStrFullPath = Path.Combine(pvStrPath, pvStrName);
            if (!File.Exists(vStrFullPath))
                return;
            cWordprocessing = WordprocessingDocument.Open(vStrFullPath, true);
            cDocument = cWordprocessing.MainDocumentPart.Document;
            cBookmarks = GetAllBookmarks(cDocument);
        }

        public static void AddTable(string fileName, DataTable data)
        {
            if (data != null && data.Rows.Count == 0)
                return;
            var filas = data.Rows.Count;
            var columnas = data.Columns.Count;

            //Abrir documento
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                //Obtener el documento
                var doc = document.MainDocumentPart.Document;
                //Innicializa el cuerpo del documento si es null
                if (document.MainDocumentPart.Document == null)
                    document.MainDocumentPart.Document.AppendChild(new Body());

                Table table = new Table();
                //definir las propiedades  de la tabla
                TableProperties props = GetTableProperties();
                //Adjuntar las propiedades a la tabla
                table.AppendChild(props);

                for (var i = 0; i < filas; i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j < columnas; j++)
                    {
                        //Crear una celda
                        var tc = new TableCell();
                        //agrega un parrafo con un elemento text con el valor a la celda
                        tc.Append(new Paragraph(new Run(new Text(data.Rows[i][j]?.ToString()))));
                        // Agrega propiedades a la celda de autoajuste
                        tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                        //agrega la celda a la fila de la tabla
                        tr.Append(tc);
                    }
                    //Agrega la fila a la tabla
                    table.Append(tr);
                }
                //Agregar la tabla al cuerpo del documento
                doc.Body.Append(table);
                doc.Save();
            }
        }

        public void AddTable(DataTable data)
        {
            if (data != null && data.Rows.Count == 0)
                return;
            if (cWordprocessing == null || cWordprocessing?.MainDocumentPart?.Document == null)
                return;
            //Obtener el documento
            var doc = cWordprocessing.MainDocumentPart.Document;
            var columnas = data.Columns.Count;

            Table table = new Table();
            //definir las propiedades  de la tabla
            TableProperties props = GetTableProperties();
            //Adjuntar las propiedades a la tabla
            table.AppendChild(props);
            foreach (DataRow vFila in data.Rows)
            {
                var tr = new TableRow();
                for (var j = 0; j < columnas; j++)
                {
                    //Crear una celda
                    var tc = new TableCell();
                    //agrega un parrafo con un elemento text con el valor a la celda
                    tc.Append(new Paragraph(new Run(new Text(vFila[j]?.ToString()))));
                    // Agrega propiedades a la celda de autoajuste
                    tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                    //agrega la celda a la fila de la tabla
                    tr.Append(tc);
                }
                //Agrega la fila a la tabla
                table.Append(tr);
            }
            //Agregar la tabla al cuerpo del documento
            doc.Body.Append(table);
            doc.Save();
        }

        public List<string> GetBookmarks()
        {
            List<string> vLstBookMarks = new List<string>();
            if (cDocument == null) return vLstBookMarks;

            var BookMarks = GetDocumentBookmarkValues();
            foreach (var mark in BookMarks)
                vLstBookMarks.Add(mark.Key);
            return vLstBookMarks;
        }

        public void WriteBookMark(string pvStrNameBookMark, string value)
        {
            if (string.IsNullOrWhiteSpace(pvStrNameBookMark) || cDocument == null || cBookmarks == null) return;

            var bookMark = cBookmarks.Where(mark => mark.Name.Equals(pvStrNameBookMark)).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(bookMark?.Name)) return;

            SetText(bookMark, value);
        }


        public void Dispose()
        {
            if (cWordprocessing == null)
                return;
            cWordprocessing.MainDocumentPart?.Document?.Save();
            cWordprocessing.Close();
            cWordprocessing.Dispose();
            cWordprocessing = null;
        }

        public IDictionary<string, string> GetDocumentBookmarkValues(bool includeHiddenBookmarks = false)
        {
            return GetDocumentBookmarkValues(cDocument, includeHiddenBookmarks);
        }

        #region Metodos Estaticos

        public static IDictionary<string, string> GetDocumentBookmarkValues(Document document, bool includeHiddenBookmarks = false)
        {
            IDictionary<string, string> bookmarks = new Dictionary<string, string>();

            foreach (var bookmark in GetAllBookmarks(document))
            {
                if (includeHiddenBookmarks || !IsHiddenBookmark(bookmark.Name))
                    bookmarks[bookmark.Name] = GetText(bookmark);
            }
            return bookmarks;
        }

        public static void SetDocumentBookmarkValues(Document document, IDictionary<string, string> bookmarkValues)
        {
            foreach (var bookmark in GetAllBookmarks(document))
                SetBookmarkValue(bookmark, bookmarkValues);
        }

        public static void SetText(BookmarkStart bookmark, string value)
        {
            //Retorna un elemento Text del Marcador que se esta buscando
            var text = FindBookmarkText(bookmark);
            if (text != null)
            {
                //Asigna el texto
                text.Text = value;
                //
                RemoveOtherTexts(bookmark, text);
            }
            else
                InsertBookmarkText(bookmark, value);
        }

        private static IEnumerable<BookmarkStart> GetAllBookmarks(Document document)
        {
            return document.MainDocumentPart.RootElement.Descendants<BookmarkStart>();
        }

        private static bool IsHiddenBookmark(string bookmarkName) => bookmarkName.StartsWith("_");

        public static string GetText(BookmarkStart bookmark)
        {
            var text = FindBookmarkText(bookmark);
            return text != null ? text.Text : string.Empty;
        }

        private static void SetBookmarkValue(BookmarkStart bookmark, IDictionary<string, string> bookmarkValues)
        {
            if (bookmarkValues.TryGetValue(bookmark.Name, out string value))
                SetText(bookmark, value);
        }


        private static Text FindBookmarkText(BookmarkStart bookmark)
        {
            if (bookmark.ColumnFirst != null)
                return FindTextInColumn(bookmark);
            else
            {
                var run = bookmark.NextSibling<Run>();

                if (run != null)
                    //obtener el primer elemento de tipo Text si es distinto de null
                    return run.GetFirstChild<Text>();
                else
                {
                    Text text = null;
                    //Obtiene el siguiente elemento Openxml (Si es null significa que no hay siguiente)
                    var nextSibling = bookmark.NextSibling();
                    //Itera los elementos hasta encontrar el ultimo de tipo Text
                    while (text == null && nextSibling != null)
                    {
                        if (nextSibling.IsEndBookmark(bookmark))
                            return null;
                        //Obtiene el texto del siguiente elemento
                        text = nextSibling.GetFirstDescendant<Text>();
                        //Obtiene el siguiente elemento
                        nextSibling = nextSibling.NextSibling();
                    }

                    return text;
                }
            }
        }

        /// <summary>
        /// Obtiene la propiedad Text de un marcador que se encuentre en una tabla
        /// </summary>
        /// <param name="bookmark"></param>
        /// <returns></returns>
        private static Text FindTextInColumn(BookmarkStart bookmark)
        {
            //Obtiene las celdas de la tabla
            var cell = bookmark.GetParent<TableRow>().GetFirstChild<TableCell>();

            for (int i = 0; i < bookmark.ColumnFirst; i++)
                cell = cell.NextSibling<TableCell>();
            return cell.GetFirstDescendant<Text>();
        }

        private static void RemoveOtherTexts(BookmarkStart bookmark, Text keep)
        {
            if (bookmark.ColumnFirst != null) return;

            Text text = null;
            //Obtiene el siguiente elemento
            var nextSibling = bookmark.NextSibling();
            while (text == null && nextSibling != null)
            {
                //si es un marcador de fin, se termina el ciclo
                if (nextSibling.IsEndBookmark(bookmark))
                    break;
                //Obtiene los elementos Text
                foreach (var item in nextSibling.Descendants<Text>())
                {
                    //Si es distinto del Text del marcador se elimina
                    if (item != keep)
                        item.Remove();
                }
                nextSibling = nextSibling.NextSibling();
            }
        }

        private static void InsertBookmarkText(BookmarkStart bookmark, string value)
        {
            //Inserta el elemento Text despues 
            bookmark.Parent.InsertAfter(new Run(new Text(value)), bookmark);
        }
        #endregion Metodos Estaticos

    }
}
