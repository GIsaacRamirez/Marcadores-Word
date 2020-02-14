using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordPrueba
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //Word.CreateDocument(@"C:\Users\Isac\Desktop", "prueba.docx");
            Word document = new Word();
            document.OpenDoc(@"C:\Users\Isac\Desktop", "prueba.docx");

            DataTable dt = new DataTable();//suppose data comes frome database
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            dt.Columns.Add("Sex");
            dt.Rows.Add(1, "Tom", "male");
            dt.Rows.Add(2, "Jim", "male");
            dt.Rows.Add(3, "LiSa", "female");
            dt.Rows.Add(4, "LiLi", "female");
            //Word.AddTable(Path.Combine(@"C:\Users\Isac\Desktop", "prueba.docx"), dt);
            document.AddTable(dt);

            var listaMarcadores = document.GetBookmarks( );


            document.WriteBookMark(listaMarcadores?.FirstOrDefault(), "Hola Isaac");

            document.Dispose();
        }
    }
}
