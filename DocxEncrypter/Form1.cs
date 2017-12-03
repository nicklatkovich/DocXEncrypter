using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace DocxEncrypter {
    public partial class Form1 : Form {

        DocX sourceDocX = null;
        DocX targetDocX = null;

        public Form1( ) {
            InitializeComponent( );
            LoadSourceDocX(@"E:/temp.docx");
            foreach (var paragraph in targetDocX.Paragraphs) {
                paragraph.Append("qwe").SpacingAfter(0.5f);
            }
            targetDocX.SaveAs(@"E:/tempRes.docx");
            Close( );
        }

        public void LoadSourceDocX(string filename) {
            sourceDocX?.Dispose( );
            targetDocX?.Dispose( );
            sourceDocX = DocX.Load(filename);
            targetDocX = sourceDocX.Copy( );
        }
    }
}
