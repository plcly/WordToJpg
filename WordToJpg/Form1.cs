using Aspose.Words;
using Aspose.Words.Saving;
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

namespace WordToJpg
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files.Length > 0)
            {
                var extension = Path.GetExtension(files[0]);
                if (extension == ".docx" || extension == ".doc" || extension == ".wps")
                {
                    var path = files[0];
                    var jpgName = Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path)+".jpg");
                    var doc = new Document(path);
                    ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);
                    options.Resolution = 300;
                    options.JpegQuality = 100;
                    options.UseHighQualityRendering = true;
                    doc.Save(jpgName, options);
                }
            }
        }
    }
}
