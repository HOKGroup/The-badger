using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Graph = Microsoft.Office.Interop.Graph;
using System.Runtime.InteropServices;

namespace PowerRhodes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //fields
        string[] arrImages;
        string[] arrTexts;
        string strTemplate;

        private void btnGetPictures_Click(object sender, EventArgs e)
        {
            openFolder.Filter = "Images (*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF";
            openFolder.Multiselect = true;
            openFolder.FileName = "Select Pictures";

            DialogResult dr = openFolder.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                arrImages = openFolder.FileNames;
                lblImageCount.Text = arrImages.Length + "  Bloody pictures";

            }
        }

        private void btnGetTemplate_Click(object sender, EventArgs e)
        {
            openFolder.Filter = "PPt file (*.pptx)|*.pptx";
            openFolder.Multiselect = false;
            openFolder.FileName = "Select PowerPoint for template";
            DialogResult dr = openFolder.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                strTemplate = openFolder.FileName;
                lblTemplate.Text = "template loaded, awesome!!";
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            bool yn = GetLines();

            if (yn)
            {
                ShowPresentation();
                GC.Collect();
            }
        }

        private void ShowPresentation()
        {
            PowerPoint.Application objApp;
            PowerPoint.Presentations objPresSet;
            PowerPoint._Presentation objPres;
            PowerPoint.Slides objSlides;
            PowerPoint._Slide objSlide;
            PowerPoint.TextRange objTextRng;

            //create new presentation based on a template...
            objApp = new PowerPoint.Application();
            objApp.Visible = MsoTriState.msoTrue;
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(strTemplate, MsoTriState.msoTrue, MsoTriState.msoTrue);
            objSlides = objPres.Slides;


            //we need to create the page...
            int count = 1;
            int countText = 0;

            string mess = "";
            for (int i = 0; i < arrImages.Length; i++)
            {
                //add fixed images
                objSlide = objSlides.Add(count, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                objSlide.Shapes.AddPicture(arrImages[i], MsoTriState.msoFalse, MsoTriState.msoTrue, 51.1475f, 14.4339f, 618.0155f, 437.0040f);

                // Add title
                if (arrTexts != null)
                {
                    objTextRng = objSlide.Shapes[1].TextFrame.TextRange;

                    if (countText > arrTexts.Length - 1)
                    {
                        mess = arrTexts[arrTexts.Length - 1];
                    }
                    else
                    {
                        mess = arrTexts[countText];
                    }

                    objTextRng.Text = mess;

                    count++;
                    countText++;
                }
            }

        }

        private bool GetLines()
        {
            string[] allLines = txtBoxTitles.Text.Split('\n');
            bool yn = false;
            
            string ms = txtBoxTitles.Text;
 
            if (ms.Length == 0) //have to review....
            {
                DialogResult dialogResult = MessageBox.Show("No bloody titles?, do you want to create the bloody presentation?", "Badger says: ", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    yn = true;
                }
                else
                {
                    yn = false;
                }
            }
            else
            {
                arrTexts = allLines;
                yn = true;
            }

            return yn;
        }

        private void btnShit_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            int seed = dt.Second;
            Random rnd = new Random(seed);
            int num = rnd.Next(1, 3801);

            string adress = "http://montondemierda.com/page/" + num + "#.VC8IWfldU-L";
            System.Diagnostics.Process.Start(adress);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
