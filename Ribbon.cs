using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace PPTProgressMaker
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            setEnabled(false);
            Globals.ThisAddIn.Application.ColorSchemeChanged += Application_ColorSchemeChanged;
            Globals.ThisAddIn.Application.PresentationOpen += Application_PresentationOpen;
            Globals.ThisAddIn.Application.PresentationClose += Application_PresentationClose;
            Globals.ThisAddIn.Application.PresentationBeforeSave += Application_PresentationBeforeSave;
        }

        private void Application_PresentationBeforeSave(Microsoft.Office.Interop.PowerPoint.Presentation Pres, ref bool Cancel)
        {
        }

        private void Application_PresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            setEnabled(true);
            Application_ColorSchemeChanged(null);
            loadStyle();
        }

        
        private void Application_PresentationClose(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            setEnabled(false);
        }

        private void Application_ColorSchemeChanged(Microsoft.Office.Interop.PowerPoint.SlideRange SldRange)
        {
            var presentation = Globals.ThisAddIn.Application.ActivePresentation;
            if (presentation != null)
            {
                ColorSchemeChanged(presentation);
            }
        }

        private void setEnabled(bool enabled)
        {
            foreach (var t in this.Tabs)
            {
                foreach (var g in t.Groups)
                {
                    foreach (var c in g.Items)
                    {
                        c.Enabled = enabled;
                    }
                }
            }            
        }

        private void ColorSchemeChanged(Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {
            var colors = presentation.Designs[1].SlideMaster.Theme.ThemeColorScheme;
            var coloridx = new int[]
            {
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent1).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent2).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent3).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent4).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent5).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent6).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark1).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark2).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight1).RGB,
                colors.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight2).RGB,
            };

            initNormalColors(coloridx, 0, glNormal);
            initNormalColors(coloridx, 5, glActive);
        }

        private void initNormalColors(int[] coloridx, int active, RibbonGallery gallery)
        {
            var items = gallery.Items;
            items.Clear();

            for (int i = 0; i < coloridx.Length; ++i)
            {
                items.Add(Factory.CreateRibbonDropDownItem());
                items[i].Label = String.Format("Accent {0}", i);
                items[i].Image = get16pxImage(coloridx[i]);
                items[i].Tag = coloridx[i];
            }
            setGalleryColor(active, gallery);
        }

        private static void setGalleryColor(int active, RibbonGallery gallery)
        {
            var items = gallery.Items;
            gallery.Image = items[active].Image;
            gallery.Tag = items[active].Tag;
        }

        private static Image get16pxImage(int rgb)
        {
            int size = 16;
            var bmp = new Bitmap(size, size);
            using (Graphics gfx = Graphics.FromImage(bmp))
            {
                using (SolidBrush brush = new SolidBrush(ColorTranslator.FromOle(rgb)))
                {
                    gfx.FillRectangle(brush, 0, 0, size, size);
                }
            }
            return bmp;
        }

        private void btnHorizTOC_Click(object sender, RibbonControlEventArgs e)
        {
            addToC(ToCStyle.Types.Horizontal);
        }

        private void btnVertToc_Click(object sender, RibbonControlEventArgs e)
        {
            addToC(ToCStyle.Types.Vertical);
        }

        private void addToC(ToCStyle.Types type)
        {
            try
            {
                var style = new ToCStyle() { Type = type, Style = Style, RTL = RTL, FirstSlide = FirstSlide, SlideNumbers = SlideNumbers, IgnoreLastSection = IgnoreLastSection, NormalColor =  NormalColor, ActiveColor = ActiveColor};
                storeStyle(style);
                ThisAddIn.instance.addToC(style);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private ToCStyle.Styles Style
        {
            get
            {
                if (cbGradient.Checked)
                    return ToCStyle.Styles.Gradient;
                else
                    return ToCStyle.Styles.Solid;
            }
        }

        private bool RTL
        {
            get { return cbRTL.Checked; }
        }

        private bool FirstSlide
        {
            get
            {
                return cbFirstSlide.Checked;
            }
        }
        public bool IgnoreLastSection { get { return cbIgnoreLastSec.Checked; } }

        private bool SlideNumbers
        {
            get
            {
                return cbSlideNumbers.Checked;
            }
        }
        private void glActive_Click(object sender, RibbonControlEventArgs e)
        {
            setGalleryColor(glActive.SelectedItemIndex, glActive);
        }

        private void glNormal_Click(object sender, RibbonControlEventArgs e)
        {
            setGalleryColor(glNormal.SelectedItemIndex, glNormal);
        }

        private void setColor(RibbonGallery gl, int value)
        {
            for(int i = 0; i < gl.Items.Count; ++ i)
            {
                if((int)gl.Items[i].Tag == value)
                {
                    gl.SelectedItemIndex = i;
                    setGalleryColor(i, gl);
                    return;
                }
            }

            int pos = gl.Items.Count;
            gl.Items.Add(Factory.CreateRibbonDropDownItem());
            gl.Items[pos].Label = "Custom";
            gl.Items[pos].Image = get16pxImage(value);
            gl.Items[pos].Tag = value;
            setGalleryColor(pos, gl);
        }

        public int NormalColor
        {
            get {
                return (int)glNormal.Tag;
            }
            set
            {
                setColor(glNormal, value);
            }
        }

        public int ActiveColor
        {
            get {
                return (int)glActive.Tag;
            }
            set
            {
                setColor(glActive, value);
            }
        }


        private void setPresTag(Microsoft.Office.Interop.PowerPoint.Presentation Pres, string tag, string value)
        {
            try
            {
                Pres.Tags.Delete(tag);
            }
            catch
            {
            }
            Pres.Tags.Add(tag, value);
        }

        private string getPresTag(Microsoft.Office.Interop.PowerPoint.Presentation Pres, string tag)
        {
            try
            {
                return Pres.Tags[tag];
            }
            catch
            {
                return null;
            }
        }
        private void loadStyle()
        {
            try
            {
                var state = Convert.FromBase64String(getPresTag(Globals.ThisAddIn.Application.ActivePresentation, "ae_state"));
                using (var stream = new MemoryStream(state))
                {
                    var bf = new BinaryFormatter();
                    var style = (ToCStyle)bf.Deserialize(stream);
                    applyUIStyle(style);
                }
            }
            catch { }
        }

        private void applyUIStyle(ToCStyle style)
        {
            cbGradient.Checked = style.Style == ToCStyle.Styles.Gradient;
            cbRTL.Checked = style.RTL;
            cbFirstSlide.Checked = style.FirstSlide;
            cbSlideNumbers.Checked = style.SlideNumbers;
            cbIgnoreLastSec.Checked = style.IgnoreLastSection;
            ActiveColor = style.ActiveColor;
            NormalColor = style.NormalColor;
        }

        private void storeStyle(ToCStyle style)
        {
            BinaryFormatter bf = new BinaryFormatter();
            using (MemoryStream stream = new MemoryStream())
            {
                bf.Serialize(stream, style);
                byte[] data = stream.ToArray();
                string state = Convert.ToBase64String(data);
                setPresTag(Globals.ThisAddIn.Application.ActivePresentation, "ae_state", state);
            }
        }

    }

}
