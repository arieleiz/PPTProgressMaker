using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;

namespace PPTProgressMaker
{
    public partial class Ribbon
    {
        private static readonly string AE_ACTIVE_COL_TAG = "ae_active_color";
        private static readonly string AE_NORMAL_COL_TAG = "ae_normal_color";

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
            setPresTag(Pres, AE_ACTIVE_COL_TAG, ActiveColor.ToString());
            setPresTag(Pres, AE_NORMAL_COL_TAG, NormalColor.ToString());
        }

        private void Application_PresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            setEnabled(true);
            Application_ColorSchemeChanged(null);
            loadPresColor(Pres, AE_ACTIVE_COL_TAG, glActive);
            loadPresColor(Pres, AE_NORMAL_COL_TAG, glNormal);
        }

        private void loadPresColor(Microsoft.Office.Interop.PowerPoint.Presentation Pres, string tag, RibbonGallery gl)
        {
            int value;
            string str = getPresTag(Pres, tag);
            if (!String.IsNullOrEmpty(str) && int.TryParse(str, out value))
                setColor(gl, value);

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
                ThisAddIn.instance.addToC(new ToCStyle() { Type = type, Style = getStyle(), RTL = getRTL(), FirstSlide = isFirstSlide() });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private ToCStyle.Styles getStyle()
        {
            if (cbGradient.Checked)
                return ToCStyle.Styles.Gradient;
            else
                return ToCStyle.Styles.Solid;
        }

        private bool getRTL()
        {
            return cbRTL.Checked;
        }

        private bool isFirstSlide()
        {
            return cbFirstSlide.Checked;
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

    }

}
