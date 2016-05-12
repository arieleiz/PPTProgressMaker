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
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            setEnabled(false);
            Globals.ThisAddIn.Application.ColorSchemeChanged += Application_ColorSchemeChanged;
            Globals.ThisAddIn.Application.PresentationOpen += Application_PresentationOpen;
            Globals.ThisAddIn.Application.PresentationClose += Application_PresentationClose;
        }

        private void Application_PresentationOpen(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
        {
            setEnabled(true);
            Application_ColorSchemeChanged(null);
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
            try
            {
                ThisAddIn.instance.addToC(ThisAddIn.TocType.StyleHorizontal, getStyle(), getRTL());
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVertToc_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ThisAddIn.instance.addToC(ThisAddIn.TocType.StyleVertical, getStyle(), getRTL());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private ThisAddIn.TocStyle getStyle()
        {
            if (cbGradient.Checked)
                return ThisAddIn.TocStyle.StyleGradient;
            else
                return ThisAddIn.TocStyle.StyleSolid;
        }

        private bool getRTL()
        {
            return cbRTL.Checked;
        }

        private void glActive_Click(object sender, RibbonControlEventArgs e)
        {
            setGalleryColor(glActive.SelectedItemIndex, glActive);
        }

        private void glNormal_Click(object sender, RibbonControlEventArgs e)
        {
            setGalleryColor(glNormal.SelectedItemIndex, glNormal);
        }

        public int getNormalColor()
        {
            return (int)glNormal.Tag;
        }

        public int getActiveColor()
        {
            return (int)glActive.Tag;
        }

    }

}
