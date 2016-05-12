﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Drawing;

namespace PPTProgressMaker
{
    public partial class ThisAddIn
    {
        public const string HORIZ_SMART_ART_OBJECT_NAME = "Closed Chevron Process";
        public const string VERT_SMART_ART_OBJECT_NAME = "Vertical Curved List";
        public const string SHAPE_TAG = "AEContents";

        public static ThisAddIn instance;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public enum TocType { StyleHorizontal, StyleVertical };
        public enum TocStyle { StyleSolid, StyleGradient };

        public void addToC(TocType type, TocStyle style, bool rtl)
        {
            applyStyleDelegate styler;
            switch (style)
            {
                case TocStyle.StyleSolid: styler = getSolidStyle(); break;
                case TocStyle.StyleGradient: styler = getGradientStyle(); break;
                default:
                    throw new Exception("Unknown color style");
            }

            string[] secnames;
            double[] pcnt;

            deleteOldToC();
            buildToCData(out secnames, out pcnt);

            switch(type)
            {
                case TocType.StyleHorizontal:
                    addHorizontalToC(secnames, pcnt, rtl, styler);
                    break;
                case TocType.StyleVertical:
                    addVerticalToC(secnames, pcnt, rtl, styler);
                    break;

            }

        }

        delegate void applyStyleDelegate(Microsoft.Office.Core.ShapeRange shapes, int shapeidx, int sldidx, double sldpcnt);
            
        private applyStyleDelegate getSolidStyle()
        {
            var color1 = Globals.Ribbons.Ribbon.getNormalColor();
            var color2 = Globals.Ribbons.Ribbon.getActiveColor();

            return (shapes, shapeidx, sldidx, sldpcnt) =>
                {
                    var col = (sldidx == shapeidx) ? color2 : color1;
                    shapes.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientVertical, 1);
                    var gs = shapes.Fill.GradientStops;
                    gs.Insert(col, 0f, 0);
                    gs.Insert(col, 1f);
                    gs.Delete(1);
                    gs.Delete(1);
                };
        }

        private applyStyleDelegate getGradientStyle()
        {
            var color1 = Globals.Ribbons.Ribbon.getNormalColor();
            var color2 = Globals.Ribbons.Ribbon.getActiveColor();

            return (shapes, shapeidx, sldidx, sldpcnt) =>
            {
                var col1 = color1;
                var col2 = color1;
                var col3 = color1;
                if (sldidx == shapeidx)
                {
                    col3 = color2;
                }
                else if (sldidx == shapeidx)
                {
                    col1 = col2 = col3 = color2;
                }

                shapes.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientVertical, 1);
                var gs = shapes.Fill.GradientStops;
                gs.Insert(col1, 0f, 0);
                gs.Insert(col2, 0.5f);
                gs.Insert(col3, 1f);
                gs.Delete(1);
                gs.Delete(1);
            };
        }

        private void buildToCData(out string[] secnames, out double[] pcnt)
        {
            var p = Application.ActivePresentation;
            int num_slides = p.Slides.Count;

            pcnt = new double[num_slides];

            // per section
            var sections = p.SectionProperties;
            secnames = new string[sections.Count];
            for (int i = 1; i <= sections.Count; ++i)
            {
                string name = sections.Name(i);
                secnames[i - 1] = name;
                int first_slide = sections.FirstSlide(i);
                int last_slide = (i + 1 <= sections.Count) ? sections.FirstSlide(i + 1) : num_slides;
                for (int j = first_slide; j <= last_slide; ++j)
                    pcnt[j - 1] = (j - first_slide + 1.0) / (last_slide - first_slide + 1.0);
            }
        }

        private void deleteOldToC()
        {
            var p = Application.ActivePresentation;
            var slides = p.Slides;
            int num_slides = slides.Count;

            for (int i = 0; i < num_slides; ++i)
            {
                var slide = slides[i + 1];
                var old = getShapeSafe(slide, SHAPE_TAG);
                if (old != null)
                    old.Delete();
            }
        }

        private void addHorizontalToC(string[] secnames, double[] pcnt, bool rtl, applyStyleDelegate styler)
        {
            var p = Application.ActivePresentation;
            var slides = p.Slides;
            int num_slides = slides.Count;

            int id = getSmartArtObjectByName(HORIZ_SMART_ART_OBJECT_NAME);
            var layout = Application.SmartArtLayouts[id];
            float height = p.PageSetup.SlideHeight / 20;
            float width = p.PageSetup.SlideWidth;
            float left = 0;
            float top = p.PageSetup.SlideHeight - height;
            
            for (int i = 0; i < num_slides; ++i)
            {
                var slide = slides[i + 1];
               
                var shape = slide.Shapes.AddSmartArt(layout, left, top, width, height);
                shape.Name = SHAPE_TAG;
                if (rtl)
                    shape.SmartArt.Reverse = Microsoft.Office.Core.MsoTriState.msoTrue;

                formatToCShape(shape, i, slide.SectionNumber, secnames, pcnt, styler);
            }
        }
        private void addVerticalToC(string[] secnames, double[] pcnt, bool rtl, applyStyleDelegate styler)
        {
            var p = Application.ActivePresentation;
            var slides = p.Slides;
            int num_slides = slides.Count;

            int id = getSmartArtObjectByName(VERT_SMART_ART_OBJECT_NAME);
            var layout = Application.SmartArtLayouts[id];
            float height = p.PageSetup.SlideHeight;
            float width = p.PageSetup.SlideWidth / 6;
            float left = rtl ? (p.PageSetup.SlideWidth - width) : 0;
            float top = 0;

            for (int i = 0; i < num_slides; ++i)
            {
                var slide = slides[i + 1];

                var shape = slide.Shapes.AddSmartArt(layout, left, top, width, height);
                shape.Name = SHAPE_TAG;
                if(rtl)
                    shape.SmartArt.Reverse = Microsoft.Office.Core.MsoTriState.msoTrue;

                formatToCShape(shape, i, slide.SectionNumber, secnames, pcnt, styler);
            }
        }

        private void formatToCShape(PowerPoint.Shape shape, int sldindex, int secindex, string[] secnames, double[] pcnt, applyStyleDelegate styler)
        {
            for(int i = shape.SmartArt.Nodes.Count + 1; i<= secnames.Length; ++ i)
                shape.SmartArt.Nodes.Add();

            for (int i = 0; i < secnames.Length; ++i)
            {
                var node = shape.SmartArt.Nodes[i + 1];
                node.TextFrame2.TextRange.Text = secnames[i];
                styler(node.Shapes, i, secindex, pcnt[sldindex]);
            }
        }

        private PowerPoint.Shape getShapeSafe(PowerPoint.Slide slide, string name)
        {
            try
            {
                return slide.Shapes[name];
            }
            catch
            {
                return null;
            }
        }

        private int getSmartArtObjectByName(string name)
        {
            for(int i = 1; i <= Application.SmartArtLayouts.Count; ++i)
            {
                Debug.WriteLine(Application.SmartArtLayouts[i].Name);
                if (Application.SmartArtLayouts[i].Name.Equals(name))
                    return i;
            }
            throw new Exception("Could not find smart art object");
        }

        private PowerPoint.Shape findShapeByTag(PowerPoint.Slide slide, string name)
        {
            for (int i = 1; i < slide.Shapes.Count; ++i)
            {
                var s = slide.Shapes[i];
                if (s.Tags[name] != null) 
                        return s;
            }
            return null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}