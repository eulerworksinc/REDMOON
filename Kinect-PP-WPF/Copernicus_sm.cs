﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Kinect;
using Microsoft.Kinect.Toolkit;


namespace Kinect_PP_WPF
{
    /// <summary>
    /// State machine class for Copernicus
    /// </summary>
    class Copernicus_sm
    {
        private int current_slide;
        private int num_slides;
        private Application pp_app;
        private Presentation pp_presentation;
        private SlideShowView pp_slideshow;
        private List<Copernicus_slide> slide_list;
        private string pp_filename;
        private Button_list button_list;

        /// <summary>
        /// Constructor for Copernicus_sm
        /// </summary>
        /// <param name="xml_filename"></param>
        /// <param name="button_list"></param>
        public Copernicus_sm(string xml_filename, Button_list button_list)
        {
            try
            {
                XmlDocument xml_file = new XmlDocument();
                xml_file.Load(xml_filename);
                XmlElement root = xml_file.DocumentElement;
                pp_filename = Path.GetDirectoryName(xml_filename);
                pp_filename += "\\" + root.GetAttribute("filename");
                slide_list = new List<Copernicus_slide>();
                num_slides = root.ChildNodes.Count;
                this.button_list = button_list;

                /// Create new copernicus slide for each slide in xml
                foreach (XmlNode node in root.ChildNodes)
                {
                    Copernicus_slide slide = new Copernicus_slide();
                    slide.id = node.Attributes["id"].Value;
                    foreach (XmlNode button in node.ChildNodes)
                    {
                        switch (button.Attributes["id"].Value)
                        {
                            case "right":
                                slide.right_button = true;
                                break;
                            case "left":
                                slide.left_button = true;
                                break;
                            case "next":
                                slide.next_slide_button = true;
                                break;
                            case "previous":
                                slide.prev_slide_button = true;
                                break;
                            case "close":
                                slide.close_button = true;
                                break;
                            default:
                                break;
                        }
                    }
                    slide_list.Add(slide);
                }

            }
            catch
            {
                throw;
            }

        }

        /// <summary>
        /// Start powerpoint slideshow
        /// </summary>
        public void start_pp()
        {
            pp_app = new Application();
            pp_app.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            pp_presentation = pp_app.Presentations.Open(pp_filename);
            pp_presentation.SlideShowSettings.ShowPresenterView = Microsoft.Office.Core.MsoTriState.msoFalse;
            pp_presentation.SlideShowSettings.Run();
            pp_slideshow = pp_presentation.SlideShowWindow.View;
            current_slide = 1;
            goto_slide(current_slide);
        }

        /// <summary>
        /// Quit powerpoint application
        /// </summary>
        public void quit()
        {
            pp_app.Quit();
        }

        /// <summary>
        /// Goto next slide
        /// </summary>
        public void next_slide()
        {
            goto_slide(current_slide + 1);
        }

        /// <summary>
        /// Goto previous slide
        /// </summary>
        public void prev_slide()
        {
            goto_slide(current_slide - 1);
        }

        public void advance(int i)
        {
            goto_slide(current_slide + i);
        }

        /// <summary>
        /// Goto slide by index
        /// </summary>
        /// <param name="index"></param>
        private void goto_slide(int index)
        {
            if (index > num_slides || index < 1)
            {
                return;
                ///throw new ArgumentOutOfRangeException();
            }
            int list_index = index - 1;

            button_list.close_button.IsEnabled = slide_list[list_index].close_button;
            button_list.next_slide_button.IsEnabled = slide_list[list_index].next_slide_button;
            button_list.prev_slide_button.IsEnabled = slide_list[list_index].prev_slide_button;
            button_list.left_button.IsEnabled = slide_list[list_index].left_button;
            button_list.right_button.IsEnabled = slide_list[list_index].right_button;           
            pp_slideshow.GotoSlide(index);
            current_slide = index;

        }


    }

    /// <summary>
    /// Tracks which buttons are enabled in the main window by slide
    /// </summary>
    class Copernicus_slide
    {
        public bool next_slide_button { get; set; } = false;
        public bool prev_slide_button { get; set; } = false;
        public bool close_button { get; set; } = false;
        public bool left_button { get; set; } = false;
        public bool right_button { get; set; } = false;
        public string id { get; set; } = "";
    }

    /// <summary>
    /// List of main window buttons
    /// </summary>
    class Button_list
    {
        public Microsoft.Kinect.Toolkit.Controls.KinectTileButton next_slide_button { get; set; }
        public Microsoft.Kinect.Toolkit.Controls.KinectTileButton prev_slide_button { get; set; }
        public Microsoft.Kinect.Toolkit.Controls.KinectTileButton close_button { get; set; }
        public Microsoft.Kinect.Toolkit.Controls.KinectTileButton left_button { get; set; }
        public Microsoft.Kinect.Toolkit.Controls.KinectTileButton right_button { get; set; }
    }
}
