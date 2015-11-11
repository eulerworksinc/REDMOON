using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Kinect;
using Microsoft.Kinect.Toolkit;


namespace Copernicus
{
    /// <summary>
    /// State machine class for Copernicus
    /// </summary>
    class Copernicus_sm
    {
        private List<Copernicus_slide> slide_list;
        private string pp_filename;
        private Button_list button_list;
        private PowerPointControl ppControl;

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

            ppControl = new PowerPointControl();
            ppControl.Open(pp_filename);
        }

        /// <summary>
        /// Quit powerpoint application
        /// </summary>
        public void quit()
        {
            ppControl.Close();
        }

        /// <summary>
        /// Goto next slide
        /// </summary>
        public void next_slide()
        {
            ppControl.NextSlide();
        }

        /// <summary>
        /// Goto previous slide
        /// </summary>
        public void prev_slide()
        {
            ppControl.PreviousSlide();
        }

        public void advance(int i)
        {
            ppControl.GotoSlide(i);
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
