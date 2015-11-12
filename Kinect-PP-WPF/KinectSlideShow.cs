using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace Copernicus
{

    /// <summary>
    /// List of Kinect slides
    /// </summary>
    class KinectSlideShow
    {
        public List<KinectSlide> slides { get; private set; } = new List<KinectSlide>();

        /// <summary>
        /// True if a slide show has been opened
        /// </summary>
        public bool IsOpen { get; private set; } = false;

        /// <summary>
        /// PowerPoint presentation file name
        /// </summary>
        public string PresentationFileName { get; private set; }

        public void Open(string fileName)
        {
            if (IsOpen)
            {
                throw new InvalidOperationException();
            }

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(fileName);
                XmlElement root = xmlDoc.DocumentElement;
                PresentationFileName = Path.GetDirectoryName(fileName);
                PresentationFileName += "\\" + root.GetAttribute("filename");

                foreach (XmlNode node in root.ChildNodes)
                {
                    KinectSlide slide = new KinectSlide();
                    slide.Name = node.Attributes["id"].Value;
                    foreach (XmlNode button in node.ChildNodes)
                    {
                        slide.Buttons.Add(button.Attributes["id"].Value);
                    }
                    slides.Add(slide);
                }
            }

            catch (Exception)
            {
                Close();
                throw;
            }

            IsOpen = true;
        }

        public void Close()
        {
            PresentationFileName = null;
            slides = new List<KinectSlide>();
            IsOpen = false;
        }
    }

    /// <summary>
    /// Name and interface of a single slide
    /// </summary>
    class KinectSlide
    {
        /// <summary>
        /// List of button names in the slide
        /// </summary>
        public List<string> Buttons { get; set; } = new List<string>();

        /// <summary>
        /// Slide name
        /// </summary>
        public string Name { get; set; }
    }
}
