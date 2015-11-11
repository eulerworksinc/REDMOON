using System;
using Microsoft.Office.Interop.PowerPoint;

namespace Copernicus
{
    class PowerPointControl
    {
        private Application ppApp;
        private Presentation ppPresentation;
        private SlideShowView ppSlideShow;

        /// <summary>
        /// PowerPoint filename
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// Current slide number
        /// </summary>
        public int CurrentSlide { get; private set; }

        /// <summary>
        /// Number of slides in presentation
        /// </summary>
        public int SlideCount { get; private set; }

        /// <summary>
        /// True if presentation is open
        /// </summary>
        public bool IsOpen { get; private set; }
        
        /// <summary>
        /// Constructor for PowerPointControl
        /// </summary>
        /// <param name="fileName"></param>
        public PowerPointControl()
        {
            IsOpen = false;
            SlideCount = 0;
            CurrentSlide = 1;
        }

        /// <summary>
        /// Open a PowerPoint presentation
        /// </summary>
        /// <param name="fileName"></param>
        public void Open(string fileName)
        {
            if (IsOpen)
            {
                throw new InvalidOperationException();
            }

            try
            {
                ppApp = new Application();
                ppApp.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            }
            catch (Exception)
            {
                throw;
            }

            try
            {
                ppPresentation = ppApp.Presentations.Open(fileName);
                ppPresentation.SlideShowSettings.ShowPresenterView = Microsoft.Office.Core.MsoTriState.msoFalse;
                ppPresentation.SlideShowSettings.Run();
                ppSlideShow = ppPresentation.SlideShowWindow.View;
            }
            catch (Exception)
            {
                ppApp.Quit();
                ppSlideShow = null;
                ppPresentation = null;
                ppApp = null;
                throw;
            }

            CurrentSlide = 1;
            SlideCount = ppPresentation.Slides.Count;
            FileName = fileName;
            IsOpen = true;
            SlideChangedEventArgs args = new SlideChangedEventArgs();
            args.SlideNumber = CurrentSlide;
            OnSlideChanged(args);
        }

        /// <summary>
        /// Goto slide by index
        /// </summary>
        /// <param name="index"></param>
        public void GotoSlide(int index)
        {
            
            if (!IsOpen)
            {
                throw new InvalidOperationException();
            }

            if (index > SlideCount || index < 1)
            {
                throw new ArgumentOutOfRangeException();
            }

            if (index != CurrentSlide)
            {
                ppSlideShow.GotoSlide(index);
                CurrentSlide = index;
                SlideChangedEventArgs args = new SlideChangedEventArgs();
                args.SlideNumber = CurrentSlide;
                OnSlideChanged(args);
            }

        }

        /// <summary>
        /// Goto next slide
        /// </summary>
        public void NextSlide()
        {
            try
            {
                GotoSlide(CurrentSlide + 1);
            }
            catch (ArgumentOutOfRangeException)
            {
                /// Do nothing
            }
        }

        /// <summary>
        /// Goto previous slide
        /// </summary>
        public void PreviousSlide()
        {
            try
            {
                GotoSlide(CurrentSlide - 1);
            }
            catch (ArgumentOutOfRangeException)
            {
                /// Do nothing
            }
        }

        /// <summary>
        /// Close presentation
        /// </summary>
        public void Close()
        {
            if (!IsOpen)
            {
                throw new InvalidOperationException();
            }

            ppApp.Quit();
            ppSlideShow = null;
            ppPresentation = null;
            ppApp = null;
            IsOpen = false;
        }
        
        /// <summary>
        /// Raise the SlideChanged event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnSlideChanged(SlideChangedEventArgs e)
        {
            EventHandler<SlideChangedEventArgs> handler = SlideChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        /// <summary>
        /// Raised when the slide is changed
        /// </summary>
        public event EventHandler<SlideChangedEventArgs> SlideChanged;
    }

    /// <summary>
    /// Event arguments for SlideChanged event
    /// </summary>
    public class SlideChangedEventArgs : EventArgs
    {
        public int SlideNumber { get; set; }
    }
}
