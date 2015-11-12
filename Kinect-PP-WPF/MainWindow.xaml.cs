using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Kinect.Toolkit;
using Microsoft.Kinect;

namespace Copernicus
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private KinectSensorChooser sensorChooser;
        private KinectSlideShow slideShow;
        private PowerPointControl ppControl;

        /// <summary>
        /// Constructor for MainWindow() 
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            Loaded += OnLoaded;

            /// Open xml data
            var fileDiag = new System.Windows.Forms.OpenFileDialog();
            fileDiag.Filter = "XML File|*.xml";
            fileDiag.FilterIndex = 1;

            if (fileDiag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                slideShow = new KinectSlideShow();
                ppControl = new PowerPointControl();
                ppControl.SlideChanged += OnSlideChanged;
                slideShow.Open(fileDiag.FileName);
                ppControl.Open(slideShow.PresentationFileName);
            }
            else
            {
                Close();
            }

            /// Maximize kinect window and make it the topmost window
            WindowState = WindowState.Maximized;
            Topmost = true;


        }

        /// <summary>
        /// Event handler for Loaded
        /// Initializes and starts the sensor chooser
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="routedEventArgs"></param>
        private void OnLoaded(object sender, RoutedEventArgs routedEventArgs)
        {
            sensorChooser = new KinectSensorChooser();
            sensorChooser.KinectChanged += SensorChooserOnKinectChanged;
            sensorChooserUi.KinectSensorChooser = sensorChooser;
            sensorChooser.Start();
        }

        /// <summary>
        /// Event handler for changes to connected Kinects
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void SensorChooserOnKinectChanged(object sender, KinectChangedEventArgs args)
        {
            bool error = false;
            if (args.OldSensor != null)
            {
                try
                {
                    args.OldSensor.DepthStream.Range = DepthRange.Default;
                    args.OldSensor.SkeletonStream.EnableTrackingInNearRange = false;
                    args.OldSensor.DepthStream.Disable();
                    args.OldSensor.SkeletonStream.Disable();
                }
                catch (InvalidOperationException)
                {
                    // KinectSensor might enter an invalid state while enabling/disabling streams or stream features.
                    // E.g.: sensor might be abruptly unplugged.
                    error = true;
                }
            }

            if (args.NewSensor != null)
            {
                try
                {
                    args.NewSensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30);
                    args.NewSensor.SkeletonStream.Enable();

                    try
                    {
                        args.NewSensor.DepthStream.Range = DepthRange.Near;
                        args.NewSensor.SkeletonStream.EnableTrackingInNearRange = true;
                        args.NewSensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                    }
                    catch (InvalidOperationException)
                    {
                        // Non Kinect for Windows devices do not support Near mode, so reset back to default mode.
                        args.NewSensor.DepthStream.Range = DepthRange.Default;
                        args.NewSensor.SkeletonStream.EnableTrackingInNearRange = false;
                    }
                }
                catch (InvalidOperationException)
                {
                    error = true;
                    // KinectSensor might enter an invalid state while enabling/disabling streams or stream features.
                    // E.g.: sensor might be abruptly unplugged.
                }
            }
            if (!error)
            {
                kinectRegion.KinectSensor = args.NewSensor;
                kinectRegion.KinectSensor.SkeletonFrameReady += OnSkeletonFrameReady;
            }


        }

        /// <summary>
        /// Handler for SkeletonFrameReady
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnSkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs args)
        {

        }

        private void OnSlideChanged(object sender, SlideChangedEventArgs args)
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(controlGrid); ++i)
            {
                DependencyObject child = VisualTreeHelper.GetChild(controlGrid, i);
                if (child != null && child is Microsoft.Kinect.Toolkit.Controls.KinectTileButton)
                {
                    Microsoft.Kinect.Toolkit.Controls.KinectTileButton control = (child as Microsoft.Kinect.Toolkit.Controls.KinectTileButton);

                    if (slideShow.slides[args.SlideNumber - 1].Buttons.Contains(control.Name))                 
                    {
                        control.IsEnabled = true;
                    }
                    else
                    {
                        control.IsEnabled = false;
                    }
                }

            }
        }


        /// <summary>
        /// Next page button event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void NextButtonOnClick(object sender, RoutedEventArgs args)
        {
            ppControl.NextSlide();
        }

        /// <summary>
        /// Previous page button event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void PreviousButtonOnClick(object sender, RoutedEventArgs args)
        {
            ppControl.PreviousSlide();
        }

        /// <summary>
        /// Close button event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void CloseButtonOnClick(object sender, RoutedEventArgs args)
        {
            ppControl.Close();
            Close();

        }

        /// <summary>
        /// Right button event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void RightButtonOnClick(object sender, RoutedEventArgs args)
        {
            try
            {
                ppControl.GotoSlide(ppControl.CurrentSlide + 5);
            }
            catch (ArgumentOutOfRangeException)
            {
                /// do nothing
            }

        }

        /// <summary>
        /// Left button event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void LeftButtonOnClick(object sender, RoutedEventArgs args)
        {
            try
            {
                ppControl.GotoSlide(ppControl.CurrentSlide - 5);
            }
            catch (ArgumentOutOfRangeException)
            {
                /// do nothing
            }
        }
    }
}

