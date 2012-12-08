using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.Timers;
using System.Runtime.InteropServices;


namespace SPprotoype
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Member Variables
        private KinectSensor _Kinect;
        private readonly Brush[] _SkeletonBrushes;
        private Skeleton[] _FrameSkeletons;
        string pptname = "D:\\acads\\1stsemSY12-13\\CMSC190-1\\SP\\prototype\\SPprotoype\\sppresentation.pptx";
        PowerPoint.Application ppt;
        PowerPoint.Presentations spPt;
        PowerPoint.Presentation openPt;
        bool opened = false;
        int crt = 1;
        int rx = 0 , ry = 0 , oldry = 0;
        static Timer _timer = new Timer(3000);
        #endregion Member Variables

        #region Constructor
        public MainWindow()
        {
            InitializeComponent();
            this._SkeletonBrushes = new[] { Brushes.White,Brushes.Crimson, Brushes.Indigo, Brushes.DodgerBlue, Brushes.Purple, Brushes.Pink };
            DiscoverKinectSensor();
            this.Loaded += (s, e) => { DiscoverKinectSensor(); };
            this.Unloaded += (s, e) => { this.Kinect = null; };
        }

        #endregion Constructor

        #region Methods
        private void DiscoverKinectSensor()
        {
            KinectSensor.KinectSensors.StatusChanged += KinectSensors_StatusChanged;
            this.Kinect = KinectSensor.KinectSensors.FirstOrDefault(x => x.Status == KinectStatus.Connected);
        }
        private void KinectSensors_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            switch (e.Status)
            {
                case KinectStatus.Connected:
                    if (this.Kinect == null)
                    {
                        this.Kinect = e.Sensor;
                        kinectDetection.Text = "Kinect Detected";
                    }
                    break;
                case KinectStatus.Disconnected:
                    if (this.Kinect == e.Sensor)
                    {
                        this.Kinect = null;
                        this.Kinect = KinectSensor.KinectSensors.FirstOrDefault(x => x.Status == KinectStatus.Connected);
                        if (this.Kinect == null)
                        {
                            kinectDetection.Text = "Kinect Disconnected";
                        }
                    }
                    break;
            }
        }
        #endregion Methods

        #region Properties
        public KinectSensor Kinect
        {
            get { return this._Kinect; }

            set
            {
                if (this._Kinect != value)
                {
                    //uninitialize
                    if (this._Kinect != null)
                    {
                        this._Kinect.Stop();
                        this._Kinect.SkeletonFrameReady -= KinectDevice_SkeletonFrameReady;
                        this._Kinect.SkeletonStream.Disable();
                        this._FrameSkeletons = null;
                        UninitializeKinectSensor(this._Kinect);
                    }
                    this._Kinect = value;
                    //initialize
                    if (value != null && value.Status == KinectStatus.Connected)
                    {
                        this._Kinect.SkeletonStream.Enable();
                        this._FrameSkeletons = new Skeleton[this._Kinect.SkeletonStream.FrameSkeletonArrayLength];
                        this._Kinect.SkeletonFrameReady += KinectDevice_SkeletonFrameReady;
                        InitializeKinectSensor(this._Kinect);
                    }
                }
            }
        }

        private void InitializeKinectSensor(KinectSensor sensor)
        {
            if (sensor != null)
            {
               // sensor.ColorStream.Enable();
               // sensor.ColorFrameReady += Kinect_ColorFrameReady;
                sensor.Start();
            }
        }
        private void UninitializeKinectSensor(KinectSensor sensor)
        {
            if (sensor != null)
            {
                sensor.Stop();
                //sensor.ColorFrameReady -= Kinect_ColorFrameReady;
            }
        }

        private void Kinect_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame frame = e.OpenColorImageFrame())
            {
                if (frame != null)
                {
                    byte[] pixelData = new byte[frame.PixelDataLength];
                    frame.CopyPixelDataTo(pixelData);
                    ColorImageElement.Source = BitmapImage.Create(frame.Width, frame.Height, 96, 96, PixelFormats.Bgr32, null, pixelData, frame.Width * frame.BytesPerPixel);
                }
            }
        }
        #endregion Properties

        private void KinectDevice_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
 
            using(SkeletonFrame frame = e.OpenSkeletonFrame())
            {
                if (frame != null)
                {
                    Polyline figure;
                    Brush userBrush;
                    Skeleton skeleton;
                    JointType[] joints;

                    LayoutRoot.Children.Clear();
                    frame.CopySkeletonDataTo(this._FrameSkeletons);

                    for (int i = 0; i < this._FrameSkeletons.Length; i++)
                    {
                        skeleton = this._FrameSkeletons[i];

                        if (skeleton.TrackingState == SkeletonTrackingState.Tracked)
                        {
                            userBrush = this._SkeletonBrushes[i % this._SkeletonBrushes.Length];

                            //Draws the skeleton's head and torso
                            joints = new[] { JointType.Head, JointType.ShoulderCenter, JointType.ShoulderLeft, JointType.Spine, JointType.ShoulderRight,
                                              JointType.ShoulderCenter, JointType.HipCenter, JointType.HipLeft, JointType.Spine, JointType.HipRight,
                                             JointType.HipCenter};
                            figure = CreateFigure(skeleton, userBrush, joints);
                            LayoutRoot.Children.Add(figure);
                            //Draws the skeleton's left leg
                            joints = new[] { JointType.HipLeft,JointType.KneeLeft, JointType.AnkleLeft, 
                                             JointType.FootLeft};
                            figure = CreateFigure(skeleton, userBrush, joints);
                            LayoutRoot.Children.Add(figure);

                            //Draws the skeleton's right leg
                            joints = new[] {JointType.HipRight,JointType.KneeRight,
                                            JointType.AnkleRight, JointType.FootRight};
                            figure = CreateFigure(skeleton, userBrush, joints);
                            LayoutRoot.Children.Add(figure);

                            //Draws the skeleton's left arm
                            joints = new[] {JointType.ShoulderLeft, JointType.ElbowLeft,
                                            JointType.WristLeft, JointType.HandLeft};
                            figure = CreateFigure(skeleton, userBrush, joints);
                            LayoutRoot.Children.Add(figure);

                            //Draws the skeleton's right arm
                            joints = new[] {JointType.ShoulderRight, JointType.ElbowRight,
                                            JointType.WristRight, JointType.HandRight};
                            figure = CreateFigure(skeleton, userBrush, joints);
                            LayoutRoot.Children.Add(figure);
                            Joint rhand = skeleton.Joints[JointType.HandRight];
                            rx = (int)Math.Ceiling(rhand.Position.X*1000);
                            ry = (int)Math.Ceiling(rhand.Position.Y*1000) * -1;
                            if (ry < oldry)
                            {
                                ry = ry / 2;
                            }
                            else
                            {
                                ry = ry * 2;
                            }
                            textBox1.Text = rhand.Position.X.ToString();
                            textBox2.Text = rhand.Position.Y.ToString();
                            SetCursorPos(rx+100,ry+100);
                            controlPpt(skeleton.Joints[JointType.Head],
                                          skeleton.Joints[JointType.HandLeft], skeleton.Joints[JointType.HandRight], skeleton.Joints[JointType.HipCenter]);
                            oldry = ry;
                        }
                    }

                }
            }
        } // endof KinectDevice_SkeletonFrameready

        private Polyline CreateFigure(Skeleton skeleton, Brush brush, JointType[] joints)
        {
            Polyline figure = new Polyline();
            figure.StrokeThickness = 8;
            figure.Stroke = brush;

            for (int i = 0; i < joints.Length; i++ )
            {
                figure.Points.Add(GetJointPoint(skeleton.Joints[joints[i]]));
            }
            return figure;
        }//end CreateFigure

        private Point GetJointPoint(Joint joint)
        {
            DepthImagePoint point = this.Kinect.MapSkeletonPointToDepth(joint.Position, this.Kinect.DepthStream.Format);

            point.X *= (int)this.LayoutRoot.ActualWidth / this.Kinect.DepthStream.FrameWidth;
            point.Y *= (int)this.LayoutRoot.ActualHeight / this.Kinect.DepthStream.FrameHeight;

            return new Point(point.X,point.Y);
        }//end GetJointPoint

        #region Commands
        void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            crt = 0;
        }
        void controlPpt(Joint head, Joint lhand, Joint rhand, Joint hcenter)
        {
            bool lActive = false;
            bool rActive = false;
            if ((head.Position.X > lhand.Position.X) && (head.Position.Y < lhand.Position.Y) && !lActive)
            {
                if(opened)
                    if (openPt.SlideShowWindow.View.Slide.SlideIndex > 1)
                    {
                        _timer.Elapsed += new ElapsedEventHandler(_timer_Elapsed);
                        _timer.Enabled = true;
                        if (crt == 0)
                        {
                            openPt.SlideShowWindow.View.Previous();
                            lActive = true;
                            rActive = false;
                            crt = 1;
                        }
                    }
            }

            if ((head.Position.X < rhand.Position.X) && (head.Position.Y < rhand.Position.Y) && !rActive)
            {
                if (opened)
                    if (openPt.SlideShowWindow.View.Slide.SlideIndex < 6)
                    {
                        _timer.Elapsed += new ElapsedEventHandler(_timer_Elapsed);
                        _timer.Enabled = true; // Enable it
                        if (crt == 0)
                        {
                            openPt.SlideShowWindow.View.Next();
                            rActive = true;
                            lActive = false;
                            crt = 1;
                        }
                    }
            }

            if ((hcenter.Position.X > rhand.Position.X) && (hcenter.Position.X < lhand.Position.X) && (hcenter.Position.Y < lhand.Position.Y) && (hcenter.Position.Y < rhand.Position.Y))
            {
                if (opened)
                {
                    openPt.Close();
                    ppt.Quit();
                    opened = false;
                }
            }
            if((head.Position.Y < rhand.Position.Y) && (head.Position.Y < lhand.Position.Y))
            {
                if (!opened)
                {
                    openPPt();
                    opened = true;
                }
            }
        }// end controlPpt
        void openPPt()
        {
            ppt = new PowerPoint.Application();
            ppt.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

            spPt = ppt.Presentations;
            openPt = spPt.Open(pptname, MsoTriState.msoFalse, MsoTriState.msoFalse);
            openPt.SlideShowSettings.Run();
        }//end openPPT
        //setmouseposition
        [DllImport("User32.dll")]
        private static extern bool SetCursorPos(int x, int y);
        //endsetmouseposition
        #endregion Commands
    }
}
