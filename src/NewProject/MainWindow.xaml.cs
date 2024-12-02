using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Xml.Serialization;
using ArmadaPack;

namespace NewProject;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public static readonly NewProjectStatus Status = new();

    public static readonly Uri SelectDiskUri = new("pack://application:,,,/SelectDisk.xaml", UriKind.Absolute);
    public static readonly Uri SelectProjectUri = new("pack://application:,,,/SelectProject.xaml", UriKind.Absolute);
    public static readonly Uri CreateProjectUri = new("pack://application:,,,/CreateProject.xaml", UriKind.Absolute);
    public static readonly Uri SkippedConditionsUri = new("pack://application:,,,/SetSkippedConditions.xaml", UriKind.Absolute);

    public const string Hierarchy = @"LeonardoOS2\FlyTprj.100\PRJ\";
    public const string NextText = "Next";
    public const string FinishText = "Finish";

    public static Uri? NextButtonUri { get; set; }
    public static Uri? BackButtonUri { get; set; }

    public MainWindow()
    {
        InitializeComponent();

        DataContext = Status;

        PresentationFrame.Source = SelectDiskUri;

        BackButton.Click += delegate
        {
            PresentationFrame.Source = BackButtonUri;
        };

        NextButton.Click += delegate
        {
            if (Status.NextButtonText == FinishText)
            {
                var cl = new List<TestPart>();
                foreach (var item in CreateProject.CheckComponentCollection)
                {
                    if (item is { IsChecked: true, Part: not null })
                        cl.Add(new TestPart
                        {
                            DrawingReference = item.PartName,
                            BarycenterX = item.Barycenter.X,
                            BarycenterY = item.Barycenter.Y,
                            CoordinateX = item.Coordinates.X,
                            CoordinateY = item.Coordinates.Y,
                            Edge = item.Coordinates.Edge
                        });
                }

                if (cl.Count == 0)
                {
                    MessageBox.Show("Part Test List can not be empty.", "Caution", MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var fullPath = Status.SelectedProject[..3] + Hierarchy + Status.SelectedProject[3..];
                var ftpDir = Path.Combine(fullPath, "PROGRAM\\FTP", "Vision");
                var dir = Directory.CreateDirectory(ftpDir);
                var file = Path.Combine(dir.FullName, "vision.xml");
                var sr = new XmlSerializer(typeof(Vision));
                using var sw = new StreamWriter(file);
                sr.Serialize(sw, new Vision { TestParts = cl.ToArray() });

                MessageBox.Show("Project file created, now you can safely close this window.", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                PresentationFrame.Source = NextButtonUri;
            }
        };
    }
}