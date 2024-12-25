using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Xml.Serialization;
using ArmadaPack;
using NewProject.Util.LeonardoProperty;
using NewProject.Util.PartList;
using NewProject.Util.TpList;

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
            try
            {
                if (Status.NextButtonText == FinishText)
                {
                    var sites = LeoProp.GetNumberOfBoards(LeonardoType.FlyingProbes, Status.SelectedProject);
                    var testPoints = TestPoint.GetAllTestPoints(LeonardoType.FlyingProbes, Status.SelectedProject);

                    var testGroups = new TestPartsGroup[sites];
                    for (var i = 0; i < sites; i++)
                    {
                        var group = new TestPartsGroup { Site = i + 1 };

                        var cl = new List<TestPart>();
                        foreach (var item in CreateProject.CheckComponentCollection)
                        {
                            if (item is { IsChecked: true, Part: not null })
                            {
                                if (item.Part.TryCalculateBarycenterCoordinates(testPoints, i + 1, out var co))
                                {
                                    cl.Add(new TestPart
                                    {
                                        DrawingReference = item.PartName,
                                        CoordinateX = co.X,
                                        CoordinateY = co.Y,
                                        Edge = co.Edge
                                    });
                                }
                            }
                        }

                        if (cl.Count == 0)
                            throw new Exception("Part Test List can not be empty.");

                        group.TestParts = cl.ToArray();
                        testGroups[i] = group;
                    }

                    var vision = new Vision
                    {
                        PixelRatio = Status.PixelRatio,
                        BorderWidthPixels = Status.BorderWidth,
                        SimilarityRate = Status.SimilarityRate,
                        Groups = testGroups
                    };

                    var fullPath = Status.SelectedProject[..3] + Hierarchy + Status.SelectedProject[3..];
                    var ftpDir = Path.Combine(fullPath, "PROGRAM\\FTP");

                    // create vision\vision.xml
                    var dir = Directory.CreateDirectory(Path.Combine(ftpDir, "Vision"));
                    var file = Path.Combine(dir.FullName, "vision.xml");
                    var sr = new XmlSerializer(typeof(Vision));
                    using var sw = new StreamWriter(file);
                    sr.Serialize(sw, vision);
                   
                    // copy templates
                    var templateDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TPGM_TEMPLATES");
                    var psi = new ProcessStartInfo("xcopy", $"\"{templateDir}\" \"{ftpDir}\" /Q /E /Y")
                    {
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    Process.Start(psi)?.WaitForExit();

                    MessageBox.Show("Project file created, now you can safely close this window.", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    PresentationFrame.Source = NextButtonUri;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        };
    }
}