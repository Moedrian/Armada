using System;
using System.Windows;

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
            PresentationFrame.Source = NextButtonUri;
        };
    }
}