using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Path = System.IO.Path;

namespace NewProject;

/// <summary>
/// Interaction logic for SelectDisk.xaml
/// </summary>
public partial class SelectDisk : UserControl
{
    private static readonly ObservableCollection<string> ProjectDisks = new();

    public SelectDisk()
    {
        var drives = DriveInfo.GetDrives()
            .Where(driveInfo => driveInfo.DriveType is DriveType.Fixed or DriveType.Network);

        ProjectDisks.Clear();

        foreach (var drive in drives)
        {
            const string mark = "LeonardoOS2\\FlyTprj.100";
            var fullname = drive.RootDirectory.FullName;
            var projectsDir = Path.Combine(fullname, mark);
            if (Directory.Exists(projectsDir))
                ProjectDisks.Add(fullname);
        }

        InitializeComponent();

        // reset
        DiskListComboBox.Items.Clear();
        MainWindow.Status.BackVisibility = Visibility.Hidden;
        MainWindow.Status.IsBackEnabled = false;
        MainWindow.Status.IsNextEnabled = false;
        MainWindow.Status.SelectedDisk = string.Empty;
        MainWindow.Status.SelectedProject = string.Empty;
        MainWindow.Status.SkippedPrefixes = [];
        MainWindow.Status.SkippedKeywords = [];
        MainWindow.Status.SkippedSuffixes = [];

        // register
        DiskListComboBox.ItemsSource = ProjectDisks;
        DiskListComboBox.DropDownClosed += (_, _) =>
        {
            var disk = (string)DiskListComboBox.SelectedItem;
            if (disk is { Length: > 0 })
            {
                MainWindow.Status.SelectedDisk = disk;
                MainWindow.Status.IsNextEnabled = true;
                MainWindow.NextButtonUri = MainWindow.SelectProjectUri;
            }
        };
    }
}