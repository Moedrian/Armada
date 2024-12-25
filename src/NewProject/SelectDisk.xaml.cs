using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Path = System.IO.Path;

namespace NewProject;

/// <summary>
/// Interaction logic for SelectDisk.xaml
/// </summary>
public partial class SelectDisk : UserControl
{
    private static readonly ObservableCollection<string> ProjectDisks = [];

    private static void AllowOnlyDigits(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox tb)
        {
            tb.Text = Regex.Replace(tb.Text, @"[^\d]", string.Empty);
            tb.CaretIndex = tb.Text.Length;
        }
    }

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
        
        RatioBox.Text = "9";
        BorderWidthBox.Text = "60";
        SimilarityRateBox.Text = "10";

        RatioBox.TextChanged += AllowOnlyDigits;
        BorderWidthBox.TextChanged += AllowOnlyDigits;
        SimilarityRateBox.TextChanged += AllowOnlyDigits;

        DiskListComboBox.DropDownClosed += (_, _) =>
        {
            var disk = (string)DiskListComboBox.SelectedItem;
            if (disk is { Length: > 0 })
            {
                try
                {
                    MainWindow.Status.SelectedDisk = disk;

                    if (int.TryParse(RatioBox.Text, out var ratio))
                        MainWindow.Status.PixelRatio = ratio;
                    else
                        throw new Exception("Ratio format not correct!");

                    if (int.TryParse(BorderWidthBox.Text, out var border))
                        MainWindow.Status.BorderWidth = border;
                    else
                        throw new Exception("Border width format not correct!");

                    if (int.TryParse(SimilarityRateBox.Text, out var rate))
                        MainWindow.Status.SimilarityRate = rate;
                    else
                        throw new Exception("Similarity format not correct!");

                    MainWindow.Status.IsNextEnabled = true;
                    MainWindow.NextButtonUri = MainWindow.SelectProjectUri;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        };
    }
}