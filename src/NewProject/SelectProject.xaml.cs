using NewProject.Util.PartList;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Path = System.IO.Path;

namespace NewProject;

/// <summary>
/// Interaction logic for SelectProject.xaml
/// </summary>
public partial class SelectProject : UserControl
{
    private static readonly ObservableCollection<string> Projects = [];

    private static string[] IndexTestPrograms(string disk)
    {
        var hierarchy = Path.Combine(disk, MainWindow.Hierarchy);
        const StringComparison comp = StringComparison.OrdinalIgnoreCase;

        var list = new List<string>();
        var rootDi = new DirectoryInfo(hierarchy);
        foreach (var archive in rootDi.GetDirectories())
            foreach (var testProgram in archive.GetDirectories())
                if (!testProgram.Name.EndsWith(".zip", comp))
                    list.Add(testProgram.FullName.Replace(hierarchy, string.Empty, comp));

        return [..list];
    }

    public SelectProject()
    {
        var testPrograms = IndexTestPrograms(MainWindow.Status.SelectedDisk);

        Projects.Clear();
        foreach (var tpgm in testPrograms)
            Projects.Add(tpgm);

        InitializeComponent();

        MainWindow.Status.BackVisibility = Visibility.Visible;
        MainWindow.Status.IsBackEnabled = true;
        MainWindow.Status.IsNextEnabled = false;
        MainWindow.BackButtonUri = MainWindow.SelectDiskUri;

        // reset
        MainWindow.Status.SelectedProject = string.Empty;
        MainWindow.Status.SkippedPrefixes = [];
        MainWindow.Status.SkippedKeywords = [];
        MainWindow.Status.SkippedSuffixes = [];

        AvailableProjectList.ItemsSource = Projects;

        AvailableProjectList.SelectionChanged += (_, _) =>
        {
            MainWindow.Status.IsNextEnabled = true;

            var selection = (string)AvailableProjectList.SelectedItem;
            if (string.IsNullOrEmpty(selection))
                return;

            var disk = MainWindow.Status.SelectedDisk;
            MainWindow.Status.SelectedProject = Path.Combine(disk, selection);

            MainWindow.NextButtonUri = MainWindow.SkippedConditionsUri;
        };
    }
}