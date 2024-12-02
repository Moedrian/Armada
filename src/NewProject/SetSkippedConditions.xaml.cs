using System;
using System.Windows;
using System.Windows.Controls;

namespace NewProject;

/// <summary>
/// Interaction logic for SetSkippedConditions.xaml
/// </summary>
public partial class SetSkippedConditions : Page
{
    public SetSkippedConditions()
    {
        InitializeComponent();

        // reset
        // MainWindow.Status.SkippedPrefixes = [];
        // MainWindow.Status.SkippedKeywords = [];
        // MainWindow.Status.SkippedSuffixes = [];

        MainWindow.Status.BackVisibility = Visibility.Visible;
        MainWindow.BackButtonUri = MainWindow.SelectProjectUri;

        MainWindow.Status.NextButtonText = MainWindow.NextText;
        MainWindow.Status.IsNextEnabled = false;

        CreateProject.CheckComponentCollection.Clear();

        ConfirmButton.Click += delegate
        {
            MainWindow.Status.SkippedPrefixes = Split(SkippedPrefixesBoxes.Text);
            MainWindow.Status.SkippedKeywords = Split(SkippedKeywordsBoxes.Text);
            MainWindow.Status.SkippedSuffixes = Split(SkippedSuffixesBoxes.Text);

            MainWindow.Status.IsNextEnabled = true;
            MainWindow.NextButtonUri = MainWindow.CreateProjectUri;
        };
    }

    private static string[] Split(string input)
    {
        return input.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    }
}