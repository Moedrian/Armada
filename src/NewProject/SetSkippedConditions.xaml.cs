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
        MainWindow.Status.IsNextEnabled = false;

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