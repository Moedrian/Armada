using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using NewProject.Util.PartList;

namespace NewProject;

/// <summary>
/// Interaction logic for CreateProject.xaml
/// </summary>
public partial class CreateProject : UserControl
{
    public static readonly ObservableCollection<CheckPart> CheckComponentCollection = new();

    public CreateProject()
    {
        InitializeComponent();

        MainWindow.BackButtonUri = MainWindow.SkippedConditionsUri;
        MainWindow.Status.NextButtonText = MainWindow.FinishText;

        Task.Run(delegate
        {
            Dispatcher.Invoke(delegate
            {
                var prj = MainWindow.Status.SelectedProject;
                var components = PartListGetter.GetTopSideComponents(LeonardoType.FlyingProbes, prj, out var bomExists);

                if (components.Length == 0)
                {
                    const string message = "Unable to get component list.";
                    MessageBox.Show(message, "Notice", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    if (!bomExists)
                    {
                        const string message =
                            "BOM file does not existing, displaying all parts.";
                        MessageBox.Show(message, "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    }

                    var prefixes = MainWindow.Status.SkippedPrefixes;
                    var keywords = MainWindow.Status.SkippedKeywords;
                    var suffixes = MainWindow.Status.SkippedSuffixes;

                    var ccl = GetCandidates(components, prefixes, suffixes, keywords).OrderBy(c => c.PartName);
                    foreach (var cc in ccl)
                        CheckComponentCollection.Add(cc);

                    ComponentList.ItemsSource = CheckComponentCollection;

                    AttachEvents();
                }
            });
        });
    }

    private void AttachEvents()
    {
        CheckAllBox.Checked += delegate
        {
            if (CheckComponentCollection.Count == 0) return;

            foreach (var part in CheckComponentCollection)
                part.IsChecked = true;

            ComponentList.Items.Refresh();
        };

        CheckAllBox.Unchecked += delegate
        {
            if (CheckComponentCollection.Count == 0) return;

            foreach (var part in CheckComponentCollection)
                part.IsChecked = false;

            ComponentList.Items.Refresh();
        };
    }

    private const StringComparison StrCmp = StringComparison.Ordinal;

    private static bool MatchPrefixes(string input, string[] prefixes)
    {
        return prefixes.Any(str => input.StartsWith(str, StrCmp));
    }

    private static bool MatchSuffixes(string input, string[] suffixes)
    {
        return suffixes.Any(str => input.EndsWith(str, StrCmp));
    }
    private static bool MatchKeywords(string input, string[] prefixes, string[] suffixes, string[] keywords)
    {
        return keywords.Any(str => input.Contains(str, StrCmp))
               && !prefixes.Any(str => input.StartsWith(str, StrCmp))
               && !suffixes.Any(str => input.EndsWith(str, StrCmp));
    }

    private static CheckPart[] GetCandidates(Part[] components, string[] prefixes, string[] suffixes, string[] keywords)
    {
        var noKeywords = (from c in components where !MatchKeywords(c.DrawingReference, prefixes, suffixes, keywords) select c);
        var noPrefixes = (from c in noKeywords where !MatchPrefixes(c.DrawingReference, prefixes) select c);
        var noSuffixes = (from c in noPrefixes where !MatchSuffixes(c.DrawingReference, suffixes) select c);
        return noSuffixes.Select(c => new CheckPart(c)).ToArray();
    }
}

