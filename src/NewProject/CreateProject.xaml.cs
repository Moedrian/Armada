using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using NewProject.Serializable;
using NewProject.Util.PartList;

namespace NewProject;

/// <summary>
/// Interaction logic for CreateProject.xaml
/// </summary>
public partial class CreateProject : UserControl
{
    public CreateProject()
    {
        InitializeComponent();
        MainWindow.BackButtonUri = MainWindow.SkippedConditionsUri;
        MainWindow.Status.NextVisibility = Visibility.Hidden;

        var prj = MainWindow.Status.SelectedProject;
        var components = PartListGetter.GetTopSideComponents(LeonardoType.FlyingProbes, prj, out var bomExists);

        if (!bomExists)
        {
            const string message =
                "BOM file does not existing, displaying all parts.";
            MessageBox.Show(message, "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        var prefixes = MainWindow.Status.SkippedPrefixes;
        var keywords = MainWindow.Status.SkippedKeywords;
        var suffixes = MainWindow.Status.SkippedSuffixes;

        ComponentList.ItemsSource = GetCandidates(components, prefixes, suffixes, keywords);

        CreateButton.Click += delegate
        {
            var cl = new List<Component?>();
            foreach (var item in ComponentList.Items)
            {
                var c = item as CheckComponent;
                if  (c?.IsChecked ?? false)
                   cl.Add(c.Component); 
            }

            var valid = cl.Select(c => c != null).ToList();
            if (valid.Count == 0)
            {
                MessageBox.Show("Component Test List can not be empty.", "Caution", MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }


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

    private static CheckComponent[] GetCandidates(Component[] components, string[] prefixes, string[] suffixes, string[] keywords)
    {
        var noKeywords = (from c in components where !MatchKeywords(c.DrawingReference, prefixes, suffixes, keywords) select c);
        var noPrefixes = (from c in noKeywords where !MatchPrefixes(c.DrawingReference, prefixes) select c);
        var noSuffixes = (from c in noPrefixes where !MatchSuffixes(c.DrawingReference, suffixes) select c);
        return noSuffixes.Select(c => new CheckComponent(c)).ToArray();
    }
}

