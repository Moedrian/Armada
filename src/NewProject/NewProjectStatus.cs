using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace NewProject;

public class NewProjectStatus : INotifyPropertyChanged
{
    private bool _isBackEnabled;
    public bool IsBackEnabled
    {
        get => _isBackEnabled;
        set => SetField(ref _isBackEnabled, value);
    }

    private bool _isNextEnabled;
    public bool IsNextEnabled
    {
        get => _isNextEnabled;
        set => SetField(ref _isNextEnabled, value);
    }

    private Visibility _backVisibility = Visibility.Visible;

    public Visibility BackVisibility
    {
        get => _backVisibility;
        set => SetField(ref _backVisibility, value);
    }

    private Visibility _nextVisibility = Visibility.Visible;

    public Visibility NextVisibility
    {
        get => _nextVisibility;
        set => SetField(ref _nextVisibility, value);
    }

    private string _nextButtonText = "Next";

    public string NextButtonText
    {
        get => _nextButtonText;
        set => SetField(ref _nextButtonText, value);
    }

    public string SelectedDisk { get; set; } = string.Empty;
    public string SelectedProject { get; set; } = string.Empty;

    public string[] SkippedPrefixes { get; set; } = [];
    public string[] SkippedKeywords { get; set; } = [];
    public string[] SkippedSuffixes { get; set; } = [];

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}