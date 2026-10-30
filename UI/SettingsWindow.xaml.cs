using System.IO;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ContentMigrator.Config;

namespace ContentMigrator.UI;

/// <summary>
/// Edits appsettings.json. The tabs and fields are generated from the [SettingsSection] and
/// [Setting] attributes on <see cref="AppSettings"/>, so a new setting only has to be added there.
/// Edits apply to a copy; nothing is written until Save.
/// </summary>
public partial class SettingsWindow : Window
{
    private sealed class FieldBinding
    {
        public object Section { get; init; }
        public PropertyInfo Property { get; init; }
        public SettingAttribute Meta { get; init; }
        public string SectionTitle { get; init; }
        public Func<object> Read { get; init; }
    }

    private readonly List<FieldBinding> _bindings = new();
    private AppSettings _working;

    public SettingsWindow(AppSettings current)
    {
        InitializeComponent();
        DarkTitleBar.Attach(this);
        lblPath.Text = SettingsStore.SettingsPath;
        _working = SettingsStore.Clone(current);
        BuildTabs();
    }

    /// <summary>The settings that were saved (valid after DialogResult == true).</summary>
    public AppSettings SavedSettings { get; private set; }

    // ======================================================================= build

    private void BuildTabs()
    {
        var selected = tabs.SelectedIndex;
        tabs.Items.Clear();
        _bindings.Clear();

        foreach (var sectionProp in typeof(AppSettings).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            var title = sectionProp.GetCustomAttribute<SettingsSectionAttribute>()?.Title ?? sectionProp.Name;
            var section = sectionProp.GetValue(_working);
            var panel = new StackPanel { Margin = new Thickness(14, 12, 14, 12) };

            foreach (var prop in sectionProp.PropertyType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                var meta = prop.GetCustomAttribute<SettingAttribute>();
                if (meta == null) continue;
                panel.Children.Add(BuildEditor(section, prop, meta, title));
            }

            tabs.Items.Add(new TabItem
            {
                Header = title,
                Content = new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }
            });
        }
        tabs.SelectedIndex = selected >= 0 && selected < tabs.Items.Count ? selected : 0;
    }

    private UIElement BuildEditor(object section, PropertyInfo prop, SettingAttribute meta, string sectionTitle)
    {
        var block = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
        var value = prop.GetValue(section);
        Func<object> read;

        if (prop.PropertyType == typeof(bool))
        {
            var check = new CheckBox { Content = meta.Label, IsChecked = (bool)value, FontWeight = FontWeights.Bold };
            block.Children.Add(check);
            read = () => check.IsChecked == true;
        }
        else
        {
            block.Children.Add(new TextBlock
            {
                Text = meta.Label + (meta.Required ? " *" : ""),
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 3)
            });

            if (prop.PropertyType == typeof(List<string>))
            {
                var box = new TextBox
                {
                    Text = string.Join(Environment.NewLine, (List<string>)value ?? new List<string>()),
                    AcceptsReturn = true,
                    Height = 130,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    FontFamily = new FontFamily("Consolas"),
                    Padding = new Thickness(4, 2, 4, 2)
                };
                block.Children.Add(box);
                read = () => box.Text
                    .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
                    .Select(s => s.Trim())
                    .Where(s => s.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else if (prop.PropertyType == typeof(int))
            {
                var box = new TextBox
                {
                    Text = ((int)value).ToString(CultureInfo.InvariantCulture),
                    Width = 120,
                    Height = 26,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(4, 2, 4, 2)
                };
                block.Children.Add(box);
                read = () => int.TryParse(box.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var i)
                    ? i
                    : throw new FormatException($"{sectionTitle} > {meta.Label} must be a whole number.");
            }
            else
            {
                var box = new TextBox
                {
                    Text = (string)value ?? "",
                    Height = 26,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(4, 2, 4, 2)
                };
                if (meta.IsFile)
                {
                    var row = new DockPanel();
                    var browse = new Button { Content = "Browse...", Width = 80, Height = 26, Margin = new Thickness(6, 0, 0, 0) };
                    browse.Click += (_, _) =>
                    {
                        var dialog = new Microsoft.Win32.OpenFileDialog
                        {
                            Title = meta.Label,
                            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
                        };
                        if (File.Exists(box.Text)) dialog.InitialDirectory = Path.GetDirectoryName(box.Text);
                        if (dialog.ShowDialog(this) == true) box.Text = dialog.FileName;
                    };
                    DockPanel.SetDock(browse, Dock.Right);
                    row.Children.Add(browse);
                    row.Children.Add(box);
                    block.Children.Add(row);
                }
                else if (meta.IsFolder)
                {
                    var row = new DockPanel();
                    var browse = new Button { Content = "Browse...", Width = 80, Height = 26, Margin = new Thickness(6, 0, 0, 0) };
                    browse.Click += (_, _) =>
                    {
                        var dialog = new Microsoft.Win32.OpenFolderDialog { Title = meta.Label };
                        if (Directory.Exists(box.Text)) dialog.InitialDirectory = box.Text;
                        if (dialog.ShowDialog(this) == true) box.Text = dialog.FolderName;
                    };
                    DockPanel.SetDock(browse, Dock.Right);
                    row.Children.Add(browse);
                    row.Children.Add(box);
                    block.Children.Add(row);
                }
                else
                {
                    block.Children.Add(box);
                }
                read = () => box.Text.Trim();
            }
        }

        if (!string.IsNullOrWhiteSpace(meta.Description))
        {
            block.Children.Add(new TextBlock
            {
                Text = meta.Description,
                TextWrapping = TextWrapping.Wrap,
                Foreground = (Brush)FindResource("Brush.TextMuted"),
                FontSize = 11,
                Margin = new Thickness(prop.PropertyType == typeof(bool) ? 20 : 0, 3, 0, 0)
            });
        }

        _bindings.Add(new FieldBinding { Section = section, Property = prop, Meta = meta, SectionTitle = sectionTitle, Read = read });
        return block;
    }

    // ======================================================================= buttons

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        var errors = new List<string>();
        foreach (var b in _bindings)
        {
            try
            {
                b.Property.SetValue(b.Section, b.Read());
            }
            catch (FormatException ex)
            {
                errors.Add(ex.Message);
            }
        }
        errors.AddRange(SettingsStore.Validate(_working));
        if (errors.Count > 0)
        {
            MessageBox.Show(this, string.Join(Environment.NewLine, errors), "Settings not saved", MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        try
        {
            SettingsStore.Save(_working);
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
        {
            MessageBox.Show(this,
                $"Could not write {SettingsStore.SettingsPath}:\n\n{ex.Message}\n\n" +
                "The settings file lives next to the .exe, so that folder must be writable by your account. " +
                "Move the program to a writable folder or adjust the folder permissions.",
                "Settings not saved", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        SavedSettings = _working;
        DialogResult = true;
    }

    private void BtnReload_Click(object sender, RoutedEventArgs e)
    {
        var result = SettingsStore.Load();
        if (!string.IsNullOrEmpty(result.Error))
        {
            MessageBox.Show(this, result.Error, "Reload", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        _working = result.Settings;
        BuildTabs();
    }

    private void BtnDefaults_Click(object sender, RoutedEventArgs e)
    {
        var answer = MessageBox.Show(this,
            "Replace every value on every tab with the built-in defaults?\n\nNothing is written until you click Save.",
            "Restore defaults", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (answer != MessageBoxResult.Yes) return;
        _working = new AppSettings();
        BuildTabs();
    }

    private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
    {
        var path = SettingsStore.SettingsPath;
        if (!File.Exists(path))
        {
            MessageBox.Show(this, $"{path} does not exist yet. Click Save first.", "Open file", MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
            Process.Start(new ProcessStartInfo("notepad.exe", $"\"{path}\"") { UseShellExecute = true });
        }
        MessageBox.Show(this, "After editing the file by hand, save it and click 'Reload from file'.", "Open file",
            MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
