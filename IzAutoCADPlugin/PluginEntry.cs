using System;
using System.IO;
using System.Reflection;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

namespace IzAutoCADPlugin
{
    public class PluginEntry : IExtensionApplication
    {
        private const string TabId = "IZ_TOOLS_TAB";
        private const string PanelSourceId = "IZ_TOOLS_PANEL";

        public void Initialize()
        {
            Application.Idle += OnFirstIdle;
        }

        public void Terminate()
        {
            Application.Idle -= OnFirstIdle;
        }

        [CommandMethod("IZSHOW")]
        public void ShowWindowCommand()
        {
            ShowExportDialog();
        }

        private void OnFirstIdle(object sender, EventArgs e)
        {
            Application.Idle -= OnFirstIdle;
            TryCreateRibbonUi();
        }

        private static void TryCreateRibbonUi()
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon == null)
            {
                return;
            }

            RibbonTab tab = FindTab(ribbon, TabId);
            if (tab == null)
            {
                tab = new RibbonTab
                {
                    Title = "IZ Tools",
                    Id = TabId
                };
                ribbon.Tabs.Add(tab);
            }

            if (FindPanel(tab, PanelSourceId) != null)
            {
                return;
            }

            RibbonPanelSource panelSource = new RibbonPanelSource
            {
                Title = "Demo",
                Id = PanelSourceId
            };

            RibbonPanel panel = new RibbonPanel
            {
                Source = panelSource
            };

            RibbonButton button = new RibbonButton
            {
                Text = "DWG Export CSV",
                ShowText = true,
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Size = RibbonItemSize.Large,
                Image = LoadIcon("icon16.png"),
                LargeImage = LoadIcon("icon32.png"),
                CommandHandler = new RibbonButtonHandler(ShowExportDialog)
            };

            panelSource.Items.Add(button);
            tab.Panels.Add(panel);
        }

        private static RibbonTab FindTab(RibbonControl ribbon, string tabId)
        {
            foreach (RibbonTab tab in ribbon.Tabs)
            {
                if (tab.Id == tabId)
                {
                    return tab;
                }
            }

            return null;
        }

        private static RibbonPanel FindPanel(RibbonTab tab, string panelId)
        {
            foreach (RibbonPanel panel in tab.Panels)
            {
                if (panel.Source != null && panel.Source.Id == panelId)
                {
                    return panel;
                }
            }

            return null;
        }

        private static void ShowExportDialog()
        {
            using (var dialog = new ExportDialog())
            {
                Application.ShowModalDialog(dialog);
            }
        }

        private static BitmapImage LoadIcon(string fileName)
        {
            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrWhiteSpace(assemblyDir))
            {
                return null;
            }

            string iconPath = Path.Combine(assemblyDir, "Resources", fileName);
            if (!File.Exists(iconPath))
            {
                return null;
            }

            return new BitmapImage(new Uri(iconPath, UriKind.Absolute));
        }

        private sealed class RibbonButtonHandler : ICommand
        {
            private readonly Action _onExecute;

            public RibbonButtonHandler(Action onExecute)
            {
                _onExecute = onExecute;
            }

            public bool CanExecute(object parameter) => true;

            public event EventHandler CanExecuteChanged
            {
                add { }
                remove { }
            }

            public void Execute(object parameter)
            {
                _onExecute?.Invoke();
            }
        }
    }
}
