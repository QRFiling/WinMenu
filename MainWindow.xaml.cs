using IWshRuntimeLibrary;
using Microsoft.Win32;
using Shell32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using File = System.IO.File;
using Folder = Shell32.Folder;
using Path = System.IO.Path;

namespace WinMenu
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;

            appList.ItemsSource = apps;
            pinnedList.ItemsSource = pinned;
            lastOpenedList.ItemsSource = lastOpened;
            recentFilesList.ItemsSource = recent;
            fastLinksList.ItemsSource = fastLinks;
        }

        bool errorMessage;

        void Current_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            errorMessage = true;

            MessageBox.Show("Произошла внутреняя необработаная ошибка:\n\r" + e.Exception.Message + "\n\r\n\rНо ничего, программа продолжит" +
                " работать в штатном режиме (наверное)", "WinMenu", MessageBoxButton.OK, MessageBoxImage.Error);

            errorMessage = false;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        const byte keyControl = 0x11;
        const byte keyEscape = 0x1B;
        const uint KEYEVENTF_KEYUP = 0x02;

        Rect GetClassicStartMenuButtonBounds()
        {
            try
            {
                AutomationElement tray = AutomationElement.FromHandle(FindWindow("Shell_TrayWnd", null));
                AutomationElement startButton = tray.FindFirst(TreeScope.Subtree, new AndCondition
                (
                    new PropertyCondition(AutomationElement.ClassNameProperty, "Button"),
                    new PropertyCondition(AutomationElement.AccessKeyProperty, "Ctrl + Esc"))
                );

                return startButton.Current.BoundingRectangle;
            }
            catch { return Rect.Empty; }
        }

        async void OpenClassicStartMenu()
        {
            window_Deactivated(null, null);
            await Task.Delay(100);
            
            keybd_event(keyControl, 0, 0, UIntPtr.Zero);
            keybd_event(keyEscape, 0, 0, UIntPtr.Zero);

            keybd_event(keyControl, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event(keyEscape, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        Rect startButtonBounds;

        protected override void OnContentRendered(EventArgs e)
        {
            startButtonBounds = GetClassicStartMenuButtonBounds();
            WindowHook();

            UserActivityHook hook = new UserActivityHook();
            hook.Start(true, false);

            hook.OnMouseActivity += (s, a) =>
            {
                if (a.Button == System.Windows.Forms.MouseButtons.Middle && startButtonBounds.Contains(a.X, a.Y))
                    OpenMenu();
            };

            UpdateAppsList();
            UpdatePinnedList();
            JumpListHelper.UpdateAppIds();
            base.OnContentRendered(e);
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            Application.Current.Shutdown();
        }

        double newLeft = 0;

        void SetMenuSize()
        {
            Rect workArea = SystemParameters.WorkArea;
            Top = workArea.Top;
            Left = workArea.Left;

            Height = workArea.Height;
            Width = workArea.Width / 3.5;
        }

        async void OpenMenu()
        {
            await UpdateColor();
            SetMenuSize();

            Show();
            WindowState = WindowState.Normal;
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                BeginAnimation(LeftProperty, new DoubleAnimation
                {
                    Duration = TimeSpan.FromSeconds(1),
                    From = newLeft - 15,
                    To = newLeft,
                    EasingFunction = new PowerEase { Power = 10 }
                });

                if ((DateTime.Now - closeTime).TotalMinutes > 1)
                {
                    UpdateRecentFiles();
                    UpdateAutoRun();
                }

                if ((DateTime.Now - closeTime).TotalMinutes >= 10)
                {
                    JumpListHelper.UpdateAppIds();
                    UpdateRecycleBin();
                    UpdateTemp();
                }
            }
            
            base.OnStateChanged(e);
        }

        DateTime closeTime = DateTime.Now - TimeSpan.FromMinutes(10);
        bool firstRun = true;

        void window_Deactivated(object sender, EventArgs e)
        {
            if (firstRun == false && errorMessage == false)
            {
                DoubleAnimation animation = new DoubleAnimation
                {
                    Duration = TimeSpan.FromSeconds(0.1),
                    From = Left,
                    To = Left - 15,
                };

                animation.Completed += (s, a) =>
                {
                    WindowState = WindowState.Minimized;
                    Hide();

                    closeTime = DateTime.Now;
                    GetScrollViewer(appList).ScrollToTop();
                    GetScrollViewer(pinnedList).ScrollToTop();
                };

                if (searchPanel.Visibility == Visibility.Visible)
                    Border_MouseLeftButtonUp_2(null, null);

                BeginAnimation(LeftProperty, animation);
            }

            firstRun = false;
        }

        ScrollViewer GetScrollViewer(DependencyObject o)
        {
            if (o is ScrollViewer)
                return o as ScrollViewer;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(o); i++)
            {
                var child = VisualTreeHelper.GetChild(o, i);
                var result = GetScrollViewer(child);

                if (result == null) continue;
                else return result;
            }

            return null;
        }

        public class App
        {
            FileSystemInfo file;

            public FileSystemInfo File
            {
                get { return file; }
                set
                {
                    file = value;
                    NewApp = (DateTime.Now - value.CreationTime).TotalDays <= 3;

                    if (string.IsNullOrEmpty(DisplayName))
                    {
                        DisplayName = Directory.Exists(value.FullName) ? value.Name :
                            (KeepExtension ? value.Name : Path.GetFileNameWithoutExtension(value.FullName));
                    }

                    if (StartTime == DateTime.MinValue)
                        DisplayDate = value.CreationTime.ToString("dd.MM.yy");
                }
            }

            public string DisplayName { get; set; }
            public string DisplayDate { get; set; }
            public Visibility DisplayVisibility { get; set; }
            public string LetterSeparator { get; set; }
            public BitmapImage Icon { get; set; }
            public bool NewApp { get; set; }
            public DateTime StartTime { get; set; }
            public bool KeepExtension { get; set; }
        }
        class PinnedApp
        {
            FileSystemInfo file;

            public FileSystemInfo File
            {
                get { return file; }
                set
                {
                    file = value;
                    DisplayName = Path.GetFileNameWithoutExtension(value.FullName);
                }
            }

            public string DisplayName { get; set; }
            public BitmapImage Icon { get; set; }
        }
        class FastLink
        {
            public string DisplayName { get; set; }
            public string Path { get; set; }
            public string Args { get; set; }
            public BitmapImage Icon { get; set; }
            public bool HaveAlternativePaths { get; set; }
        }

        static BitmapImage GetIcon(string name) =>
             new BitmapImage(new Uri($"pack://application:,,,/WinMenu;component/Resources/{name}.png"));

        ObservableCollection<App> apps = new ObservableCollection<App>();
        ObservableCollection<PinnedApp> pinned = new ObservableCollection<PinnedApp>();
        ObservableCollection<App> lastOpened = new ObservableCollection<App>();
        ObservableCollection<App> recent = new ObservableCollection<App>();

        ObservableCollection<FastLink> fastLinks = new ObservableCollection<FastLink>
        {
            new FastLink { DisplayName = "Мой компьютер", Icon = GetIcon("PC"), Path = "explorer", Args = "::{20d04fe0-3aea-1069-a2d8-08002b30309d}" },
            new FastLink { DisplayName = "Панель управления", Icon = GetIcon("Control"), Path = "control" },
            new FastLink { DisplayName = "Командная строка", Icon = GetIcon("Terminal"), Path = "cmd" },
            new FastLink { DisplayName = "Реестр", Icon = GetIcon("Struct"), Path = "regedit" },
            new FastLink { DisplayName = "Звук", Icon = GetIcon("Sound"), Path = "control", Args = "mmsys.cpl sounds" },
            new FastLink { DisplayName = "Устройства", Icon = GetIcon("USB"), Path = "devmgmt.msc" },
            new FastLink { DisplayName = "Игры", Icon = GetIcon("Controller"), HaveAlternativePaths = true,
                Path = $@"D:\Games|{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\Saved Games" }
        };

        void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _ = UpdateColor();
            SetMenuSize();
            WindowState = WindowState.Minimized;
            Hide();

            offList.ItemsSource = new App[]
            {
                new App { DisplayName = "Перезагрузка", Icon = GetIcon("Reboot") },
                new App { DisplayName = "Смена пользователя", Icon = GetIcon("UserSwitch") },
                new App { DisplayName = "Спящий режим", Icon = GetIcon("Sleep") },
                new App { DisplayName = "Гибернация", Icon = GetIcon("Time") }
            };

            autorunList.ItemsSource = new App[]
            {
                new App { DisplayName = "Автозагрузка в реестре", Icon = GetIcon("Fast") },
                new App { DisplayName = "Автозагрузка в пуске", Icon = GetIcon("Folder") },
                new App { DisplayName = "Планировщик задач", Icon = GetIcon("Task") },
                new App { DisplayName = "Службы", Icon = GetIcon("Service") }
            };

            appMenuList.ItemsSource = new App[]
            {
                new App { DisplayName = "Найти в проводнике", Icon = GetIcon("Folder") },
                new App { DisplayName = "От имени администратора", Icon = GetIcon("Shield") },
                new App { DisplayName = "Ярлык на рабочем столе", Icon = GetIcon("Share") },
                new App { DisplayName = "Удалить из этого списка", Icon = GetIcon("Delete") },
                new App { DisplayName = "Свойства", Icon = GetIcon("InfoList") }
            };

            moreMenuList.ItemsSource = new App[]
            {
                new App { DisplayName = "Добавить в автозапуск", Icon = GetIcon("Autostart") },
                new App { DisplayName = "Открыть классическое меню", Icon = GetIcon("Windows") },
                new App { DisplayName = "Закрыть программу", Icon = GetIcon("Close") }
            };

            recycleBinList.ItemsSource = new App[] { new App { DisplayName = "Очистить корзину", Icon = GetIcon("Delete") } };
            tempList.ItemsSource = new App[] { new App { DisplayName = "Очистить папку", Icon = GetIcon("Delete") } };

            ICollectionView view = CollectionViewSource.GetDefaultView(apps);
            view.Filter = (o) =>
            {
                App app = o as App;
                return app.File.Name.ToLower().Contains(searchBox.Text.ToLower().Trim());
            };

            searchBox.TextChanged += (s, a) => view.Refresh();

            userInfo.Text = $"Учётная запись: {Environment.UserName}";
            UpdateWorkTime();

            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(5);
            timer.Start();

            timer.Tick += (s, a) =>
            {
                if (Visibility == Visibility.Visible)
                    UpdateWorkTime();
            };

            void UpdateWorkTime()
            {
                DateTime osStartTime = DateTime.Now - new TimeSpan(10000 * GetTickCount64());
                pcWorkInfo.Text = $"ПК работает {DateSting(osStartTime)}";
            }

            InitializeWathcers();
        }

        readonly DirectoryInfo startMenuApps1 = new DirectoryInfo(Environment.GetFolderPath(
            Environment.SpecialFolder.CommonStartMenu) + "\\Programs");

        readonly DirectoryInfo startMenuApps2 = new DirectoryInfo(Environment.GetFolderPath(
            Environment.SpecialFolder.StartMenu) + "\\Programs");

        readonly DirectoryInfo startMenuPinnedAps = new DirectoryInfo(Environment.GetFolderPath(
            Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu");

        readonly DirectoryInfo recentFiles = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Recent));
        readonly BitmapImage folderIcon = GetIcon("FolderIcon");

        FileSystemWatcher watcher1;
        FileSystemWatcher watcher2;
        FileSystemWatcher watcher3;

        void InitializeWathcers()
        {
            watcher1 = new FileSystemWatcher(startMenuApps1.FullName);
            watcher2 = new FileSystemWatcher(startMenuApps2.FullName);
            watcher3 = new FileSystemWatcher(startMenuPinnedAps.FullName);

            foreach (var item in new FileSystemWatcher[] { watcher1, watcher2, watcher3 })
            {
                item.Created += FolderContentChanged;
                item.Changed += FolderContentChanged;
                item.Deleted += FolderContentChanged;
                item.Renamed += FolderContentChanged;

                item.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                item.EnableRaisingEvents = true;
            }
        }

        void FolderContentChanged(object sender, FileSystemEventArgs e)
        {
            if (sender == watcher3)
            {
                if (pinnedListUpdated)
                    UpdatePinnedList();
            }
            else
            {
                if (appListUpdated)
                    UpdateAppsList();
            }
        }

        async Task UpdateColor()
        {
            var colors = await ScreenColors.GetColors();

            background.GradientStops[1].Color = colors.Item1;
            background.GradientStops[0].Color = colors.Item2;

            FindVisualChilds<Popup>(this).Select(s => (s.Child as Grid).Children.
                OfType<Border>().First()).ToList().ForEach(f => f.Background = colors.Item3);
        }

        IEnumerable<T> FindVisualChilds<T>(DependencyObject depObj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                DependencyObject ithChild = VisualTreeHelper.GetChild(depObj, i);
                if (ithChild == null) continue;
                if (ithChild is T t) yield return t;
                foreach (T childOfChild in FindVisualChilds<T>(ithChild)) yield return childOfChild;
            }
        }

        bool appListUpdated = true;

        async void UpdateAppsList()
        {
            appListUpdated = false;
            
            await Task.Run(() =>
            {
                Dispatcher.Invoke(() => apps.Clear());
                char prevfirstLetter = char.MinValue;

                List<FileSystemInfo> files = GetStartMenuFiles(startMenuApps1)
                    .Concat(GetStartMenuFiles(startMenuApps2)).OrderBy(o => o.Name).ToList();

                foreach (var item in files)
                {
                    bool isFile = File.Exists(item.FullName);
                    System.Drawing.Bitmap bitmap = isFile ? System.Drawing.Icon.ExtractAssociatedIcon(item.FullName).ToBitmap() : null;

                    char currFirstLetter = char.ToUpper(item.Name.First());
                    if (char.IsDigit(currFirstLetter) || char.IsSymbol(currFirstLetter)) currFirstLetter = '#';

                    Dispatcher.Invoke(() =>
                    {
                        apps.Add(new App
                        {
                            File = item,
                            Icon = isFile ? BitmapToBitmapImage(bitmap) : folderIcon,
                            LetterSeparator = prevfirstLetter != currFirstLetter ? currFirstLetter.ToString() : null
                        });

                        appsTip.Visibility = Visibility.Collapsed;
                        appsTipIcon.Visibility = Visibility.Collapsed;
                    });

                    prevfirstLetter = currFirstLetter;
                }
            });

            appListUpdated = true;

            IEnumerable<FileSystemInfo> GetStartMenuFiles(DirectoryInfo directory)
            {
                return directory.GetFiles().Select(s => (FileSystemInfo)s)
                    .Concat(directory.GetDirectories());
            }
        }

        bool pinnedListUpdated = true;

        async void UpdatePinnedList()
        {
            pinnedListUpdated = false;

            await Task.Run(() =>
            {
                Dispatcher.Invoke(() => pinned.Clear());

                List<FileSystemInfo> files = startMenuPinnedAps.GetFiles().Select(s => (FileSystemInfo)s).
                    Concat(startMenuPinnedAps.GetDirectories()).OrderByDescending(o => o.CreationTime).ToList();

                foreach (var item in files)
                {
                    bool isFile = File.Exists(item.FullName);
                    System.Drawing.Bitmap bitmap = isFile ? System.Drawing.Icon.ExtractAssociatedIcon(item.FullName).ToBitmap() : null;

                    Dispatcher.Invoke(() =>
                    {
                        pinned.Add(new PinnedApp
                        {
                            File = item,
                            Icon = isFile ? BitmapToBitmapImage(bitmap) : folderIcon,
                        });

                        pinTipIcon.Visibility = Visibility.Collapsed;
                        pinTip.Visibility = Visibility.Collapsed;
                    });
                }
            });

            pinnedListUpdated = true;
        }

        async void UpdateRecentFiles()
        {
            await Task.Run(() =>
            {
                Dispatcher.Invoke(() => recent.Clear());

                List<FileSystemInfo> files = recentFiles.GetFiles().OrderByDescending(od => od.LastWriteTime)
                    .Take(15).Select(s => (FileSystemInfo)s).ToList();

                foreach (var item in files)
                {
                    System.Drawing.Bitmap bitmap = System.Drawing.Icon.ExtractAssociatedIcon(item.FullName).ToBitmap();

                    Dispatcher.Invoke(() =>
                    {
                        recent.Add(new App
                        {
                            File = item,
                            Icon = BitmapToBitmapImage(bitmap),
                        });
                    });
                }
            });

            recentFilesTip.Visibility = Visibility.Collapsed;
        }

        async void UpdateRecycleBin()
        {
            FolderItems items = null;
            string size = string.Empty;

            await Task.Run(() =>
            {
                try
                {
                    items = new Shell().NameSpace(10).Items();
                    size = BytesToString(items.OfType<FolderItem>().Sum(s => (long)s.Size));
                }
                catch { }
            });

            if (items != null)
            {
                recycleBinSize.Text = $"Размер: {size}";
                recycleBinCount.Text = $"Всего файлов: {items.Count}";
            }
        }

        [DllImport("Kernel32.dll")]
        static extern long GetTickCount64();

        async void UpdateAutoRun()
        {
            string[] pathsCurrent = new string[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
            };

            string[] pathsAll = new string[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
                @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run"
            };

            int countPrograms = 0;
            int countServices = 0;

            await Task.Run(() =>
            {
                foreach (var item in pathsCurrent)
                    countPrograms += GetValuesCount(item, true);

                foreach (var item in pathsAll)
                    countPrograms += GetValuesCount(item, false);

                ServiceController[] servises = ServiceController.GetServices();
                countServices = servises.Count(c => c.StartType == ServiceStartMode.Automatic);
            });

            autoRunProgramCount.Text = $"Всего программ: {countPrograms}";
            autoRunServiceCount.Text = $"Всего служб: {countServices}";

            int GetValuesCount(string path, bool currentUser)
            {
                using (RegistryKey key = currentUser ? Registry.CurrentUser.OpenSubKey(path) :
                    Registry.LocalMachine.OpenSubKey(path))
                    return key.GetValueNames().Length;
            }
        }

        async void UpdateTemp()
        {
            await Task.Run(() =>
            {
                DirectoryInfo di = new DirectoryInfo(Path.GetTempPath());
                long size = di.EnumerateFiles("*.*", SearchOption.AllDirectories).Sum(fi => fi.Length);

                DriveInfo drive = DriveInfo.GetDrives().FirstOrDefault(f =>
                    f.Name == Path.GetPathRoot(Environment.SystemDirectory));

                if (drive != null && drive.IsReady)
                {
                    double percent = Math.Round((double)(100 * size) / drive.TotalSize, 1);
                    Dispatcher.Invoke(() => tempPercentage.Text = $"{percent}% от диска {drive.Name}");
                }

                Dispatcher.Invoke(() => tempSize.Text = $"Размер: {BytesToString(size)}");
            });
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        static string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

        void WindowHook()
        {
            HookForm hook = new HookForm();
            hook.WindowEvent += (i) =>
            {
                try
                {
                    GetWindowThreadProcessId(i, out uint id);
                    string fileName = Process.GetProcessById((int)id).MainModule.FileName;

                    if (Path.GetFileNameWithoutExtension(fileName).ToLower() != "explorer" && fileName != exePath)
                    {
                        FileInfo file = new FileInfo(fileName);
                        
                        lastOpened.Insert(0, new App
                        {
                            StartTime = DateTime.Now,
                            DisplayName = FileVersionInfo.GetVersionInfo(fileName).ProductName,
                            File = file,
                            Icon = BitmapToBitmapImage(System.Drawing.Icon.ExtractAssociatedIcon(fileName).ToBitmap())
                        });

                        for (int j = 1; j < lastOpened.Count; j++)
                        {
                            if (lastOpened[j].File.FullName == fileName)
                            {
                                lastOpened.RemoveAt(j);
                                break;
                            }
                        }

                        lastProgramsTip.Visibility = Visibility.Collapsed;

                        if (lastOpened.Count > 20)
                            lastOpened.RemoveAt(20);
                    }
                }
                catch { }
            };
        }

        string DateSting(DateTime date)
        {
            const int SECOND = 1;
            const int MINUTE = 60 * SECOND;
            const int HOUR = 60 * MINUTE;
            const int DAY = 24 * HOUR;

            var ts = new TimeSpan(DateTime.Now.Ticks - date.Ticks);
            double delta = Math.Abs(ts.TotalSeconds);

            if (delta < 1 * MINUTE)
                return ts.Seconds == 1 ? "сек" : ts.Seconds + " сек";

            if (delta < 2 * MINUTE)
                return "1 мин";

            if (delta < 45 * MINUTE)
                return ts.Minutes + " мин";

            if (delta < 90 * MINUTE)
                return "1 ч";

            if (delta < 24 * HOUR)
                return ts.Hours + " ч";

            if (delta < 48 * HOUR)
                return "вчера";

            if (delta < 30 * DAY)
                return ts.Days + " дн";

            return date.ToString("dd.MM.yy");
        }
        string BytesToString(long byteCount)
        {
            string[] suf = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
            if (byteCount == 0) return "0" + suf[0];
            long bytes = Math.Abs(byteCount);
            int place = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
            double num = Math.Round(bytes / Math.Pow(1024, place), 1);
            return (Math.Sign(byteCount) * num).ToString("#.#") + suf[place];
        }
        public static BitmapImage BitmapToBitmapImage(System.Drawing.Bitmap bitmap)
        {
            if (bitmap == null)
                return null;

            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        void itemGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            App app = (sender as FrameworkElement).DataContext as App;
            OpenApp(app.File.FullName);
        }

        void parentGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            PinnedApp app = (sender as FrameworkElement).DataContext as PinnedApp;
            OpenApp(app.File.FullName);
        }

        void OpenApp(string path, string args = null, bool runas = false)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = path;
                process.StartInfo.Arguments = args;

                if (runas)
                {
                    process.StartInfo.UseShellExecute = true;
                    process.StartInfo.Verb = "runas";
                }

                try { process.Start(); }
                catch (Exception ex)
                {
                    if (ex is Win32Exception && File.Exists(path))
                    {
                        string target = GetShortcutTarget(path);
                        if (string.IsNullOrEmpty(target)) throw ex;

                        if (File.Exists(target) == false)
                            target = target.Replace("Program Files (x86)", "Program Files");

                        OpenApp(target, args, runas);
                        return;
                    }
                }

                window_Deactivated(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть файл или папку: " + ex.Message, "WinMenu",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void Button_Click_1(object sender, RoutedEventArgs e)
        {
            foreach (var item in lastOpened)
                item.DisplayDate = DateSting(item.StartTime);

            lastOpenedList.Items.Refresh();
            lastOpenedMenu.IsOpen = true;

            Button button = sender as Button;
            Point screen = button.PointToScreen(new Point(button.ActualWidth, 0));

            lastOpenedMenu.HorizontalOffset = screen.X + 5;
            lastOpenedMenu.VerticalOffset = screen.Y;
        }

        void Button_Click_2(object sender, RoutedEventArgs e)
        {
            foreach (var item in recent)
                item.DisplayDate = DateSting(item.File.LastWriteTime);

            recentFilesList.Items.Refresh();
            recentFilesMenu.IsOpen = true;

            Button button = sender as Button;
            Point screen = button.PointToScreen(new Point(button.ActualWidth, 0));

            recentFilesMenu.HorizontalOffset = screen.X + 5;
            recentFilesMenu.VerticalOffset = screen.Y;
        }

        void lastOpenedList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListBox listBox = sender as ListBox;
            if (listBox.SelectedItem == null) return;

            OpenApp((listBox.SelectedItem as App).File.FullName);
            listBox.SelectedItem = null;

            if (listBox == lastOpenedList) lastOpenedMenu.IsOpen = false;
            else if (listBox == recentFilesList) recentFilesMenu.IsOpen = false;
        }

        void Button_Click_3(object sender, RoutedEventArgs e) => OpenApp("explorer", "shell:RecycleBinFolder");
        void Button_Click_5(object sender, RoutedEventArgs e) => OpenApp(Path.GetTempPath());

        void Button_Click_4(object sender, RoutedEventArgs e)
        {
            string file = "msconfig.exe";
            string path = Environment.SystemDirectory + "\\" + file;

            if (!File.Exists(path))
                path = Environment.ExpandEnvironmentVariables(@"%windir%\Sysnative\" + file);

            OpenApp(path, "-4");
        }

        void Border_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            recentFilesMenu.IsOpen = false;
            OpenApp(recentFiles.FullName);
        }

        void Border_MouseLeftButtonUp_1(object sender, MouseButtonEventArgs e)
        {
            FastLink link = (sender as FrameworkElement).DataContext as FastLink;
            string[] paths = link.Path.Split('|');

            if (link.HaveAlternativePaths)
            {
                foreach (var item in paths)
                {
                    if (File.Exists(item) || Directory.Exists(item))
                    {
                        OpenApp(item, link.Args);
                        return;
                    }
                }
            }
            else OpenApp(link.Path, link.Args);
        }

        void Border_PreviewMouseLeftButtonUp_1(object sender, MouseButtonEventArgs e)
        {
            if (offButtonExpend.IsMouseOver)
            {
                Point screen = offButtonExpend.PointToScreen(new Point(0, 0));
                offMenu.IsOpen = true;

                offMenu.HorizontalOffset = screen.X;
                offMenu.VerticalOffset = screen.Y - (offMenu.Child as FrameworkElement).ActualHeight - 5;

                return;
            }

            StartShutDown("/s /t 0");
        }

        void offList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = offList.SelectedIndex;
            offList.SelectedItem = null;
            offMenu.IsOpen = false;

            if (index == 0) StartShutDown("-f -r -t 0");
            else if (index == 1) StartShutDown("-l");
            else if (index == 2) StartShutDown("sl");
            else if (index == 3) StartShutDown("hi");
        }

        void StartShutDown(string param)
        {
            window_Deactivated(null, null);

            try
            {
                if (param == "hi")
                    System.Windows.Forms.Application.SetSuspendState(System.Windows.Forms.PowerState.Suspend, true, true);

                if (param == "sl")
                    System.Windows.Forms.Application.SetSuspendState(System.Windows.Forms.PowerState.Hibernate, true, true);

                ProcessStartInfo proc = new ProcessStartInfo();
                proc.FileName = "cmd";
                proc.WindowStyle = ProcessWindowStyle.Hidden;
                proc.Arguments = "/C shutdown " + param;
                Process.Start(proc);
            }
            catch
            {
                MessageBox.Show("Не удалось завершить работу компьютера", "WinMenu",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        void userInfo_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e) => OpenApp("control", "/name Microsoft.UserAccounts");
        void Button_MouseRightButtonUp(object sender, MouseButtonEventArgs e) => recycleBinMenu.IsOpen = true;
        void Button_MouseRightButtonUp_1(object sender, MouseButtonEventArgs e) => autorunMenu.IsOpen = true;
        void Button_MouseRightButtonUp_2(object sender, MouseButtonEventArgs e) => tempMenu.IsOpen = true;


        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        static extern uint SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, int dwFlags);

        async void recycleBitList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (recycleBinList.SelectedItem != null)
            {
                recycleBinList.SelectedItem = null;
                recycleBinMenu.IsOpen = false;
                SHEmptyRecycleBin(IntPtr.Zero, null, 0);

                await Task.Delay(1000);
                UpdateRecycleBin();
            }
        }

        void autorunList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = autorunList.SelectedIndex;
            autorunList.SelectedItem = null;
            window_Deactivated(null, null);
            autorunMenu.IsOpen = false;

            if (index == 0) FindInRegistry(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run");
            else if (index == 1) OpenApp(Environment.GetFolderPath(Environment.SpecialFolder.Startup));
            else if (index == 2) OpenApp("taskschd.msc");
            else if (index == 3) OpenApp("services.msc");
        }

        void FindInRegistry(string pathToOpen)
        {
            string path = @"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit";
            Registry.CurrentUser.CreateSubKey(path).SetValue("LastKey", "Компьютер\\" + pathToOpen);

            Process.GetProcessesByName("regedit").ToList().ForEach(p => p.Kill());
            Process.Start("regedit");
        }

        async void tempList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (tempList.SelectedItem != null)
            {
                tempList.SelectedItem = null;
                tempMenu.IsOpen = false;

                MessageBoxResult result = MessageBox.Show("Очистить все возможные файлы из папки Temp?",
                    Title, MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    await ClearFolder(new DirectoryInfo(Path.GetTempPath()));
                    UpdateTemp();
                }
            }
        }

        async Task ClearFolder(DirectoryInfo directory)
        {
            await Task.Run(() =>
            {
                foreach (FileInfo file in directory.EnumerateFiles())
                {
                    try { file.Delete(); }
                    catch { }
                }

                foreach (DirectoryInfo dir in directory.EnumerateDirectories())
                {
                    try { dir.Delete(true); }
                    catch { }
                }
            });
        }

        void Border_MouseLeftButtonUp_2(object sender, MouseButtonEventArgs e)
        {
            DoubleAnimation animation = new DoubleAnimation
            {
                Duration = TimeSpan.FromSeconds(0.3),
                From = searchPanel.IsVisible ? 40 : 10,
                To = searchPanel.IsVisible ? 0 : 40,
                EasingFunction = new PowerEase { Power = 5 }
            };

            if (searchPanel.Visibility == Visibility.Visible)
            {
                animation.Completed += (s, a) =>
                {
                    searchPanel.Visibility = Visibility.Collapsed;
                    searchBox.Text = string.Empty;
                };
            }

            if (searchPanel.Visibility == Visibility.Collapsed)
            {
                searchPanel.Visibility = Visibility.Visible;
                searchBox.Focus();
            }

            searchPanel.BeginAnimation(HeightProperty, animation);
            appList.BeginAnimation(MarginProperty, new ThicknessAnimation
            {
                Duration = TimeSpan.FromSeconds(0.7),
                To = animation.To == 40 ? new Thickness(0, 60, 0, 5) : new Thickness(0, 10, 0, 5),
                EasingFunction = new PowerEase { Power = 10 }
            });
        }

        void Border_MouseLeftButtonUp_3(object sender, MouseButtonEventArgs e) => searchBox.Text = string.Empty;

        void appMenuList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (appMenuList.SelectedItem != null)
            {
                int index = appMenuList.SelectedIndex;
                appMenuList.SelectedItem = null;
                window_Deactivated(null, null);
                appMenu.IsOpen = false;

                if (index == 0) OpenApp("explorer", $"/select, \"{appMenuPath.Text}\"");
                else if (index == 1) OpenApp(appMenuPath.Text, null, true);
                else if (index == 2) CreateShortcut(appMenuPath.Text);
                else if (index == 4) FileProperties.Show(appMenuPath.Text);
                else if (index == 3)
                {
                    if (File.Exists(shortcutPath))
                    {
                        Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(shortcutPath,
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing);
                    }
                    else if (Directory.Exists(shortcutPath))
                    {
                        Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(shortcutPath,
                            Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin,
                            Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing);
                    }
                }
            }
        }

        void appMenuRecentList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (appMenuRecentList.SelectedItem != null)
            {
                window_Deactivated(null, null);
                appMenu.IsOpen = false;
                OpenApp(appMenuPath.Text, $"\"{(appMenuRecentList.SelectedItem as App).File.FullName}\"");
                appMenuRecentList.SelectedItem = null;
            }
        }

        string shortcutName = string.Empty;
        string shortcutPath = string.Empty;

        async void Border_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            appMenuName.Text = "...";
            appMenuPath.Text = "...";
            appMenuIcon.Source = null;
            appMenuRecentStatus.Text = "Загрузка элементов...";
            appMenu.IsOpen = true;

            FrameworkElement element = sender as FrameworkElement;
            App app = element.DataContext as App;
            PinnedApp pinnedApp = app == null ? element.DataContext as PinnedApp : null;
            FileSystemInfo file = app == null ? pinnedApp.File : app.File;

            shortcutName = app == null ? pinnedApp.DisplayName : app.DisplayName;
            string path = shortcutPath = file.FullName;
            string name = file.Name;

            await Task.Run(() =>
            {
                path = GetShortcutTarget(file.FullName);
                if (string.IsNullOrEmpty(path)) path = file.FullName;
                name = Directory.Exists(file.FullName) ? file.Name : Path.GetFileNameWithoutExtension(path);
            });

            appMenuName.Text = name;
            appMenuPath.Text = path;
            appMenuIcon.Source = app == null ? pinnedApp.Icon : app.Icon;

            List<App> jumplist = Directory.Exists(file.FullName) ? null :
                await JumpListHelper.GetJumpListFiles(Path.GetFileName(path));

            if (jumplist == null)
            {
                appMenuRecentStatus.Text = "Связанных элементов нет";
                appMenuRecentList.ItemsSource = null;
                return;
            }

            foreach (var item in jumplist)
            {
                item.DisplayDate = DateSting(item.StartTime);

                if (item.File.Exists)
                {
                    item.Icon = BitmapToBitmapImage(System.Drawing.Icon.
                        ExtractAssociatedIcon(item.File.FullName).ToBitmap());
                }
                else item.Icon = (BitmapImage)appMenuIcon.Source;
            }

            appMenuRecentList.Items.Refresh();
            appMenuRecentList.ItemsSource = jumplist;
            appMenuRecentStatus.Text = string.Empty;
        }

        string GetShortcutTarget(string shortcutFilename)
        {
            string pathOnly = Path.GetDirectoryName(shortcutFilename);
            string filenameOnly = Path.GetFileName(shortcutFilename);

            Shell shell = new Shell();
            Folder folder = shell.NameSpace(pathOnly);
            FolderItem folderItem = folder.ParseName(filenameOnly);

            if (folderItem != null)
            {
                if (folderItem.IsLink)
                {
                    ShellLinkObject link = (ShellLinkObject)folderItem.GetLink;
                    return link.Path;
                }

                return shortcutFilename;
            }

            return string.Empty;
        }

        void CreateShortcut(string path)
        {
            object shDesktop = "Desktop";
            WshShell shell = new WshShell();
            string shortcutAddress = (string)shell.SpecialFolders.Item(ref shDesktop) + $@"\{shortcutName}.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            shortcut.TargetPath = path;
            shortcut.Save();
        }

        bool lastAutorunValue;

        void Border_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            bool autorun = InAutorun();

            if (autorun != lastAutorunValue)
            {
                App app = (moreMenuList.ItemsSource as App[])[0];
                lastAutorunValue = autorun;

                app.DisplayName = lastAutorunValue ? "Убрать из автозапуска" : "Добавить в автозапуск";
                app.Icon = GetIcon(lastAutorunValue ? "AutostartOff" : "Autostart");

                moreMenuList.Items.Refresh();
            }

            moreMenu.IsOpen = true;
            Border button = sender as Border;
            Point screen = button.PointToScreen(new Point(0, 0));

            moreMenu.HorizontalOffset = screen.X;
            moreMenu.VerticalOffset = screen.Y - moreMenuList.ActualHeight - 15;
        }

        async void moreMenuList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (moreMenuList.SelectedItem != null)
            {
                int index = moreMenuList.SelectedIndex;
                moreMenuList.SelectedItem = null;

                if (index == 0)
                {
                    SetAutorun(!lastAutorunValue);
                    UpdateAutoRun();
                }
                else if (index == 1) OpenClassicStartMenu();
                else if (index == 2) Close();

                while (System.Windows.Forms.Control.MouseButtons != System.Windows.Forms.MouseButtons.None)
                    await Task.Delay(1);

                moreMenu.IsOpen = false;
            }
        }

        void SetAutorun(bool add)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (add) key.SetValue("WinMenu", "\"" + exePath + "\"");
                    else key.DeleteValue("WinMenu", false);
                }
            }
            catch { }
        }

        bool InAutorun()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"))
                    return key.GetValue("WinMenu") != null;
            }
            catch { return false; }
        }
    }
}
