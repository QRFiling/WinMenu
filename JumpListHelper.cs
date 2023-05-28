using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WinMenu
{
    class JumpListHelper
    {
        static DirectoryInfo jumplistDirectory = new DirectoryInfo(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Recent), "AutomaticDestinations"));

        static List<string> automaticDestinations = new List<string>();
        static Dictionary<string, string> actualAppIds = new Dictionary<string, string>();

        public static async void UpdateAppIds()
        {
            await Task.Run(() =>
            {
                if (jumplistDirectory.Exists == false)
                    return;

                automaticDestinations.Clear();
                actualAppIds.Clear();

                foreach (var item in jumplistDirectory.GetFiles())
                {
                    automaticDestinations.Add(item.FullName);
                    string appId = item.Name.Replace(item.Extension, string.Empty);

                    if (AppIDsData.Data.TryGetValue(appId, out string value))
                        actualAppIds.Add(appId, value);
                }
            });
        }

        async public static Task<List<MainWindow.App>> GetJumpListFiles(string exeName)
        {
            List<MainWindow.App> output = new List<MainWindow.App>();
            var pair = actualAppIds.FirstOrDefault(f => f.Value.ToLower() == exeName.ToLower());

            if (string.IsNullOrEmpty(pair.Key))
                return null;

            await Task.Run(() =>
            {
                var list = JumpList.JumpList.LoadAutoJumplist(automaticDestinations.
                    First(f => Path.GetFileNameWithoutExtension(f) == pair.Key)).DestListEntries;

                foreach (var item in list.Take(10))
                {
                    MainWindow.App app = new MainWindow.App
                    {
                        KeepExtension = true,
                        File = new FileInfo(item.Path),
                        StartTime = item.LastModified.LocalDateTime
                    };
   
                    output.Add(app);
                }
            });

            return output;
        }
    }
}
