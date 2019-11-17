using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Linq;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFirewallHelper;
using Newtonsoft.Json.Linq;
using ffxiv_act_installer.Extensions;

namespace ffxiv_act_installer
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            WebClient webclient = new WebClient();
            int selection;
            var installpath = @"C:\Program Files (x86)\Advanced Combat Tracker";
            var overlaypath = "overlay.zip";
            var cactbotpath = "cactbot.zip";
            var ffxivpluginpath = "FFXIV_ACT_Plugin.dll";
            var configpath = $"{Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}/Advanced Combat Tracker/Config";
            var architecture = "x64";
            if (!System.Environment.Is64BitOperatingSystem)
            {
                architecture = "x32";
            }

            if (IsAdministrator())
            {
                Console.WriteLine("Welcome to the FFXIV ACT Installer. Please select an installer package:");
                Console.WriteLine("(1) ACT + parse overlay - recommended");
                Console.WriteLine("(2) ACT + Cactbot");
            }

            else
            {
                Console.WriteLine("Please launch this program as an administrator. Press any key to exit.");
                Console.ReadKey();
                Environment.Exit(0);
            }

            while (!int.TryParse(Console.ReadLine(), out selection) || selection > 2 || selection < 1)
            {
                Console.Clear();
                Console.WriteLine("Invalid option selected. Please select an installer package:");
                Console.WriteLine("(1) ACT + parse overlay - recommended");
                Console.WriteLine("(2) ACT + Cactbot");
            }
            Console.Clear();
            Console.WriteLine();
            Console.WriteLine("Downloading ACT...");
            try
            {
                webclient.DownloadFile("https://advancedcombattracker.com/includes/page-download.php?id=56", "act.exe");
                Console.Clear();
                Console.WriteLine("PLEASE UNINSTALL ACT IF IT'S INSTALLED BEFORE PROCEEDING.");
                Console.WriteLine($"It is strongly recommended to install in the default location ({installpath})");
                Console.WriteLine("Please do not exit this screen until prompted.");
                Console.WriteLine("Press any button to continue...");
                Console.ReadKey();
                Console.Clear();
                Console.WriteLine("Waiting for ACT installation to complete.");
                var process = Process.Start("act.exe");
                process.WaitForExit();
                Console.Clear();

                while (!Directory.Exists(installpath))
                {
                    Console.WriteLine("Can't find ACT in the usual directory. Please specify its install directory.");
                    FolderBrowserDialog browser2 = new FolderBrowserDialog();
                    if (browser2.ShowDialog() == DialogResult.OK)
                    {
                        foreach (var path in Directory.GetDirectories(browser2.SelectedPath))
                        {
                            Console.WriteLine(path); // full path
                            installpath = path;
                        }
                    }

                    Console.Write("Searching");
                    for (var x = 0; x < 3; x++)
                    {
                        Thread.Sleep(700);
                        Console.Write(".");
                    }
                }

                Console.WriteLine("Found ACT install!");
                Console.Write("Fetching ffxiv plugin...");
                webclient.DownloadFile("https://advancedcombattracker.com/includes/page-download.php?id=73", "FFXIV_ACT_Plugin.dll");
                Console.Write("Done!");
                Console.WriteLine();
                File.Copy(ffxivpluginpath, $"{installpath}/{ffxivpluginpath}", true);
                Console.Write("Fetching Hibiyasleep overlay...");
                FetchLatestGithubRelease(architecture, "https://api.github.com/repos/hibiyasleep/OverlayPlugin/releases/latest", "overlay.zip").Wait();
                Console.Write("Done!");
                Console.WriteLine();
                Console.Write("Extracting overlay files...");
                Directory.CreateDirectory($"{installpath}/OverlayPlugin/");
                Directory.CreateDirectory(configpath);
                var overlayarchive = new ZipArchive(new FileStream(overlaypath, FileMode.Open));
                overlayarchive.ExtractToDirectory($"{installpath}/OverlayPlugin/", true);
                Console.Write("Done!");
                Console.WriteLine();
                Console.WriteLine("Copying config files...");
                var configfiles = Directory.GetFiles($"{Directory.GetCurrentDirectory()}/act-config", "*config.xml");
                foreach (var config in configfiles)
                {
                    string rawtext = File.ReadAllText(config);
                    rawtext = rawtext.Replace("Luke", Environment.UserName); //replace author name with user's
                    File.WriteAllText(config, rawtext);
                    var dest = $"{configpath}/{config.Substring(config.LastIndexOf("\\"))}";
                    File.Copy(config, dest, true);
                }

                switch (selection)
                {       
                    case 2:
                        
                        Console.Write("Fetching Cactbot...");
                        FetchLatestGithubRelease(architecture, "https://api.github.com/repos/quisquous/cactbot/releases/latest", "cactbot.zip").Wait();
                        Console.Write("Done!");
                        Console.WriteLine();
                        Console.Write("Extracting cactbot...");
                        ZipFile.ExtractToDirectory(cactbotpath, $"{installpath}/OverlayPlugin/");
                        var cactbotversion = string.Empty;
                        var dirs = Directory.GetDirectories($"{installpath}/OverlayPlugin/");
                        foreach (var dir in dirs)
                        {
                            if (dir.Contains("cactbot-"))
                            {
                                cactbotversion = dir.Substring(dir.LastIndexOf("/") + 1);
                            }
                        }
                        cactbotpath = $"{installpath}/OverlayPlugin/{cactbotversion}/OverlayPlugin";
                        DirectoryCopy(cactbotpath, $"{installpath}/OverlayPlugin", true);
                        Directory.Delete($"{installpath}/OverlayPlugin/{cactbotversion}", true);
                        Console.Write("Done!");
                        Console.WriteLine();
                        break;
                    default:
                        break;
                }

                
                var installexe = @"\Advanced Combat Tracker.exe";
                var firewallpath = installpath + installexe;
                var allRules = FirewallManager.Instance.Rules.ToArray();
                if (allRules.Any(t => t.Name.Contains("ACT Inbound Rule") || t.Name.Contains("ACT Outbound Rule")))
                {
                    Console.WriteLine("Deleting old firewall exceptions...");
                    var oldoutboundrule = allRules.FirstOrDefault(t => t.Name.Contains("ACT Outbound Rule"));
                    var oldinboundrule = allRules.FirstOrDefault(t => t.Name.Contains("ACT Inbound Rule"));
                    if (oldoutboundrule != null)
                    {
                        FirewallManager.Instance.Rules.Remove(oldoutboundrule);
                    }

                    if (oldinboundrule != null)
                    {
                        FirewallManager.Instance.Rules.Remove(oldinboundrule);
                    }
                }

                Console.WriteLine("Adding firewall exceptions...");
                var outboundrule = FirewallManager.Instance.CreateApplicationRule(
                    FirewallManager.Instance.GetProfile().Type, @"ACT Outbound Rule",
                    FirewallAction.Allow, @firewallpath);
                outboundrule.Direction = FirewallDirection.Outbound;
                FirewallManager.Instance.Rules.Add(outboundrule);

                var inboundrule = FirewallManager.Instance.CreateApplicationRule(
                    FirewallManager.Instance.GetProfile().Type, @"ACT Inbound Rule",
                    FirewallAction.Allow, @firewallpath);
                inboundrule.Direction = FirewallDirection.Inbound;
                FirewallManager.Instance.Rules.Add(inboundrule);

                Console.WriteLine("Installation complete! ACT will now run. Press any button to exit.");
                Process.Start($"{installpath}/Advanced Combat Tracker.exe");
                Console.ReadKey();
                Environment.Exit(0);

            }
            catch (Exception e)
            {
                Console.WriteLine($"An error has occurred: {e.Message}");
                Console.WriteLine(e.InnerException);
                Console.WriteLine(e.StackTrace);
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
            }
        }
        public static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public static async Task FetchLatestGithubRelease(string architecture, string repo, string filename)
        {
            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Add("User-Agent", "ffxiv-act-installer");
                using (var request = new HttpRequestMessage(new HttpMethod("GET"), repo))
                {
                    try
                    {
                        var response = await httpClient.SendAsync(request);
                        var content = response.Content.ReadAsStringAsync();
                        var releaseurls = JObject.Parse(content.Result)
                            .SelectToken("assets")
                            .Select(t => t.SelectToken("browser_download_url")
                                .Value<string>())
                            .ToList();
                        var webclient = new WebClient();
                        var url = releaseurls.FirstOrDefault().ToString();
                        if (repo.Contains("hibiyasleep"))
                        {
                            if (architecture == "x32")
                            {
                                url = releaseurls.FirstOrDefault(t => t.Contains($"{architecture}-full") || t.Contains("x86-full"));
                            }
                            else
                            {
                                url = releaseurls.FirstOrDefault(t => t.Contains($"{architecture}-full"));
                            }
                        }
                        
                        if (url == null)
                        {
                            throw new ApplicationException("Could not retrieve github url.");
                        }

                        webclient.DownloadFile(url, filename);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Web request error: {e.Message}");
                    }
                }
            }
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName,
            bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
    }
}
