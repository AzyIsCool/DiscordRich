using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using static DiscordRpc;
using EnvDTE;
using EnvDTE80;

namespace DiscordRich
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class EnableDiscordRich
    {
        public static RichClass richClass = new RichClass();
        public static bool shutdown = false;
        static bool initialize = false;

        static RichPresence presence;
        static int callbackCalls;
        static EventHandlers handlers;

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("9cead51b-d641-4969-a680-992f7edef830");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="EnableDiscordRich"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private EnableDiscordRich(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static EnableDiscordRich Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new EnableDiscordRich(package);
        }

        public static void StartUp()
        {
            Console.WriteLine("Discord: init");
            callbackCalls = 0;

            handlers = new EventHandlers();
            handlers.readyCallback = ReadyCallback;
            handlers.disconnectedCallback += DisconnectedCallback;
            handlers.errorCallback += ErrorCallback;

            shutdown = false;
            initialize = true;

            DiscordRpc.Initialize(richClass.applicationId, ref handlers, true, String.Empty);
            UpdatePresenceVS();
            Update();
        }

        static async void UpdatePresenceVS()
        {
            if (!shutdown)
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    DTE2 dte2 = Package.GetGlobalService(typeof(DTE)) as DTE2;
                    Solution solution = dte2.Solution;

                    if (solution.IsOpen)
                    {
                        if (!richClass.secretmode)
                        {
                            string[] d = solution.FullName.Split("\\".ToCharArray());
                            presence.details = string.Format("I'm currently working on");
                            presence.state = $"{d[d.Length - 1].Replace(".sln", "")}";
                            presence.instance = false;
                        }
                        else
                        {
                            presence.details = string.Format($"I'm working on something you're not");
                            presence.state = "allowed to know about!";
                        }
                    }
                    else
                    {
                        presence.details = string.Format("I'm currently working on");
                        presence.state = "Nothing";
                    }
                    UpdatePresence(ref presence);
                    System.Timers.Timer timer = new System.Timers.Timer();
                    timer.AutoReset = false;
                    timer.Interval = 1000 * 5;
                    timer.Elapsed += Timer1_Elapsed;
                    timer.Start();
                });
            }
        }

        private static void Timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (!shutdown)
            {
                UpdatePresenceVS();
            }
        }

        public static void ReadyCallback()
        {
            ++callbackCalls;
            Console.WriteLine("Discord Client: ready");
        }

        public static void DisconnectedCallback(int errorCode, string message)
        {
            ++callbackCalls;
            Console.WriteLine(string.Format("Discord Client: disconnect {0}: {1}", errorCode, message));
        }

        public static void ErrorCallback(int errorCode, string message)
        {
            ++callbackCalls;
            Console.WriteLine(string.Format("Discord Client: error {0}: {1}", errorCode, message));
        }

        static async void Update()
        {
            if (!shutdown)
            {
            await System.Threading.Tasks.Task.Run(() =>
            {
                RunCallbacks();
                System.Timers.Timer timer = new System.Timers.Timer();
                timer.AutoReset = false;
                timer.Interval = 1000 * 6;
                timer.Elapsed += Timer_Elapsed;
                timer.Start();
            });
            }
        }

        private static void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (!shutdown)
            {
                Update();
            }
            else
            {
                Shutdown();
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = "DiscordRich is already enabled!";
            string title = "DiscordRich";

            if (shutdown | !initialize)
            {
                StartUp();
                message = "DiscordRich has been enabled";
            }

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_NOICON,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
