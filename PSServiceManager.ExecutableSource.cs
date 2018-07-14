using System;
using System.ComponentModel;
using System.IO;
using System.IO.Pipes;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Threading;

namespace PSServiceManager
{
    public enum ServiceType : int
    {
        SERVICE_WIN32_OWN_PROCESS = 0x00000010,
        SERVICE_WIN32_SHARE_PROCESS = 0x00000020,
    };

    public enum ServiceState : int
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    };

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public ServiceType dwServiceType;
        public ServiceState dwCurrentState;
        public int dwControlsAccepted;
        public int dwWin32ExitCode;
        public int dwServiceSpecificExitCode;
        public int dwCheckPoint;
        public int dwWaitHint;
    };

    public enum Win32Error : int
    {
        NO_ERROR = 0,
        ERROR_APP_INIT_FAILURE = 575,
        ERROR_FATAL_APP_EXIT = 713,
        ERROR_SERVICE_NOT_ACTIVE = 1062,
        ERROR_EXCEPTION_IN_SERVICE = 1064,
        ERROR_SERVICE_SPECIFIC_ERROR = 1066,
        ERROR_PROCESS_ABORTED = 1067,
    };

    public class PSMessageAgent
    {

        static private Thread asyncThread;
        static private PSMessageAgent instance;

        private PipeSecurity pipeSec;

        static public void Init(string[] messageWriteAccess)
        {
            if (instance != null)
                return;

            instance = new PSMessageAgent();

            instance.pipeSec = new PipeSecurity();

            foreach (string accessor in messageWriteAccess)
            {
                instance.pipeSec.AddAccessRule(
                    new PipeAccessRule(
                        new NTAccount(accessor),
                        PipeAccessRights.ReadWrite | PipeAccessRights.Synchronize,
                        AccessControlType.Allow
                    )
                );
            }

            asyncThread = new Thread(new ThreadStart(instance.RunAsync));
            asyncThread.Start();
            while (!asyncThread.IsAlive) ;
        }

        static public void Abort()
        {
            if (instance == null)
                return;

            asyncThread.Abort();

            while (asyncThread.IsAlive) ;
        }

        static public PSMessageAgent GetInstance()
        {
            if (instance == null)
                Init(new string[] { });

            return instance;
        }

        public delegate void EventHandler(string message);
        public event EventHandler HasMessage;

        public bool IsReady = false;

        public void RunAsync()
        {
            while (true)
            {
                using (NamedPipeServerStream pipe = new NamedPipeServerStream("svcPipe_" + PSService.Instance.ServiceName, PipeDirection.In, 1, PipeTransmissionMode.Message, PipeOptions.None, 0, 0, instance.pipeSec))
                {
                    IsReady = true;

                    pipe.WaitForConnection();

                    using (StreamReader sr = new StreamReader(pipe))
                    {
                        while (sr.Peek() == -1)
                            continue;

                        HasMessage(sr.ReadLine());
                    }
                }
            }
        }
    }

    public class PSService : ServiceBase
    {
        private readonly string _ServiceName = @"%ServiceName%";
        private readonly string _InstallPath = @"%InstallPath%";
        private readonly string _DataPath = @"%DataPath%";
        private readonly string _LogRoot = @"%LogRoot%";

        private readonly bool _SourceIsAvailable = @"%SourceIsAvailable%" == "true";

        public static PSService Instance { get; private set; }

        private string ServiceLogPath;

        private ServiceStatus ServiceStatus;
        private PSObject ServiceDefinition;
        private Runspace ServiceRunspace;

        private string ServiceScript_Process;

        public PSService()
        {
            ServiceName = _ServiceName;
            CanStop = true;
            CanPauseAndContinue = false;
            AutoLog = true;

            Instance = this;

            if (!Directory.Exists(_DataPath))
                Directory.CreateDirectory(_DataPath);

            if (!Directory.Exists(_LogRoot))
                Directory.CreateDirectory(_LogRoot);

            ServiceLogPath = Path.Combine(_LogRoot, _ServiceName + "_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss_fff") + ".log");

            WriteToServiceLog("Service '" + _ServiceName + "' w/ exe '" + Assembly.GetEntryAssembly().Location + "'.");
            WriteToServiceLog("Running as '" + WindowsIdentity.GetCurrent().Name + "'.");
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        private void Process(object trigger)
        {
            Pipeline pl = ServiceRunspace.CreatePipeline();

            Command trgCmd = new Command("param($trigger)", true);
            trgCmd.Parameters.Add(new CommandParameter("trigger", trigger));

            pl.Commands.Add(trgCmd);
            pl.Commands.AddScript(ServiceScript_Process);

            try
            {
                pl.Invoke();
            }
            catch (ActionPreferenceStopException e)
            {
                WriteToServiceLog("Exception: " + e.ErrorRecord.Exception.Message);

                throw e;
            }
            catch (Exception e)
            {
                WriteToServiceLog("Exception: " + e.Message);

                throw e;
            }
        }

        int IntervalCount = 0;

        private void proc_OnInterval(object sender, System.Timers.ElapsedEventArgs args)
        {
            IntervalCount++;

            WriteToServiceLog("Interval trigger #" + IntervalCount + ". Invoking 'Process'. [");
            Process(new
            {
                Source = "Interval",
                IntervalCount = IntervalCount
            });
            WriteToServiceLog("] Finished 'Process'.");
        }

        private void proc_OnMessage(string message)
        {
            WriteToServiceLog("Message trigger. Invoking 'Process'. [");
            Process(new
            {
                Source = "Message",
                Message = message
            });
            WriteToServiceLog("] Finished 'Process'.");
        }

        private static string GetResourceAsString(string resourceName)
        {
            return new StreamReader(
                Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName),
                Encoding.UTF8
            ).ReadToEnd();
        }

        private void WriteToServiceLogBase(string source, string message)
        {
            File.AppendAllLines(ServiceLogPath, new string[] { DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + source.PadRight(6) + " | " + message });
        }

        private void WriteToServiceLog(string message)
        {
            WriteToServiceLogBase("SvcMgr", message);
        }

        public void WriteToServiceLogFromScript(string message)
        {
            WriteToServiceLogBase("Svc", "  " + message);
        }

        protected override void OnStart(string[] args)
        {
            WriteToServiceLog("Starting.");
            ServiceStatus.dwServiceType = ServiceType.SERVICE_WIN32_OWN_PROCESS;
            ServiceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            ServiceStatus.dwWin32ExitCode = 0;
            ServiceStatus.dwWaitHint = 2000;
            SetServiceStatus(ServiceHandle, ref ServiceStatus);

            try
            {
                WriteToServiceLog("Loading service definition from embedded resources.");

                string svcCfgCmd = GetResourceAsString("PSServiceManager.ServiceDefinition.ConfigurationCommand.ps1");
                string svcDefEmbedded = GetResourceAsString("PSServiceManager.ServiceDefinition.ps1");

                if (_SourceIsAvailable)
                {
                    string svcDefFromFile = File.ReadAllText(
                        Path.Combine(
                            _InstallPath,
                            _ServiceName + ".ServiceDefinition.ps1"
                        ),
                        Encoding.UTF8
                    );

                    if (svcDefEmbedded != svcDefFromFile)
                        throw new Exception("The embedded service definition did not match the content of the exposed definition file.");
                }

                using (Runspace defRs = RunspaceFactory.CreateRunspace())
                {
                    defRs.Open();

                    defRs.CreatePipeline(svcCfgCmd).Invoke();
                    ServiceDefinition = defRs.CreatePipeline(svcDefEmbedded).Invoke()[0];
                    ServiceScript_Process = (string) ServiceDefinition.Properties["Process"].Value;
                }

                WriteToServiceLog("Initializing runspace.");
                ServiceRunspace = RunspaceFactory.CreateRunspace();
                ServiceRunspace.Open();

                ServiceRunspace.CreatePipeline("$ErrorActionPreference = \"Stop\"").Invoke();
                ServiceRunspace.CreatePipeline("function Write-Log ($message) {[PSServiceManager.PSService]::Instance.WriteToServiceLogFromScript($message)}").Invoke();

                Pipeline pl = ServiceRunspace.CreatePipeline();
                Command svcCmd = new Command("param($service)", true);
                svcCmd.Parameters.Add(new CommandParameter("service", new
                {
                    DataPath = _DataPath
                }));
                pl.Commands.Add(svcCmd);
                pl.Invoke();

                string begin = (string) ServiceDefinition.Properties["Begin"].Value;

                if (begin.Length > 0)
                {
                    WriteToServiceLog("Invoking 'Begin'. [");
                    ServiceRunspace.CreatePipeline(begin).Invoke();
                    WriteToServiceLog("] Finished 'Begin'.");
                }
                else
                    WriteToServiceLog("Would invoke 'Begin' if defined.");

                WriteToServiceLog("Initializing process trigger(s).");
                if ((bool)ServiceDefinition.Properties["ProcessOnTimer"].Value)
                {
                    System.Timers.Timer timer = new System.Timers.Timer();
                    timer.Interval = (double)ServiceDefinition.Properties["TimerInterval"].Value;
                    timer.Elapsed += new System.Timers.ElapsedEventHandler(proc_OnInterval);
                    timer.Start();
                }
                if ((bool)ServiceDefinition.Properties["ProcessOnMessage"].Value)
                {
                    PSMessageAgent.Init((string[])ServiceDefinition.Properties["MessageWriteAccess"].Value);
                    PSMessageAgent.GetInstance().HasMessage += proc_OnMessage;

                    while (!PSMessageAgent.GetInstance().IsReady)
                        continue;
                }

                WriteToServiceLog("Started.");

                ServiceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            }
            catch (ActionPreferenceStopException e)
            {
                WriteToServiceLog("Exception: " + e.ErrorRecord.Exception.Message);

                ServiceStatus.dwWin32ExitCode = (int)(Win32Error.ERROR_APP_INIT_FAILURE);

                throw e;
            }
            catch (Exception e)
            {
                WriteToServiceLog("Exception: " + e.Message);

                ServiceStatus.dwWin32ExitCode = (int)(Win32Error.ERROR_APP_INIT_FAILURE);

                throw e;
            }
            finally
            {
                ServiceStatus.dwWaitHint = 0;
                SetServiceStatus(ServiceHandle, ref ServiceStatus);
            }
        }

        protected override void OnStop()
        {
            try
            {
                WriteToServiceLog("Stopping.");

                string end = (string) ServiceDefinition.Properties["End"].Value;

                if (end.Length > 0)
                {
                    WriteToServiceLog("Invoking 'End'. [");
                    ServiceRunspace.CreatePipeline(end).Invoke();
                    WriteToServiceLog("] Finished 'end'.");
                }
                else
                    WriteToServiceLog("Would invoke 'End' if defined.");

                // Success. Set the service state to Stopped.
                ServiceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;

                WriteToServiceLog("Stopped.");
            }
            catch (ActionPreferenceStopException e)
            {
                WriteToServiceLog("Exception: " + e.ErrorRecord.Exception.Message);

                ServiceStatus.dwWin32ExitCode = (int)(Win32Error.ERROR_APP_INIT_FAILURE);

                throw e;
            }
            catch (Exception e)
            {
                WriteToServiceLog("Exception: " + e.Message);

                ServiceStatus.dwWin32ExitCode = (int)(Win32Error.ERROR_APP_INIT_FAILURE);

                throw e;
            }
            finally
            {
                ServiceStatus.dwWaitHint = 0;
                SetServiceStatus(ServiceHandle, ref ServiceStatus);
            }
        }

        public static void Main()
        {
            Run(new PSService());
        }
    }
}