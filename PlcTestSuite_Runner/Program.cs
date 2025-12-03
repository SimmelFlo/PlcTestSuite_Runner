using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using EnvDTE;
using TCatSysManagerLib;

namespace PlcTestSuite_Runner
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                string callDirectory = System.IO.Directory.GetCurrentDirectory();
                string projectDirectory = System.IO.Directory.GetParent(callDirectory).FullName;
                Console.WriteLine($"[INFO] Call directory: {callDirectory}");
                Console.WriteLine($"[INFO] Project directory: {projectDirectory}");

                if (args.Length >= 2)
                {
                    Automation a = new Automation();
                    string prj = projectDirectory + @"\" + args[0];
                    string projectName = args[1];
                    Console.WriteLine($"[INFO] Project: {prj}");
                    Console.WriteLine($"[INFO] Target project name: {projectName}");
                    a.ActivateProject(prj, projectName);
                    Console.WriteLine("[SUCCESS] Automation completed successfully.");
                }
                else
                {
                    Console.WriteLine("[ERROR] Usage: PlcTestSuite_Runner.exe <solution-file.sln> <project-name>");
                    Environment.Exit(1);
                }
            }
            catch (COMException comEx)
            {
                Console.WriteLine($"[ERROR] COM Exception occurred: {comEx.Message}");
                Console.WriteLine($"[ERROR] Error Code: 0x{comEx.ErrorCode:X}");
                Console.WriteLine($"[ERROR] Stack Trace: {comEx.StackTrace}");
                Environment.Exit(1);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Unexpected error occurred: {ex.Message}");
                Console.WriteLine($"[ERROR] Stack Trace: {ex.StackTrace}");
                Environment.Exit(1);
            }
        }
    }
    class Automation
    {
        DTE dte = null;
        Solution solution = null;
        dynamic envDteProject = null;
        ITcSysManager17 sysManager = null;

        public Automation()
        {
            MessageFilter.Register();
        }
        ~Automation()
        {
            // revoke COM message filter
            MessageFilter.Revoke();
        }

        public void ActivateProject(string project, string projectName)
        {
            try
            {
                Console.WriteLine("[STEP] Initializing TwinCAT XAE Shell...");
                Type t = System.Type.GetTypeFromProgID("TcXaeShell.DTE.17.0");
                dte = (DTE)Activator.CreateInstance(t);
                dte.SuppressUI = false;
                dte.MainWindow.Visible = true;
                solution = dte.Solution;

                Console.WriteLine("[STEP] Opening Project...");
                solution.Open(project);
                Console.WriteLine("[INFO] Waiting for project to load...");
                System.Threading.Thread.Sleep(60000);

                bool projectFound = false;
                foreach (var pp in solution.Projects)
                {
                    Project prj = pp as Project;
                    Console.WriteLine($"[INFO] Found project: {prj.Name}");

                    if (prj.Name == projectName)
                    {
                        projectFound = true;
                        envDteProject = prj;
                        sysManager = envDteProject.Object as ITcSysManager17;

                        if (sysManager != null)
                        {
                            ITcSmTreeItem plc = sysManager.LookupTreeItem("TIPC");
                            foreach (ITcSmTreeItem plcProject in plc)
                            {
                                ITcPlcProject iecProjectRoot = (ITcPlcProject)plcProject;
                                iecProjectRoot.BootProjectAutostart = true;
                            }

                            Console.WriteLine($"[INFO] {projectName} found!");
                            try
                            {
                                Console.WriteLine("[STEP] Build TwinCAT...");
                                sysManager.BuildTargetPlatform("TwinCAT RT (x64)");
                                Console.WriteLine("[STEP] Activate Configuration...");
                                sysManager.ActivateConfiguration();
                                Console.WriteLine("[STEP] Start TwinCAT...");
                                sysManager.StartRestartTwinCAT();

                                Console.WriteLine("[INFO] Waiting for TwinCAT to start...");
                                if (sysManager.IsTwinCATStarted())
                                {
                                    Console.WriteLine("[SUCCESS] TwinCAT activation completed.");
                                }
                                System.Threading.Thread.Sleep(20000);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[ERROR] Build & activate failed: {ex.Message}");
                                throw;
                            }    
                        }
                        else
                        {
                            throw new InvalidOperationException($"Failed to get ITcSysManager3 interface from {projectName}.");
                        }
                    }
                }

                if (!projectFound)
                {
                    throw new InvalidOperationException($"{projectName} not found in solution.");
                }
            }
            catch (COMException comEx)
            {
                Console.WriteLine($"[ERROR] COM Exception in ActivateProject: {comEx.Message}");
                Console.WriteLine($"[ERROR] Error Code: 0x{comEx.ErrorCode:X}");
                throw;
            }
            catch (System.IO.FileNotFoundException fileEx)
            {
                Console.WriteLine($"[ERROR] Project file not found: {fileEx.Message}");
                throw;
            }
            catch (InvalidOperationException invEx)
            {
                Console.WriteLine($"[ERROR] {invEx.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Unexpected error in ActivateProject: {ex.Message}");
                throw;
            }
            finally
            {
                // Cleanup will happen in destructor
                Console.WriteLine("[INFO] ActivateProject method completed.");
            }
        }
    }

    public class MessageFilter : IOleMessageFilter
    {
        // Class containing the IOleMessageFilter
        // thread error-handling functions.
        // 

        /// <summary>
        /// Start the filter
        /// </summary>
        public static void Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            int test = CoRegisterMessageFilter(newFilter, out oldFilter);

            if (test != 0)
            {
                Console.WriteLine(string.Format("CoRegisterMessageFilter failed with error : {0}", test));
            }
        }

        /// <summary>
        /// Done with the filter, close it.
        /// </summary>
        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            int test = CoRegisterMessageFilter(null, out oldFilter);
        }

        // IOleMessageFilter functions.

        /// <summary>
        /// Handles the in coming thread requests.
        /// </summary>
        /// <param name="dwCallType">Type of the dw call.</param>
        /// <param name="hTaskCaller">The h task caller.</param>
        /// <param name="dwTickCount">The dw tick count.</param>
        /// <param name="lpInterfaceInfo">The lp interface info.</param>
        /// <returns></returns>
        int IOleMessageFilter.HandleInComingCall(int dwCallType,
          System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
          lpInterfaceInfo)
        {
            //Return the flag SERVERCALL_ISHANDLED.
            return 0;
        }

        /// <summary>
        /// Retries the rejected call.
        /// </summary>
        /// <param name="hTaskCallee">The h task callee.</param>
        /// <param name="dwTickCount">The dw tick count.</param>
        /// <param name="dwRejectType">Type of the dw reject.</param>
        /// <returns></returns>
        int IOleMessageFilter.RetryRejectedCall(System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // Thread call was rejected, so try again.

            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER.
            {
                // Retry the thread call immediately if return >=0 & 
                // <100.
                return 99;
            }
            // Too busy; cancel call.
            return -1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hTaskCallee">The h task callee.</param>
        /// <param name="dwTickCount">The dw tick count.</param>
        /// <param name="dwPendingType">Type of the dw pending.</param>
        /// <returns></returns>
        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            //Return the flag PENDINGMSG_WAITDEFPROCESS.
            return 2;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    }

    /// <summary>
    /// Definition of the IOleMessageFilter interface
    /// </summary>
    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        /// <summary>
        /// Handles the in coming call.
        /// </summary>
        /// <param name="dwCallType">Type of the dw call.</param>
        /// <param name="hTaskCaller">The h task caller.</param>
        /// <param name="dwTickCount">The dw tick count.</param>
        /// <param name="lpInterfaceInfo">The lp interface info.</param>
        /// <returns></returns>
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        /// <summary>
        /// Retries the rejected call.
        /// </summary>
        /// <param name="hTaskCallee">The h task callee.</param>
        /// <param name="dwTickCount">The dw tick count.</param>
        /// <param name="dwRejectType">Type of the dw reject.</param>
        /// <returns></returns>
        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        /// <summary>
        /// Messages the pending.
        /// </summary>
        /// <param name="hTaskCallee">The h task callee.</param>
        /// <param name="dwTickCount">The dw tick count.</param>
        /// <param name="dwPendingType">Type of the dw pending.</param>
        /// <returns></returns>
        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }
}
