using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.Diagnostics;

namespace VSIXDpiAwareness
{
   internal sealed class CommandDpiAwareness
   {
      private const string REGISTRY_KEY = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\devenv.exe";
      private const string NAME_DPI_AWARENESS = "dpiAwareness";

      private const int DPI_AWARE = 1;
      private const int DPI_VIRTUALIZED = 0;

      private readonly Package m_package;
      private readonly IServiceProvider m_serviceProvider;

      public CommandDpiAwareness(Package package)
      {
         // This value must be synchronized with the .vsct file 
         const int CommandId = 0x0100;

         // This value must be synchronized with the .vsct file 
         Guid CommandSet = new Guid("a461dc99-c7c1-484f-b420-5b6b72bad679");
         OleMenuCommandService commandService;
         CommandID menuCommandID;
         OleMenuCommand oleMenuCommand;

         m_package = package ?? throw new ArgumentNullException(nameof(package));

         m_serviceProvider = package;

         commandService = m_serviceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
         if (commandService != null)
         {
            menuCommandID = new CommandID(CommandSet, CommandId);
            oleMenuCommand = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
            commandService.AddCommand(oleMenuCommand);

            if (IsDpiAware())
            {
               oleMenuCommand.Checked = true;
            }
            else
            {
               oleMenuCommand.Checked = false;
            }
         }
      }

      private void MenuItemCallback(object sender, EventArgs e)
      {
         const string TITLE = "DpiAwareness extension";

         OleMenuCommand oleMenuCommand;
         int newDpiAwareness;

         try
         {
            oleMenuCommand = sender as OleMenuCommand;
            if (oleMenuCommand.Checked)
            {
               newDpiAwareness = DPI_VIRTUALIZED;
            }
            else
            {
               newDpiAwareness = DPI_AWARE;
            }

            ToggleDpiAwareness(newDpiAwareness);

            VsShellUtilities.ShowMessageBox(m_serviceProvider, 
               "Visual Studio will be restarted for the change to take effect.", TITLE,
               OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

            RestartVisualStudio();
         }
         catch (Exception ex)
         {
            VsShellUtilities.ShowMessageBox(m_serviceProvider, ex.Message, TITLE,
                OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
         }
      }

      private bool IsDpiAware()
      {
         RegistryKey registryKey;
         int dpiAwareness;
         bool result = true;

         registryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(REGISTRY_KEY);

         if (registryKey != null)
         {
            try
            {
               dpiAwareness = (int)registryKey.GetValue(NAME_DPI_AWARENESS, DPI_AWARE);

               if (dpiAwareness == DPI_VIRTUALIZED)
               {
                  result = false;
               }
            }
            finally
            {
               registryKey.Close();
            }
         }
         return result;
      }

      private void ToggleDpiAwareness(int newDpiAwareness)
      {
         string tempFileFullName;

         tempFileFullName = WriteTempRegFile(newDpiAwareness);

         MergeTempRegFile(tempFileFullName);
      }

      private void MergeTempRegFile(string tempFileFullName)
      {
         Process process;
         ProcessStartInfo processStartInfo;

         processStartInfo = new ProcessStartInfo()
         {
            FileName = "regedit.exe",
            Verb = "runas",
            Arguments = "-s " + "\"" + tempFileFullName + "\"",
            WindowStyle = ProcessWindowStyle.Normal,
            UseShellExecute = true
         };

         process = Process.Start(processStartInfo);
         if (process != null)
         {
            process.Dispose();
         }
      }

      private string WriteTempRegFile(int newDpiAwareness)
      {
         const string TEMP_FILE_NAME = "DpiAwarenessTemp.reg";

         string regFileContent;
         string tempFolderFullName;
         string tempFileFullName;

         tempFolderFullName = System.IO.Path.GetTempPath();
         tempFileFullName = System.IO.Path.Combine(tempFolderFullName, TEMP_FILE_NAME);
         if (System.IO.File.Exists(tempFileFullName))
         {
            System.IO.File.Delete(tempFileFullName);
         }

         regFileContent = $@"Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\{REGISTRY_KEY}]
""{NAME_DPI_AWARENESS}"" = dword:" + newDpiAwareness.ToString("X8");

         System.IO.File.WriteAllText(tempFileFullName, regFileContent);

         return tempFileFullName;
      }

      private void RestartVisualStudio()
      {
         object shell;
         IVsShell3 shell3;
         IVsShell4 shell4;

         shell = m_serviceProvider.GetService(typeof(SVsShell));
         shell3 = shell as IVsShell3;
         shell4 = shell as IVsShell4;

         shell3.IsRunningElevated(out bool elevated);

         if (elevated)
         {
            shell4.Restart((uint)__VSRESTARTTYPE.RESTART_Elevated);
         }
         else
         {
            shell4.Restart((uint)__VSRESTARTTYPE.RESTART_Normal);
         }
      }
   }
}
