using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace VSIXDpiAwareness
{
   [PackageRegistration(UseManagedResourcesOnly = true)]
   // Info on this package for Help/About
   [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
   [Guid(VSPackageDpiAwareness.PACKAGE_GUID_STRING)]
   [ProvideMenuResource("Menus.ctmenu", 1)]
   // Package must be loaded on VS startup to set the checked/unchecked state of the command
   [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)] 
   public sealed class VSPackageDpiAwareness : Package
   {
      private const string PACKAGE_GUID_STRING = "24c3cb3e-8420-4105-99ed-01643b99490a";

      private CommandDpiAwareness m_commandDpiAwareness;

      public VSPackageDpiAwareness()
      {
      }

      protected override void Initialize()
      {
         base.Initialize();

         m_commandDpiAwareness = new CommandDpiAwareness(this);
      }
   }
}
