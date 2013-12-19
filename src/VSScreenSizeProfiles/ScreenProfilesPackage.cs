using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace devcoach.Tools.ScreenProfiles
{
  /// <summary>
  /// This is the class that implements the package exposed by this assembly.
  ///
  /// The minimum requirement for a class to be considered a valid package for Visual Studio
  /// is to implement the IVsPackage interface and register itself with the shell.
  /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
  /// to do it: it derives from the Package class that provides the implementation of the
  /// IVsPackage interface and uses the registration attributes defined in the framework to
  /// register itself and its components with the shell.
  /// </summary>
  // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
  // a package.
  [PackageRegistration(UseManagedResourcesOnly = true)]
  // This attribute is used to register the information needed to show this package
  // in the Help/About dialog of Visual Studio.
  [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
  // This attribute is needed to let the shell know that this package exposes some menus.
  [ProvideMenuResource("Menus.ctmenu", 1)]
  [Guid(GuidList.guidScreenProfilesPkgString)]
  [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
  [ProvideAutoLoad(UIContextGuids80.NoSolution)]
  public sealed class ScreenProfilesPackage : Package
  {
    /// <summary>
    /// Default constructor of the package.
    /// Inside this method you can place any initialization code that does not require
    /// any Visual Studio service because at this point the package object is created but
    /// not sited yet inside Visual Studio environment. The place to do all the other
    /// initialization is the Initialize method.
    /// </summary>
    public ScreenProfilesPackage()
    {
      Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
    }

    private static readonly object _s_applicationLock = new object();
    public static DTE2 Application { get; private set; }
    public static Events Events { get; private set; }
    public static DTEEvents DTEEvents { get; private set; }
    public static WindowEvents WindowEvents { get; private set; }
    /////////////////////////////////////////////////////////////////////////////
    // Overridden Package Implementation
    
    #region Initialize()

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    protected override void Initialize()
    {
      Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
      base.Initialize();

      lock (_s_applicationLock)
      {
        Application = (DTE2)GetService(typeof(SDTE));
        Events = Application.Application.Events;
        DTEEvents = Events.DTEEvents;
        WindowEvents = Events.WindowEvents;

        WindowEvents.WindowMoved += WindowEvents_WindowMoved;
      }

      // Add our command handlers for menu (commands must exist in the .vsct file)
      OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (null != mcs)
      {
        // Create the command for the menu item.
        CommandID menuCommandID = new CommandID(GuidList.guidScreenProfilesCmdSet, (int)PkgCmdIDList.cmdidSaveScreenLayout);
        MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
        mcs.AddCommand(menuItem);
      }
    }
    #endregion

    #region WindowEvents_WindowMoved()
    void WindowEvents_WindowMoved(
      Window window, int top, int left, int width, int height)
    {
      var key = GetWindowLayoutIdentifier();
      if (window.Type == vsWindowType.vsWindowTypeMainWindow)
      {
        var bar = Application.StatusBar;

        if (!string.IsNullOrWhiteSpace(_lastImported) &&
            _lastImported.Equals(key, StringComparison.OrdinalIgnoreCase))
        {
          return;
        }
        _lastImported = key;
        bar.Text = "Loading screen layout";
        bar.Highlight(true);
        if (ImportSettings(key))
        {
          bar.Text = "Loaded screen layout";
          bar.Highlight(true);
        }
      }
      else
      {
        var vsSettingsFile = GetSettingsFile(key);
        if (File.Exists(vsSettingsFile))
        {
          // Update existing settings on resize etc...
          ExportSettings(key);
        }
      }
    }
    #endregion

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool GetWindowRect(IntPtr hWnd, out Rectangle lpRect);

    #region MenuItemCallback()
    /// <summary>
    /// This function is the callback used to execute a command when the a menu item is clicked.
    /// See the Initialize method to see how the menu item is associated to this function using
    /// the OleMenuCommandService service and the MenuCommand class.
    /// </summary>
    private void MenuItemCallback(object sender, EventArgs e)
    {
      var bar = Application.StatusBar;
      bar.Text = "Saved current screen layout";
      bar.Highlight(true);
      ExportSettings(GetWindowLayoutIdentifier());
    } 
    #endregion

    #region GetWindowLayoutIdentifier()
    private static string GetWindowLayoutIdentifier()
    {
      Rectangle rect;
      var window =
        System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
      GetWindowRect(window, out rect);
      var msg = new List<string>();

      rect.Width = (rect.Width - rect.Left);
      rect.Height = (rect.Height - rect.Top);
      msg.Add(
        (
          (Application.DTE.Debugger.CurrentMode == dbgDebugMode.dbgDesignMode)
            ? "Design"
            : "Debug"
          )
        );
      msg.Add(
        rect.Width + "x" + rect.Height +
        "@" +
        rect.Top + "x" + rect.Left
        );
      msg.AddRange(
        Screen.AllScreens.Select(
          screen =>
            screen.Bounds.Width +
            "x" +
            screen.Bounds.Height +
            (screen.Primary ? "P" : "S") +
            (screen.Bounds.IntersectsWith(rect) ? "A" : "")
          )
        );

      var text = string.Join("-", msg);
      return text;
    }
    #endregion

    private string _lastImported = null;

    #region ImportSettings()
    private bool ImportSettings(string theme)
    {
      if (string.IsNullOrWhiteSpace(theme)) return false;

      using (var hkcu = Registry.CurrentUser)
      {

        using (var vs =
            hkcu.OpenSubKey(
              @"Software\Microsoft\VisualStudio\" + Application.Version))
        {
          if (vs == null) return false;

          var visualStudioLocation =
              vs.GetValue("VisualStudioLocation");
          var vsSettingsFile =
              Path.Combine(
                  (string)visualStudioLocation,
                  "Settings",
                  string.Concat(theme, ".vssettings"));

          if (!File.Exists(vsSettingsFile))
          {
            Trace.WriteLine(
              "The theme settings file " + vsSettingsFile + " does not exist!");
            return false;
          }

          Application.DTE.ExecuteCommand(
              "Tools.ImportandExportSettings",
              string.Concat("/import:\"", vsSettingsFile, "\""));
        }
      }
      return true;
    } 
    #endregion

    #region ExportSettings()
    private void ExportSettings(string theme)
    {
      if (string.IsNullOrWhiteSpace(theme)) return;

      var vsSettingsFile = GetSettingsFile(theme);

      Application.DTE.ExecuteCommand(
        "Tools.ImportandExportSettings",
        string.Concat("/export:\"", vsSettingsFile, "\""));

      StipSettingsFileForWindowLayout(vsSettingsFile);

    } 
    #endregion

    #region GetSettingsFile()
    private static string GetSettingsFile(string theme)
    {
      using (var hkcu = Registry.CurrentUser)
      {
        using (var vs =
          hkcu.OpenSubKey(
            @"Software\Microsoft\VisualStudio\" + Application.Version))
        {
          if (vs == null) return null;

          var visualStudioLocation =
            vs.GetValue("VisualStudioLocation");
          return
             Path.Combine(
               (string)visualStudioLocation,
               "Settings",
               string.Concat(theme, ".vssettings"));
        }
      }
    } 
    #endregion

    #region StipSettingsFileForWindowLayout()
    private static void StipSettingsFileForWindowLayout(string vsSettingsFile)
    {
      var doc = new XmlDocument();
      doc.Load(vsSettingsFile);

      var userSettingsNode = doc.SelectSingleNode("/UserSettings");

      if (userSettingsNode != null)
      {
        foreach (var node in
          userSettingsNode.ChildNodes.Cast<XmlNode>().ToList())
        {
          if (!node.Name.Equals(
            "Category",
            StringComparison.OrdinalIgnoreCase) || node.Attributes == null)
          {
            for (int i = 0; i < node.ChildNodes.Count; i++)
            {
              node.RemoveChild(node.ChildNodes[i--]);
            }

            continue;
          }

          var nameAtt = node.Attributes["name"];
          if (nameAtt != null)
          {
            if (!nameAtt.Value.Equals(
              "Environment_Group",
              StringComparison.OrdinalIgnoreCase))
            {
              userSettingsNode.RemoveChild(node);
            }
            else
            {
              for (int i = 0; i < node.ChildNodes.Count; i++)
              {
                var categoryNode = node.ChildNodes[i];
                if (categoryNode.Attributes != null)
                {
                  var catNameAtt = categoryNode.Attributes["name"];
                  if (!catNameAtt.Value.Equals(
                    "Environment_WindowLayout",
                    StringComparison.OrdinalIgnoreCase))
                  {
                    node.RemoveChild(node.ChildNodes[i--]);
                  }
                }
                else
                {
                  node.RemoveChild(node.ChildNodes[i--]);
                }
              }
            }
          }
        }
      }
      doc.Save(vsSettingsFile);
    } 
    #endregion
  }
}
