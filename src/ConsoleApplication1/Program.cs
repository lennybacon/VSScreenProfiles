using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ConsoleApplication1
{
  class Program
  {
    static void Main(string[] args)
    {
      var vsSettingsFile =
        @"S:\Dropbox\lennybacon\Settings\Visual Studio 2012\Settings\Design-960x1160-1920x1200PA-1920x1200S.vssettings";

      StipSettingsFileForWindowLayout(vsSettingsFile);
    }

    private static void StipSettingsFileForWindowLayout(string vsSettingsFile)
    {
      var doc = new XmlDocument();
      doc.Load(vsSettingsFile);

      var userSettingsNode = doc.SelectSingleNode("/UserSettings");

      if (userSettingsNode != null)
      {
        foreach (var node in userSettingsNode.ChildNodes.Cast<XmlNode>().ToList())
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
  }
}
