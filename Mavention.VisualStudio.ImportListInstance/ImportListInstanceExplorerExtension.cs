using System;
using System.ComponentModel.Composition;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;
using Mavention.VisualStudio.ImportListInstance.Commands;
using Microsoft.VisualStudio.SharePoint;
using Microsoft.VisualStudio.SharePoint.Explorer;
using Microsoft.VisualStudio.SharePoint.Explorer.Extensions;

namespace Mavention.VisualStudio.ImportListInstance {
    [Export(typeof(IExplorerNodeTypeExtension))]
    [ExplorerNodeType(ExtensionNodeTypes.ListNode)]
    public class ImportListInstanceExplorerExtension : IExplorerNodeTypeExtension {
        public static readonly XNamespace SPXN = "http://schemas.microsoft.com/sharepoint/";

        public void Initialize(IExplorerNodeType nodeType) {
            nodeType.NodeMenuItemsRequested += new EventHandler<ExplorerNodeMenuItemsRequestedEventArgs>(nodeType_NodeMenuItemsRequested);
        }

        void nodeType_NodeMenuItemsRequested(object sender, ExplorerNodeMenuItemsRequestedEventArgs e) {
            e.MenuItems.Add("Import List Instance").Click += new EventHandler<MenuItemEventArgs>(ImportListInstanceExplorerExtension_Click);
        }

        void ImportListInstanceExplorerExtension_Click(object sender, MenuItemEventArgs e) {
            IExplorerNode listNode = e.Owner as IExplorerNode;

            if (listNode != null) {
                try {
                    IListNodeInfo listInfo = listNode.Annotations.GetValue<IListNodeInfo>();

                    string listInstanceContents = listNode.Context.SharePointConnection.ExecuteCommand<Guid, string>(CommandIds.GetListInstanceXmlCommandId, listInfo.Id);
                    XNamespace xn = "http://schemas.microsoft.com/sharepoint/";
                    string moduleContents = new XElement(xn + "Elements",
                        XElement.Parse(listInstanceContents)).ToString().Replace(" xmlns=\"\"", String.Empty);

                    EnvDTE.Project activeProject = Utils.GetActiveProject();
                    if (activeProject != null) {
                        ISharePointProjectService projectService = listNode.ServiceProvider.GetService(typeof(ISharePointProjectService)) as ISharePointProjectService;
                        ISharePointProject activeSharePointProject = projectService.Projects[activeProject.FullName];
                        if (activeSharePointProject != null) {
                            string spiName = listInfo.Title;
                            ISharePointProjectItem listInstanceProjectItem = null;
                            bool itemCreated = false;
                            int counter = 0;

                            do {
                                try {
                                    listInstanceProjectItem = activeSharePointProject.ProjectItems.Add(spiName, "Microsoft.VisualStudio.SharePoint.ListInstance");
                                    itemCreated = true;
                                }
                                catch (ArgumentException) {
                                    spiName = String.Format("{0}{1}", listInfo.Title, ++counter);
                                }
                            }
                            while (!itemCreated);

                            string elementsXmlFullPath = Path.Combine(listInstanceProjectItem.FullPath, "Elements.xml"); 
                            System.IO.File.WriteAllText(elementsXmlFullPath, moduleContents);
                            ISharePointProjectItemFile elementsXml = listInstanceProjectItem.Files.AddFromFile("Elements.xml");
                            elementsXml.DeploymentType = DeploymentType.ElementManifest;
                            elementsXml.DeploymentPath = String.Format(@"{0}\", spiName);
                            listInstanceProjectItem.DefaultFile = elementsXml;

                            Utils.OpenFile(Path.Combine(elementsXmlFullPath));
                        }
                    }
                }
                catch (Exception ex) {
                    listNode.Context.ShowMessageBox(String.Format("The following exception occured while exporting List Instance: {0}", ex.Message), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
