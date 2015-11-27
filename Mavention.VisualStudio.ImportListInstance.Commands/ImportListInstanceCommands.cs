using System;
using System.Xml.Linq;
using Microsoft.SharePoint;
using Microsoft.VisualStudio.SharePoint.Commands;

namespace Mavention.VisualStudio.ImportListInstance.Commands {
    public static class ImportListInstanceCommands {
        [SharePointCommand(CommandIds.GetListInstanceXmlCommandId)]
        public static string GetListInstanceXml(ISharePointCommandContext context, Guid listId) {
            string listInstanceXml = String.Empty;

            context.Logger.WriteLine("Exporting List Instance...", LogCategory.Status);

            SPWeb web = context.Web;
            SPList list = web.Lists[listId];

            XElement xListInstance = new XElement("ListInstance",
                new XAttribute("Title", list.Title),
                new XAttribute("Description", list.Description),
                new XAttribute("Url", list.RootFolder.Url),
                new XAttribute("TemplateType", (int)list.BaseTemplate),
                new XAttribute("FeatureId", list.TemplateFeatureId),
                new XAttribute("OnQuickLaunch", list.OnQuickLaunch.ToString().ToUpper()));

            SPListItemCollection items = list.Items;

            if (items.Count > 0) {
                XElement xData = new XElement("Data");
                XElement xRows = new XElement("Rows");

                foreach (SPListItem item in items) {
                    XElement xRow = new XElement("Row");
                    foreach (SPField field in item.Fields) {
                        try {
                            string fieldValue = null;
                            object fieldValueRaw = null;

                            try {
                                fieldValueRaw = item[field.Id];
                            }
                            catch { }

                            if (fieldValueRaw != null) {
                                if (field.FieldValueType == typeof(DateTime)) {
                                    DateTime dateTime = (DateTime)fieldValueRaw;
                                    fieldValue = dateTime.ToUniversalTime().ToString(context.Web.Locale);
                                }
                                else if (field.FieldValueType == typeof(Boolean)) {
                                    fieldValue = (bool)fieldValueRaw ? "1" : "0";
                                }
                                else {
                                    fieldValue = fieldValueRaw.ToString();
                                }
                            }

                            if (!String.IsNullOrEmpty(fieldValue)) {
                                XElement xField = new XElement("Field",
                                    new XAttribute("Name", field.InternalName),
                                    fieldValue);
                                xRow.Add(xField);
                                context.Logger.WriteLine(String.Format("Exported Field: {0}; Value: {1}", field.InternalName, fieldValue), LogCategory.Verbose);
                            }
                        }
                        catch (Exception ex) {
                            context.Logger.WriteLine(String.Format("Exception while exporting field '{0}': {1}", field.InternalName, ex.StackTrace), LogCategory.Error);
                        }
                    }

                    xRows.Add(xRow);
                }

                xData.Add(xRows);
                xListInstance.Add(xData);
            }

            listInstanceXml = xListInstance.ToString();

            context.Logger.WriteLine("List Instance successfully exported", LogCategory.Status);

            return listInstanceXml;
        }
    }
}
