using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class SpfList
    {
        public static void AddField(this List list, string fieldName)
        {
            var clientContext = (ClientContext)list.Context;
            var field = list.ParentWeb.AvailableFields.GetByInternalNameOrTitle(fieldName);
            clientContext.Load(field,
                currentField => currentField.SchemaXml,
                currentField => currentField.InternalName,
                currentField => currentField.Title
            );
            var listFields = list.Fields;
            clientContext.Load(listFields, currentFields => currentFields.Include(currentField => currentField.InternalName));
            clientContext.ExecuteQuery();

            Field existField = listFields.Cast<Field>().FirstOrDefault(currentField => currentField.InternalName == fieldName);
            if (existField != null) return;

            string schemaXml = field.SchemaXml;
            schemaXml = schemaXml.ReplaceXmlAttributeValue("DisplayName", field.InternalName);
            Field listField = list.Fields.AddFieldAsXml(schemaXml, false, AddFieldOptions.AddToAllContentTypes);
            list.Update();
            clientContext.ExecuteQuery();

            schemaXml = schemaXml.ReplaceXmlAttributeValue("DisplayName", field.Title);
            listField.SchemaXml = schemaXml;
            listField.UpdateAndPushChanges(true);
            clientContext.ExecuteQuery();
        }

        public static void RemoveContentType(this List list, string contentTypeName)
        {
            var clientContext = (ClientContext)list.Context;
            var contentTypes = list.ContentTypes;
            clientContext.Load(contentTypes, currentContentTypes => currentContentTypes.Include(currentContentType => currentContentType.Name));
            clientContext.ExecuteQuery();

            ContentType contentType = contentTypes.Cast<ContentType>().FirstOrDefault(currentContentType => currentContentType.Name == contentTypeName);

            if (contentType != null)
            {
                contentType.DeleteObject();
            }
            clientContext.ExecuteQuery();
        }

        public static string ReplaceXmlAttributeValue(this string xml, string attributeName, string value)
        {
            var addAtr = "";
            if (xml.IndexOf("List") == -1)
            {
                addAtr += " List=\"\"";
            }
            if (xml.IndexOf("WebId") == -1)
            {
                addAtr += " WebId=\"\"";
            }
            if (addAtr.Length > 0)
            {
                xml = xml.Replace("<Field", "<Field" + addAtr);
            }

            if (string.IsNullOrEmpty(xml))
            {
                throw new ArgumentNullException("xml");
            }

            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentNullException("value");
            }


            int indexOfAttributeName = xml.IndexOf(attributeName, StringComparison.CurrentCultureIgnoreCase);
            if (indexOfAttributeName == -1)
            {
                throw new ArgumentOutOfRangeException("attributeName", string.Format("Attribute {0} not found in source xml", attributeName));
            }

            int indexOfAttibuteValueBegin = xml.IndexOf('"', indexOfAttributeName);
            int indexOfAttributeValueEnd = xml.IndexOf('"', indexOfAttibuteValueBegin + 1);

            return xml.Substring(0, indexOfAttibuteValueBegin + 1) + value + xml.Substring(indexOfAttributeValueEnd);
        }

        public static void UpdateLookupField(this Web web, List list, string fieldName)
        {
            ClientContext clientContext = (ClientContext)web.Context;
            Field lookupField = web.Fields.GetByInternalNameOrTitle(fieldName);
            clientContext.Load(lookupField, field => field.SchemaXml);
            if (!list.IsPropertyAvailable("Id")) clientContext.Load(list, currentList => currentList.Id);
            if (!web.IsPropertyAvailable("Id")) clientContext.Load(web, currentWeb => currentWeb.Id);
            clientContext.ExecuteQuery();

            lookupField.SchemaXml = lookupField.SchemaXml.ReplaceXmlAttributeValue("List", list.Id.ToString()).ReplaceXmlAttributeValue("WebId", web.Id.ToString());
            lookupField.UpdateAndPushChanges(true);
            clientContext.ExecuteQuery();
        }

        public static void SetFormJSLink(this List list, string formName, string JSLink)
        {
            ClientContext clientContext = (ClientContext)list.Context;
            clientContext.Load(list.Forms, currentForms => currentForms.Include(currentForm => currentForm.ServerRelativeUrl));
            clientContext.ExecuteQuery();

            foreach (Form spForm in list.Forms)
            {
                if (spForm.ServerRelativeUrl.Contains(formName))
                {
                    File formFile = clientContext.Web.GetFileByServerRelativeUrl(spForm.ServerRelativeUrl);
                    LimitedWebPartManager wpManager = formFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    clientContext.Load(wpManager.WebParts,
                        wpDefCollection => wpDefCollection.Include(
                            currentWPDef => currentWPDef.WebPart.Properties
                        )
                    );
                    clientContext.ExecuteQuery();

                    WebPartDefinition wpDef = wpManager.WebParts.Cast<WebPartDefinition>().FirstOrDefault(
                        currentWPDef => currentWPDef.WebPart.Properties["TemplateName"].StringValueOrEmpty() == "ListForm" ||
                                        currentWPDef.WebPart.Properties["TemplateName"].StringValueOrEmpty() == "TaskForm"
                    );

                    if (wpDef != null)
                    {
                        wpDef.WebPart.Properties["JSLink"] = JSLink;
                        wpDef.SaveWebPartChanges();
                        clientContext.ExecuteQuery();
                    }
                }
            }
        }

        public static void RegisterEventReceiver(this List list, string name, string serviceUrl, EventReceiverType eventType, int sequence)
        {
            ClientContext clientContext = (ClientContext)list.Context;
            EventReceiverDefinitionCreationInformation newEventReceiver = new EventReceiverDefinitionCreationInformation()
            {
                EventType = eventType,
                ReceiverName = name,
                ReceiverUrl = serviceUrl,
                SequenceNumber = sequence
            };

            list.EventReceivers.Add(newEventReceiver);
            clientContext.ExecuteQuery();
        }

        public static void UnregisterAllEventReceivers(this List list, string name)
        {
            ClientContext clientContext = (ClientContext)list.Context;
            EventReceiverDefinitionCollection erdc = list.EventReceivers;
            clientContext.Load(erdc, currentEventReceivers => currentEventReceivers.Include(currentEventReceiver => currentEventReceiver.ReceiverName));
            clientContext.ExecuteQuery();
            List<EventReceiverDefinition> toDelete = new List<EventReceiverDefinition>();
            foreach (EventReceiverDefinition erd in erdc)
            {
                if (erd.ReceiverName == name)
                {
                    toDelete.Add(erd);
                }
            }
            foreach (EventReceiverDefinition item in toDelete)
            {
                item.DeleteObject();
                clientContext.ExecuteQuery();
            }
        }
    }
}
