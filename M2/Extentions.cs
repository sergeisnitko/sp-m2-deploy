using System;
using System.Linq;
using System.Text;

using SPMeta2.Models;
using SPMeta2.Syntax.Default;
using SPMeta2.Syntax.Default.Modern;

using SPMeta2.Definitions;
using SPMeta2.Definitions.Fields;

using Microsoft.SharePoint.Client;

using SPMeta2.Definitions.Webparts;
using System.IO;
using SPF.Extentions;
using SPMeta2.CSOM.Services;
using SPMeta2.Services;
using SPMeta2.Extensions;
using SPMeta2.Common;
using SPMeta2.Definitions.ContentTypes;
using System.Collections.Generic;

namespace SPF.M2
{ 
    public static class Extentions
        /////
    {
        public static void UpdateLookupField(this Web GWeb, List GList, LookupFieldDefinition LookupDefinition)
        {
            GWeb.UpdateLookupField(GList, LookupDefinition.InternalName);
        }

        public static void UpdateLookupField(this Web GWeb, ListDefinition SpfListDefinition, LookupFieldDefinition LookupDefinition)
        {
            var SpfList = GWeb.GetListByUrl(SpfListDefinition.CustomUrl);
            GWeb.Context.Load(SpfList);
            GWeb.UpdateLookupField(SpfList, LookupDefinition);
        }

        public static void UpdateDependedLookupField(this Web GWeb, ListDefinition ListUsageDefinition, DependentLookupFieldDefinition LookupDefinition)
        {
            var Ctx = GWeb.Context;
            var SpfList = GWeb.GetListByUrl(ListUsageDefinition.CustomUrl);

            var tempFild = GWeb.Fields.GetById((Guid)LookupDefinition.PrimaryLookupFieldId);
            Ctx.Load(tempFild);
            Ctx.ExecuteQuery();

            var PrimaryLookup = Ctx.CastTo<FieldLookup>(tempFild);
            //var PrimaryLookup = (S)tempFild;
            var Lookup = GWeb.Fields.GetByInternalNameOrTitle(LookupDefinition.InternalName);
            Ctx.Load(Lookup);
            Ctx.Load(PrimaryLookup);
            Ctx.Load(SpfList);
            Ctx.Load(SpfList.Fields);
            Ctx.Load(SpfList.ParentWeb);
            Ctx.ExecuteQuery();

            var InDependedLookup = SpfList.Fields
                                        .Cast<Field>()
                                        .FirstOrDefault(c => c.InternalName == LookupDefinition.InternalName);

            if (InDependedLookup == null)
            {
                var LookupSchema = string.Format(@"<Field Type=""Lookup""
	            DisplayName=""{0}""
	            List=""{1}""
	            ShowField=""{2}""
	            FieldRef=""{3}""
	            ReadOnly=""TRUE""
	            UnlimitedLengthInDocumentLibrary=""FALSE""
	            ID=""{4}""
	            Name=""{1}"" ></Field>", LookupDefinition.InternalName, PrimaryLookup.LookupList, LookupDefinition.LookupField, LookupDefinition.PrimaryLookupFieldId, LookupDefinition.Id.ToString());
                LookupSchema = LookupSchema.Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace("  ", " ").Trim();

                SpfList.Fields.AddFieldAsXml(LookupSchema.Replace("\n\r", ""), true, AddFieldOptions.DefaultValue);
                Ctx.ExecuteQuery();
            }


            InDependedLookup = SpfList.Fields.GetByInternalNameOrTitle(LookupDefinition.InternalName);
            Ctx.Load(InDependedLookup);
            Ctx.ExecuteQuery();

            InDependedLookup.Title = LookupDefinition.Title;
            InDependedLookup.SetShowInDisplayForm((bool)LookupDefinition.ShowInDisplayForm);
            InDependedLookup.SetShowInEditForm((bool)LookupDefinition.ShowInEditForm);
            InDependedLookup.SetShowInNewForm((bool)LookupDefinition.ShowInNewForm);
            InDependedLookup.UpdateAndPushChanges(true);
            Ctx.ExecuteQuery();

            InDependedLookup.SchemaXml = InDependedLookup.SchemaXml.ReplaceXmlAttributeValue("List", PrimaryLookup.LookupList).ReplaceXmlAttributeValue("WebId", PrimaryLookup.LookupWebId.ToString());
            Ctx.ExecuteQuery();
        }


        public static TModelNode AddCustomFolder<TModelNode>(this TModelNode node, FolderDefinition FolderDef, ContentTypeDefinition ContentType, int Order)
                    where TModelNode : ModelNode, IFolderHostModelNode, new()
        {
            return AddCustomFolder(node, FolderDef, ContentType, Order, null, null);
        }
        public static TModelNode AddCustomFolder<TModelNode>(this TModelNode node, FolderDefinition FolderDef, ContentTypeDefinition ContentType, int Order, Action<FolderModelNode> FolderAction)
            where TModelNode : ModelNode, IFolderHostModelNode, new()
        {
            return AddCustomFolder(node, FolderDef, ContentType, Order, FolderAction, null);
        }

        public static TModelNode AddCustomFolder<TModelNode>(this TModelNode node, FolderDefinition FolderDef, ContentTypeDefinition ContentType, int Order, Action<FolderModelNode> FolderAction, Action<FolderModelNode> SecurityAction)
            where TModelNode : ModelNode, IFolderHostModelNode, new()
        {
            node
                .AddFolder(FolderDef, prj =>
                {
                    if (FolderAction != null)
                    {
                        var modelNode = new FolderModelNode { Value = FolderDef };
                        node.ChildModels.Add(modelNode);
                        FolderAction(modelNode);
                    }

                    prj.AddBreakRoleInheritance(new BreakRoleInheritanceDefinition
                    {
                        CopyRoleAssignments = false,
                        ClearSubscopes = true
                    }, dlWithBrokenInheritance =>
                    {
                        if (SecurityAction != null)
                        {
                            SecurityAction(dlWithBrokenInheritance);
                        }
                    });


                    prj.OnProvisioned<object>(context =>
                    {
                        ChangeFolderContentType(context, Order, ContentType);
                    });
                });

            return node;
        }
        public static void ChangeFolderContentType(OnCreatingContext<object, DefinitionBase> context, int Order, ContentTypeDefinition ContentType)
        {
            var obj = context.Object;
            var objType = context.Object.GetType();

            if (objType.ToString().Contains("Microsoft.SharePoint.Client.Folder"))
            {
                var Folder = (Folder)obj;
                var ctx = Folder.Context;

                var FolderItem = Folder.ListItemAllFields;
                ctx.Load(FolderItem);
                ctx.ExecuteQuery();
                if (FolderItem != null)
                {
                    FolderItem["ContentTypeId"] = ContentType.GetContentTypeId().ToString();
                    FolderItem.Update();
                    ctx.ExecuteQuery();
                }
            }


        }

        public static ListViewModelNode AddXsltWebPart(this ListViewModelNode node, XsltListViewWebPartDefinition WebPart, string Query)
        {
            node
                .AddWebPart(WebPart, WP =>
                {
                    WP.OnProvisioned<object>(context =>
                    {
                        ChangeQuery(context, WebPart, Query);
                    });
                });

            return node;
        }

        public static void ChangeQuery(OnCreatingContext<object, DefinitionBase> context, XsltListViewWebPartDefinition Definition, string Query)
        {
            var obj = context.Object;
            var objType = context.Object.GetType();
            if (objType.ToString().Contains("Microsoft.SharePoint.Client.WebParts.WebPart"))
            {

            }
        }


        public static void RemoveElementCTFromList(this Web SpfWeb, ListDefinition SpfListDefinition)
        {
            var SpfList = SpfWeb.GetListByUrl(SpfListDefinition.CustomUrl);
            SpfWeb.Context.Load(SpfList);
            SpfList.RemoveContentType("Элемент");
            SpfList.RemoveContentType("Item");
        }

        public static string GetFileText(string Path)
        {
            var text = "";
            if (System.IO.File.Exists(Path))
            {
                using (TextReader tw = new StreamReader(Path, Encoding.UTF8))
                {
                    text = tw.ReadToEnd();
                    tw.Close();
                }
            }
            return text;
        }


        public static void GenerateJavascriptFile(string Path, string[] JavascriptRows)
        {
            if (System.IO.File.Exists(Path))
            {
                System.IO.File.Delete(Path);
            }
            System.IO.File.Create(Path).Dispose();
            using (TextWriter tw = new StreamWriter(Path, true, Encoding.UTF8))
            {
                var Builder = new StringBuilder();
                Builder.AppendLine("\t" + string.Join("\n\t", JavascriptRows));
                tw.WriteLine(Builder.ToString());
                tw.Close();
            }
        }

        public static string AddZeros(this object Number, int Zeros)
        {
            return Number.AddBeforeSymbols(Zeros, '0');
        }

        public static string AddSpacesBefore(this object Number, int Zeros)
        {
            return Number.AddBeforeSymbols(Zeros, ' ');
        }

        public static string AddBeforeSymbols(this object Str, int Zeros, char Symbol)
        {
            var val = Str.ToString();
            for (var i = 0; i < Zeros; i += 1)
            {
                val = Symbol + val;
            }
            val = val.Substring(val.Length - Zeros);

            return val;
        }

        public static ListModelNode AddRemoveStandardContentTypes(this ListModelNode node)
        {
            node
                .AddRemoveContentTypeLinks(new RemoveContentTypeLinksDefinition
                {
                    ContentTypes = new List<ContentTypeLinkValue>
                    {
                        new ContentTypeLinkValue{ ContentTypeName = "Элемент" },
                        new ContentTypeLinkValue{ ContentTypeName = "Item" }
                    }
                })

                ;
            return node;
        }


        public static void DeployModel(this ClientContext Ctx, WebModelNode model)
        {
            DeployModel(Ctx, model, true);
        }

        public static void DeployModel(this ClientContext Ctx, SiteModelNode model)
        {
            DeployModel(Ctx, model, true);
        }

        public static void DeployModel(this ClientContext Ctx, WebModelNode model, bool Incremental)
        {
            BeforeDeployModel(Incremental, x =>
            {
                x.DeployModel(SPMeta2.CSOM.ModelHosts.WebModelHost.FromClientContext(Ctx), model);
            });
        }
        public static void DeployModel(this ClientContext Ctx, SiteModelNode model, bool Incremental)
        {
            BeforeDeployModel(Incremental, x =>
            {
                x.DeployModel(SPMeta2.CSOM.ModelHosts.SiteModelHost.FromClientContext(Ctx), model);
            });

        }
        public static void BeforeDeployModel(bool Incremental, Action<CSOMProvisionService> Deploy)
        {
            var StartedDate = DateTime.Now;
            var provisionService = new CSOMProvisionService();
            if (Incremental)
            {
                var IncProvisionConfig = new IncrementalProvisionConfig();
                IncProvisionConfig.AutoDetectSharePointPersistenceStorage = true;
                provisionService.SetIncrementalProvisionMode(IncProvisionConfig);
            }
            provisionService.OnModelNodeProcessed += (sender, args) =>
            {
                ModelNodeProcessed(sender, args, Incremental);
            };

            Deploy(provisionService);
            provisionService.SetDefaultProvisionMode();
            var FinishedDate = DateTime.Now;
            var DateDiff = (FinishedDate - StartedDate);
            var TotalHrs = Math.Round(DateDiff.TotalHours);
            var TotalMinutes = Math.Round(DateDiff.TotalMinutes);
            var TotalSeconds = Math.Round(DateDiff.TotalSeconds);

            if (TotalHrs == 0)
            {
                if (TotalMinutes == 0)
                {
                    Console.WriteLine(String.Format("It took us {0} seconds", TotalSeconds.ToString()));
                }
                else
                {
                    Console.WriteLine(String.Format("It took us {0} minutes", TotalMinutes.ToString()));
                }
            }
            else
            {
                Console.WriteLine(String.Format("It took us {0} hours", TotalHrs.ToString()));
            }
            Console.WriteLine();
            Console.WriteLine();

        }

        public static void ModelNodeProcessed(object sender, ModelProcessingEventArgs args, bool Incremental)
        {
            var ModelId = args.Model.GetPropertyBagValue(DefaultModelNodePropertyBagValue.Sys.IncrementalProvision.PersistenceStorageModelId);

            bool shouldDeploy = args.CurrentNode.GetIncrementalRequireSelfProcessingValue();

            var NodeName = args.CurrentNode.Value.ToString();
            if (NodeName.Length > 20)
            {
                NodeName = NodeName.Substring(0, 20) + "...";
            }
            if (!Incremental)
            {
                shouldDeploy = true;
            }

            Console.WriteLine(
            string.Format("{5}[{6}] [{0}/{1}] - [{2}%] - [{3}] [{4}]",
            new object[] {
                    args.ProcessedModelNodeCount.AddZeros(4),
                    args.TotalModelNodeCount.AddZeros(4),
                    Math.Round(100d * (double)args.ProcessedModelNodeCount / (double)args.TotalModelNodeCount).AddSpacesBefore(3),
                    args.CurrentNode.Value.GetType().Name,
                    NodeName,
                    (shouldDeploy == true) ? "[+]" : "[-]",
                    ModelId
           }));

        }

    }
}
