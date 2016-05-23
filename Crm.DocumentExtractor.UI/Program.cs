using CrmConnection;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace Crm.DocumentExtractor.UI
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Connecting, please wait...");

            var osmSource = new OrganizationServiceManager(ConfigurationManager.AppSettings["source.org.url"]);
            var osmDestination = new OrganizationServiceManager(ConfigurationManager.AppSettings["destination.org.url"]);

            var response = osmSource.GetProxy().Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Console.WriteLine($"Logged into source crm as: {response.UserId}");

            response = osmDestination.GetProxy().Execute(new WhoAmIRequest()) as WhoAmIResponse;
            Console.WriteLine($"Logged into destination crm as: {response.UserId}");

            Console.WriteLine("\r\nAre you sure you want to upload all templates to the destination? Y/N");
            string input = (Console.ReadLine() ?? "").ToLower();

            if (input == "y" || input == "yes")
            {
                Upload(osmSource, osmDestination);
            }
        }

        private static void Upload(OrganizationServiceManager osmSource, OrganizationServiceManager osmDestination)
        {
            var templates = GetTemplates(osmSource);
            Console.WriteLine($"\r\nFound {templates.Count} docx templates to upload");

            Parallel.ForEach(templates,
                new ParallelOptions { MaxDegreeOfParallelism = 5 },
                (template =>
                {
                    string name = template.GetAttributeValue<string>("name");
                    Console.WriteLine($"Uploading {name}");

                    try
                    {
                        string etc = template.GetAttributeValue<string>("associatedentitytypecode");

                        int? oldEtc = GetEntityTypeCode(osmSource, etc);
                        int? newEtc = GetEntityTypeCode(osmDestination, etc);

                        string fileName = ReRouteEtcViaOpenXML(template, name, etc, oldEtc, newEtc);

                        template["associatedentitytypecode"] = newEtc;
                        template["content"] = Convert.ToBase64String(File.ReadAllBytes(fileName));

                        Guid existingId = TemplateExists(osmDestination, name);
                        if (existingId != null && existingId != Guid.Empty)
                        {
                            template["documenttemplateid"] = existingId;

                            osmDestination.GetProxy().Update(template);
                            Console.WriteLine($"Updated {name}");
                        }
                        else
                        {
                            Guid id = osmDestination.GetProxy().Create(template);
                            Console.WriteLine($"Created {name}");
                        }

                        File.Delete(fileName); // delete the updated file but keep the original
                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        Console.WriteLine($"Failed to upload {name}!");
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.Detail != null ? ex.Detail.TraceText : "");
                        Console.WriteLine(ex.InnerException != null ? ex.InnerException.Message : "");
                    }
                }));
        }

        private static string ReRouteEtcViaOpenXML(Entity template, string name, string etc, int? oldEtc, int? newEtc)
        {
            byte[] content = Convert.FromBase64String(template.GetAttributeValue<string>("content"));

            string originalFileName = string.Format(@".\{0}.original.docx", name); // keep a backup just in case we need to debug
            string updatedFileName = string.Format(@".\{0}.updated.docx", name);

            File.WriteAllBytes(originalFileName, content);
            File.WriteAllBytes(updatedFileName, content);

            string toFind = string.Format("{0}/{1}", etc, oldEtc);
            string replaceWith = string.Format("{0}/{1}", etc, newEtc);

            using (var doc = WordprocessingDocument.Open(updatedFileName, true, new OpenSettings { AutoSave = true }))
            {
                // crm keeps the etc in multiple places; parts here are the actual merge fields
                doc.MainDocumentPart.Document.InnerXml = doc.MainDocumentPart.Document.InnerXml.Replace(toFind, replaceWith);
                Console.WriteLine($"Replaced '{toFind}' with '{replaceWith}' inside document.xml on {name}");

                // next is the actual namespace declaration
                doc.MainDocumentPart.CustomXmlParts.ToList().ForEach(a =>
                {
                    using (StreamReader reader = new StreamReader(a.GetStream()))
                    {
                        var xml = XDocument.Load(reader);

                        // crappy way to replace the xml, couldn't be bothered figuring out xml root attribute replacement...
                        var crappy = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" + xml.ToString();

                        if (crappy.IndexOf(toFind) > -1) // only replace what is needed
                        {
                            crappy = crappy.Replace(toFind, replaceWith);
                            Console.WriteLine($"Replaced '{toFind}' with '{replaceWith}' inside \\customXml\\*.xml {name}");

                            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(crappy)))
                            {
                                a.FeedData(stream);
                            }
                        }
                    }
                });
            }

            return updatedFileName;
        }

        private static Guid TemplateExists(OrganizationServiceManager osm, string name)
        {
            Guid result = Guid.Empty;

            QueryExpression qe = new QueryExpression("documenttemplate");
            qe.Criteria.AddCondition("status", ConditionOperator.Equal, false); // only get active templates
            qe.Criteria.AddCondition("name", ConditionOperator.Equal, name);

            var results = osm.GetProxy().RetrieveMultiple(qe);
            if (results != null && results.Entities != null && results.Entities.Count > 0)
            {
                result = results[0].Id;
            }

            return result;
        }

        public static int? GetEntityTypeCode(OrganizationServiceManager osm, string entity)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest();

            request.LogicalName = entity;
            request.EntityFilters = EntityFilters.Entity;

            RetrieveEntityResponse response = (RetrieveEntityResponse)osm.GetProxy().Execute(request);
            EntityMetadata metadata = response.EntityMetadata;

            return metadata.ObjectTypeCode;
        }

        private static List<Entity> GetTemplates(OrganizationServiceManager osm)
        {
            QueryExpression qe = new QueryExpression("documenttemplate") { ColumnSet = new ColumnSet("content", "name", "associatedentitytypecode", "documenttype", "clientdata") };
            qe.Criteria.AddCondition("status", ConditionOperator.Equal, false); // only get active templates
            qe.Criteria.AddCondition("documenttype", ConditionOperator.Equal, 2); // only word docs
            qe.Criteria.AddCondition("createdbyname", ConditionOperator.NotEqual, "SYSTEM");

            var results = osm.GetProxy().RetrieveMultiple(qe);
            if (results != null && results.Entities != null && results.Entities.Count > 0)
            {
                return results.Entities.ToList();
            }

            return new List<Entity>();
        }
    }
}
