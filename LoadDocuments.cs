using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Collections;
using System.Data;
using System.Net;
using System.IO;
using System.Runtime.Serialization.Json;
using Microsoft.SharePoint.Utilities;

namespace LoadEffective{
    /// <summary>
    /// This code is for mass uploading word documents Effective and Draft Controlled Documents Libraries and DAL forms to the DAL library
    /// </summary>
    public class Program
    {
        /// <summary>
        /// The main method that gets the user input on how many documents to add in each library, and then calls the method to add the documents to the libraries 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {     
            string QMWsite = args[0];
            
            //Get inputs for how many items to add in each of the libraries
            Console.Write("How many files in Effective Controlled Documents?\n");
            int ecd = Convert.ToInt32(Console.ReadLine());
            
            Console.Write("How many files in Draft Controlled Documents?\n");
            int dcd = Convert.ToInt32(Console.ReadLine());
            
            Console.Write("How many files in Document Activity Library?\n");
            int dal = Convert.ToInt32(Console.ReadLine());
            Console.Write("\n");
            string location = args[1] + "\\";
            BatchXml batchXml = new BatchXml();
            //Get the site where QMW is located
            using (SPSite site = new SPSite(QMWsite))
            {
                using (SPWeb root = site.RootWeb)//get the root site
                {
                    //where the libraries are located
                    using (SPWeb web = site.OpenWeb("CDM"))
                    {
                        for (int libCount = 0; libCount < 3; libCount++)
                        {
                            //add documents to each of the libraries in series
                            if (libCount == 0)
                            {
                                AddRandomDraftDocuments(root, web, ecd, location, libCount);
                            }
                            else if (libCount == 1)
                            {
                                AddRandomDraftDocuments(root, web, dcd, location, libCount);
                            }
                            else
                            {
                                AddRandomDraftDocuments(root, web, dal, location, libCount);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the id of the last item in each library
        /// </summary>
        /// <param name="list">Library/List name in SharePoint</param>
        /// <returns>Last ID of the item</returns>
        static int GetNextID(SPList list)
        {
            SPListItemCollection items = list.GetItems(
                new SPQuery
                {
                    Query = Q.OrderBy(false, "ID"),
                    ViewFields = "<FieldRef Name=\"ID\"/>",
                    ViewAttributes = Q.ScopeRecursive
                }
                );
            return items[0].ID + 1;
        }

        /// <summary>
        /// Adds the document to the library with the specified metadata
        /// </summary>
        /// <param name="folder">Sub folder in Library</param>
        /// <param name="url">URL of document to be added</param>
        /// <param name="bytes">The document</param>
        /// <param name="properties">Document Metadata</param>
        static void AddDocument(SPFolder folder, string url, byte[] bytes, Hashtable properties)
        {
            SPFile file = folder.Files.Add(url, bytes, properties, true);
            file.Update();
        }

        /// <summary>
        /// Gets the document from the local directory and assigns metadata to it, and calls the method to add the document
        /// </summary>
        /// <param name="root">Root site</param>
        /// <param name="web">Module where libraries are located</param>
        /// <param name="count">How many items to be added</param>
        /// <param name="listName">Library where documents are to be added</param>
        static void AddRandomDraftDocuments(SPWeb root, SPWeb web, int count, string location, int listName)
        {
            SPList list;
            byte[] bytes;
            string url;

            if (listName == 0)   
            {
                Console.Write("Effective Controlled Documents\n");
                list = web.Lists["Effective Controlled Documents"];
                bytes = File.ReadAllBytes(@location + "TestDoc.docx");
            }
            else if (listName == 1)
            {
                Console.Write("Draft Controlled Documents\n");
                list = web.Lists["Draft Controlled Documents"];
                bytes = File.ReadAllBytes(@location + "TestDoc.docx");
            }
            else 
            {
                Console.Write("Document Activity Logs\n");
                list = web.Lists["Document Activity Logs"];
                bytes = File.ReadAllBytes(@location + "DAL-00000.xml");
            }
            int nextIndex = GetNextID(list); //gets id of the last item
            var buDic = GetBUDeptDict(root); //gets bu and dept from root site
            var folders = list.RootFolder.SubFolders.Cast<SPFolder>().ToList(); //gets the subfolders in each library
            List<string> contentTypeIds = list.ContentTypes.Cast<SPContentType>().Where(t => t.Name != "Document").Select(c => c.Id.ToString()).ToList(); //gets content types
            for (int i = 0; i < count; i++)
            {
                Console.WriteLine(i);
                if (listName == 2)
                {
                    url = String.Format("DAL-00000{0}.xml", i + nextIndex); //creates the url based on doc type (dal form) and id
                }
                else
                {
                    url = String.Format("Test {0}.doc", i + nextIndex); //creates the url based on doc type (word doc) and id
                }
                var buPair = buDic.GetRandom();
                //sets the meta data for the doc
                Hashtable properties = new Hashtable
                {
                    { "Name", String.Format("Test {0}", i + nextIndex) },
                    { "ContentTypeId", contentTypeIds.GetRandom() },
                    { "Title", String.Format("Test {0}", i + nextIndex) },
                    { "_Revision", "00" },
                    { "Business_x0020_Unit", buPair.Key.LookupId },
                };
                if (listName == 2)
                {
                    //special method call since DAL library has no subfolder
                    AddDocument(folders.Count == 1 ? list.RootFolder : folders.GetRandom(), url, bytes, properties);
                }
                else
                {
                    //method call to add docuemnt
                    SPFolder folder = null;
                    if (folders.Count == 1)
                    {
                        folder = list.RootFolder;
                    }
                    else
                    {
                        folder = folders.GetRandom();
                        
                        while (folder.Name == "Forms")
                        {
                            folder = folders.GetRandom();
                        }
                    }
                    AddDocument(folder, url, bytes, properties);
                }

            }
        }

        /// <summary>
        /// Gets a random BU and all Dept associated with that BU
        /// </summary>
        class buDeptPair
        {
            public SPFieldLookupValue Key { get; set; }
            public List<SPFieldLookupValue> Value { get; set; }

        }

        /// <summary>
        /// Creates a list of BU and the corresponding Dept values
        /// </summary>
        /// <param name="web">Site where QMW is deployed</param>
        /// <returns>The list</returns>
        static List<buDeptPair> GetBUDeptDict(SPWeb web)
        {
            return GetDict(web, "Business Units", "Business Unit", "Departments", "Business Unit", "Department Name");
        }

        /// <summary>
        /// Creates a list based on the parameters provided. Goes into the site grabs all the items in both the lists. 
        /// It then grabs all the items in the first list based on the field name and puts them in a list
        /// It then grabs all the items in the second list, and filters based on the look up value and adds the items in the list 
        /// </summary>
        /// <param name="web">Site where QMW is deployes</param>
        /// <param name="list1">List one</param>
        /// <param name="field1">Field name in List 1</param>
        /// <param name="list2">List 2</param>
        /// <param name="lookupField">Look up Field in List 2</param>
        /// <param name="field2">Filed name in List 2</param>
        /// <returns>The parent child list</returns>
        static List<buDeptPair> GetDict(SPWeb web, string list1, string field1, string list2, string lookupField, string field2)
        {
            SPListItemCollection bus = web.Lists[list1].Items;
            SPListItemCollection es = web.Lists[list2].Items;
            List<buDeptPair> dic = new List<buDeptPair>();
            foreach (SPListItem item in bus)
            {
                string bu = item[field1] as string;
                if (!dic.Any(v => v.Key.LookupValue == bu))
                {
                    dic.Add(new buDeptPair { Key = new SPFieldLookupValue(item.ID, bu), Value = new List<SPFieldLookupValue>() });
                }
            }
            foreach (SPListItem item in es)
            {
                string e = item[field2] as string;
                try
                {
                    SPFieldLookupValue bu = new SPFieldLookupValue(item[lookupField].ToString());
                    dic.First(p => p.Key.LookupValue == bu.LookupValue).Value.Add(new SPFieldLookupValue(item.ID, e));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            return dic;
        }
    }
}
