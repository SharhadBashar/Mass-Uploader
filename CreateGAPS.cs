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
using System.Threading;

using Microsoft.SharePoint.Utilities;

namespace GapsToActivities
{
    class Program
    {
        static void Main(string[] args)
        {
            //string qmwSite = args[0];
            //int count = args[1];
            using (SPSite qmwSite = new SPSite(args[0]))
            {
                using (SPWeb qmwWeb = qmwSite.RootWeb)
                {
                    using (SPWeb tmmWeb = qmwSite.OpenWeb("TMM"))
                    {
                        List<String[]> gaps = getItemsfromGaps(tmmWeb, Convert.ToInt32(args[1]));
                        
                        for (int i = 0; i < gaps.Count; i++)
                        {
                            string[] value = new string[3];
                            value = gaps[i];
                            //Console.WriteLine(value[1] + " " + value[2]);
                            activities(tmmWeb, i + Convert.ToInt32(args[2]), value);
                        }
                    }
                }
            }
        }//End Main
/****************************************************************************************************************************************/
        public static int numberOfItem(SPWeb web, String listName)
        {
            int count = web.Lists[listName].Items.Count;
            return count;
        }//End numofitem
/****************************************************************************************************************************************/
        public static List<SPFieldLookupValue> GetLookupValues(SPWeb web, string listName, string lookupField, string lookupValue)
        {
            try
            {
                if (lookupValue == "")
                {
                    return web.Lists[listName].GetItems(
                    new SPQuery
                    {
                        Query = Q.Where(Q.IsNotNull(lookupField)),
                        ViewFields = Q.ViewFields(lookupField, "ID"),
                    }).Cast<SPListItem>().Select(i => new SPFieldLookupValue(i.ID, i[lookupField] as string)).ToList();
                }
                else
                {
                    return web.Lists[listName].GetItems(
                        new SPQuery
                        {
                            Query = Q.WhereAnd(Q.IsNotNull(lookupField), Q.Condition("Business_x0020_Unit", lookupValue)),
                            ViewFields = Q.ViewFields(lookupField, "ID"),
                        }).Cast<SPListItem>().Select(i => new SPFieldLookupValue(i.ID, i[lookupField] as string)).ToList();
                }
            }
            catch
            {
                return null;
            }
        }//End LookupValues
/****************************************************************************************************************************************/      
        public static void addListItem(SPList list, Dictionary<string, object> itemProperties, string url = null)
        {
            var item = url == null ? list.AddItem() : list.AddItem(url, SPFileSystemObjectType.File);
            foreach (var property in itemProperties) item[property.Key] = property.Value;
            item.Update();
        }//End additems
/****************************************************************************************************************************************/
        public static List<string[]> getItemsfromGaps(SPWeb web, int amount)
        {
            var cid = GetLookupValues(web, "Training Gaps", "Title", "");
            var eid = GetLookupValues(web, "Training Gaps", "Employee_x0020_ID", "");
            List<String[]> gaps = new List<string[]>();

            //int numOfGaps = numberOfItem(web, "Training Gaps");
            for (int i = 0; i < amount; i++)
            {
                String[] cidString = (cid[i].ToString()).Split('#');
                String[] eidString = (eid[i].ToString()).Split('#');
                gaps.Add(new String[] { "" + i, cidString[1], eidString[1] });
            }
            return gaps;
        }
/****************************************************************************************************************************************/
        public static SPListItem activities(SPWeb web, int i, String[] value)
        {
            SPList list = web.Lists["Training Activities"];
            String id = "TR-" + i;

            var item = new Dictionary<string, object>
                {
                    { "Activity ID", id },
                    { "Course ID", value[1]},
                    { "Autonomous Training", "Yes" },
                    { "Trainee ID", value[2]},
                };
            Console.WriteLine("Adding" + i);
            addListItem(list, item, String.Format("{0}/Lists/{1}/{2}", web.ServerRelativeUrl, list.Title, value[2]));
            return null;
        }
    }
}
