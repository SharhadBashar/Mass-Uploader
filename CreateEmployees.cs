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

namespace CreateEmployees
{
    class Program
    {
        public static Random random = new Random();
        static void Main(string[] args)   
        {
            using (SPSite qmwSite = new SPSite(args[0]))
            {
                using (SPWeb qmwWeb = qmwSite.RootWeb)
                {
                    var employees = GetLookupValues(qmwWeb, "Employee List", "ID", "");
                    int count = Convert.ToInt32(((employees.Last()).ToString()).Split(';')[0]);  
                    //Console.WriteLine("Creating Employees");
                    for (int i = 0; i < Convert.ToInt32(args[1]); i++)
                    {
                        //Console.WriteLine((Convert.ToInt32(args[1]) - 1) * 50 + i);
                        //int num = ((Convert.ToInt32(args[2]) - 1) * 50 + i);
                        if (i % 100 == 0)
                        {
                            Console.WriteLine(i);
                            Thread.Sleep(1000);
                        }
                        generateEmployee(qmwWeb, (count + i + 1));
                        
                    }
                }
            }
            //Console.ReadLine();
        }
/********************************************************************************************************************************************/
        public static String nameGenerator()
        {
            int minLength = 5, maxLength = 10;
            char[] consonants = { 'b', 'c', 'd', 'f', 'g', 'h', 'j', 'k', 'l', 'm', 'n', 'p', 'q', 'r', 's', 't', 'v', 'w', 'x', 'z' };
            char[] vowels = { 'a', 'e', 'i', 'o', 'u', 'y' };
            var length = random.Next(minLength, maxLength + 1);
            var name = new char[length];
            for (int i = 0; i < length; i++)
            {
                var index = random.Next(1, 3);
                if (index == 1)
                {
                    name[i] = vowels[random.Next(0, vowels.Length)];
                }
                else if (index == 2)
                {
                    name[i] = consonants[random.Next(0, vowels.Length)];
                }
            }
            name[0] = Char.ToUpper(name[0]);
            return new String(name);
        }// End nameGenerator
/********************************************************************************************************************************************/
        public static int numberOfItem(SPWeb web, String listName)
        {
            int count = web.Lists[listName].Items.Count;
            return count;
        }//End numofitem
/********************************************************************************************************************************************/
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
        }//End Getlookupvalues
/********************************************************************************************************************************************/
        public static void addListItem(SPList list, Dictionary<string, object> itemProperties, string url = null)
        {
            var item = url == null ? list.AddItem() : list.AddItem(url, SPFileSystemObjectType.File);
            foreach (var property in itemProperties) item[property.Key] = property.Value;
            item.Update();
        }//End additems
/********************************************************************************************************************************************/
        public static SPListItem generateEmployee(SPWeb web, int count)
        {
            SPList list = web.Lists["Employee List"];
            var jobTitles = GetLookupValues(web, "Job Titles", "Job_x0020_Title", ""); ;
            var bus = GetLookupValues(web, "Business Units", "Business_x0020_Unit", "");
            var bu = bus.GetRandom();
            String[] buString = (bu.ToString()).Split('#');
            var depts = GetLookupValues(web, "Departments", "Department_x0020_Name", buString[1]);
            var subDepts = GetLookupValues(web, "Sub-Departments", "Sub_x002d_Department", buString[1]);
            var locations = GetLookupValues(web, "Locations", "Title", buString[1]);
            var firstName = nameGenerator();
            var lastName = nameGenerator();
            
            var item = new Dictionary<string, object>
                {
                    { "Employee ID", ("EID-" + (count + 1)) },
                    { "First Name", firstName },
                    { "Last Name", lastName },
                    {"E-Mail", String.Format("{0}.{1}@fake.com",firstName.ToLower(), lastName.ToLower()) },
                    { "Job Title", jobTitles.GetRandom() },
                    { "Business Unit", bu },
                    { "Department", depts.GetRandom() },
                    { "Sub-Department", subDepts.GetRandom() },
                    { "Location", locations.GetRandom() },
                    { "Employee Status", "Not Active" },  
                };
            addListItem(list, item);
            return null;
        }
/********************************************************************************************************************************************/
    }
}
