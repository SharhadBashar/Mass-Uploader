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

namespace MassUploader
{
    class Program
    {
        //args[0] = QMW site
        //args[1] = # of Job Roles
        //args[2] = # of Job Titles
        //args[3] = # of Courses
        public static Random random = new Random();
        static void Main(string[] args)
        {
            
            /*Console.WriteLine("Please enter the number of Job Roles: ");
            int jobRoles = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Please enter the number of Job Titles: ");
            int jobTitles = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Please enter the number of Courses: ");
            int courses = Convert.ToInt32(Console.ReadLine());*/

            using (SPSite qmwSite = new SPSite(args[0]))
            {
                using (SPWeb qmwWeb = qmwSite.RootWeb)
                {
                    Console.WriteLine("Creating Job Roles");
                    for (int i = 0; i < Convert.ToInt32(args[1]); i++)
                    {
                        if (i % 100 == 0)
                        {
                            Console.WriteLine(i);
                        }
                        createJobRoles(qmwWeb);
                    }
                    Console.WriteLine("Done Job Roles");

                    Console.WriteLine("Creating Job Titles");
                    for (int i = 0; i < Convert.ToInt32(args[2]); i++)
                    {
                        if (i % 100 == 0)
                        {
                            Console.WriteLine(i);
                        }
                        createJobTitles(qmwWeb);  
                    }
                    Console.WriteLine("Done Job Titles");
                    using (SPWeb tmmWeb = qmwSite.OpenWeb("TMM"))
                    {
                        var course = GetLookupValues(tmmWeb, "Training Matrix", "ID", "");
                        int count = Convert.ToInt32(((course.Last()).ToString()).Split(';')[0]);
                        Console.WriteLine("Creating Courses");
                        //int count = numberOfItem(tmmWeb, "Training Matrix");
                        for (int i = 0; i < Convert.ToInt32(args[3]); i++)
                        {
                            if (i % 100 == 0)
                            {
                                Console.WriteLine(i);
                            }
                            createCourses(qmwWeb, tmmWeb, (count + i));
                        }
                        Console.WriteLine("Done Courses");
                    }
                }
            }
            //Console.ReadKey();
        }//End Main
/******************************************************************************************************************************************/
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
/******************************************************************************************************************************************/
        public static int numberOfItem(SPWeb web, String listName)
        {
            int count = web.Lists[listName].Items.Count;
            return count;
        }//End numofitem
/******************************************************************************************************************************************/
        public static int lastID(SPWeb web, String listName)
        {
            int ID = 0;
            int count = web.Lists[listName].Items.Count;
            SPList list = web.Lists[listName];
            SPQuery query = new SPQuery();
            Console.WriteLine("Starting Query");
            query.Query = String.Format("<OrderBy><FieldRef Name=\"ID\" Asc=\"TRUE\"/></OrderBy> ");
            query.ViewFields = "<FieldRef Name =\"ID\"/>"; //<FieldRef Name =\"Employee_x0020_ID\"/><FieldRef Name =\"First_x0020_Name\"/>
            
            SPListItemCollection itemCollection = list.GetItems(query);
            Console.WriteLine(itemCollection[count - 1]["ID"]);
            return ID;
        }
/******************************************************************************************************************************************/
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
/******************************************************************************************************************************************/
        public static void addListItem(SPList list, Dictionary<string, object> itemProperties, string url = null)
        {
            var item = url == null ? list.AddItem() : list.AddItem(url, SPFileSystemObjectType.File);
            foreach (var property in itemProperties) item[property.Key] = property.Value;
            item.Update();
        }//End additems
/******************************************************************************************************************************************/
        public static SPListItem createJobRoles(SPWeb web)
        {
            SPList list = web.Lists["Job Roles"];
            var bu = GetLookupValues(web, "Business Units", "Business_x0020_Unit", "");
            var item = new Dictionary<string, object>
            {
                { "Job Role", String.Concat(nameGenerator(), " ", nameGenerator()) },
                { "Business Unit", bu.GetRandom() },
            };
            addListItem(list, item);
            return null;
        }//End jobRoles
/******************************************************************************************************************************************/
        public static SPListItem createJobTitles(SPWeb web)
        {
            int maxTitles = random.Next(5,11);
            //if (count < 100) { maxTitles = 140; }
            SPList list = web.Lists["Job Titles"];
            var jobRoles = GetLookupValues(web, "Job Roles", "Job_x0020_Role", "");
            var bu = GetLookupValues(web, "Business Units", "Business_x0020_Unit", "");
            SPFieldLookupValueCollection jobRoleList = new SPFieldLookupValueCollection { };
            for (int i = 0; i < maxTitles; i++)
            {
                jobRoleList.Add(jobRoles.GetRandom());
            }
            var item = new Dictionary<string, object>
            {
                { "Job Title", String.Concat(nameGenerator(), " ", nameGenerator()) },
                { "Business Unit", bu.GetRandom() },
                { "Job Roles",  jobRoleList},
            };
            addListItem(list, item);
            return null;
        }//End jobTitles
/******************************************************************************************************************************************/
        public static SPListItem createCourses(SPWeb qmwWeb, SPWeb web, int count)
        {
            int maxTitles = random.Next(5, 11);
            SPList list = web.Lists["Training Matrix"];
            var bus = GetLookupValues(qmwWeb, "Business Units", "Business_x0020_Unit", "");
            var bu = bus.GetRandom();
            String[] buString = (bu.ToString()).Split('#');
            var depts = GetLookupValues(qmwWeb, "Departments", "Department_x0020_Name", buString[1]);
            var jobRoles = GetLookupValues(qmwWeb, "Job Roles", "Job_x0020_Role", "");
            SPFieldLookupValueCollection jobRoleList = new SPFieldLookupValueCollection { };
            for (int i = 0; i < maxTitles; i++)
            {
                jobRoleList.Add(jobRoles.GetRandom());
            }

            var item = new Dictionary<string, object>()
            {
                { "Course ID", ("ID-" + (count + 1)) },
                { "Course Name", String.Concat("Course-", nameGenerator() )},
                { "Course Description", nameGenerator() },
                { "Owning Business Unit", bu },
                { "Owning Department", depts.GetRandom() },
                { "Required Job Roles", jobRoleList },
                { "Course Status", "Available" },
                { "Course Type", "Read and acknowledge" },
                { "Course Level", "Beginner" },
                { "Autonomous Training", "Yes" },
            };
            addListItem(list, item);
            return null;
        }//End courses
 /******************************************************************************************************************************************/
    }
}
