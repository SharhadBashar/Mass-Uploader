using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.WebPartPages;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using IOFile = System.IO.File;
using Microsoft.SharePoint.WebPartPages;


// libraries for compression and extraction of Microsoft CAB/XSN files
using System.IO;
using System.IO.Compression;
using Microsoft.SharePoint.Utilities;
using System.Globalization;
namespace LoadEffective
{
    class DisabledItemEventsScope : SPItemEventReceiver, IDisposable
    {
        bool enabledStatus;
        public DisabledItemEventsScope()
        {
            this.enabledStatus = base.EventFiringEnabled;
            base.EventFiringEnabled = false;
        }

        #region IDisposable Members
        public void Dispose()
        {
            base.EventFiringEnabled = enabledStatus;
        }
        #endregion
    }

    public static class DataRowExt
    {
        //gets the datarow value without throwing an error
        public static string GetValueSafe(this DataRow row, string columnName)
        {
            try
            {
                return row[columnName].ToString();
            }
            catch
            {
                return null;
            }
        }

        static Random randy = new Random();
        public static T GetRandom<T>(this IList<T> list)
        {
            return list != null && list.Count > 0 ? list[randy.Next(list.Count)] : default(T);
        }
    }

    public class SqlReaderMap : Dictionary<string, int>
    {

        public SqlReaderMap()
        {
            alternateNameMap = new Dictionary<string, string>();
        }

        public SqlReaderMap(IEnumerable<string> fieldNames)
            : this()
        {
            foreach (string fieldName in fieldNames) { AddColumn(fieldName); }
        }

        public SqlReaderMap(Dictionary<string, string> fieldNameDic)
            : this()
        {
            foreach (var pair in fieldNameDic) { AddColumn(pair.Key, pair.Value); }
        }

        public void EnsureColumn(string column) { EnsureColumn(column, column); }
        public void EnsureColumn(string column, string alternateName)
        {
            if (!ContainsKey(column))
            {
                AddColumn(column, alternateName);
            }
            alternateNameMap[column] = alternateName;
        }

        public string ColumnStatement
        {
            get
            {
                return String.Join(",", Keys.Select(k => String.Format("[{0}]", k)));
            }
        }

        public Dictionary<string, int> AlternateDictionary
        {
            get
            {
                return this.ToDictionary(d => alternateNameMap[d.Key], d => d.Value);
            }
        }

        public string GetAlternateName(string name) { return alternateNameMap[name]; }

        Dictionary<string, string> alternateNameMap;
        public void AddColumn(string column) { AddColumn(column, column); }
        public void AddColumn(string column, string alternateName)
        {
            if (Keys.Contains(column)) { throw new Exception("Column \"" + column + "\" already in reader map."); }
            Add(column, Count);
            alternateNameMap.Add(column, alternateName);
        }
    }

    public static class UtilityExt
    {

        static Dictionary<Guid, string> databaseDic = new Dictionary<Guid, string>();
        public static string GetDatabaseName(this SPSite site)
        {
            if (!databaseDic.ContainsKey(site.ID))
            {
                databaseDic.Add(site.ID, site.ContentDatabase.Name);
            }
            return databaseDic[site.ID];
        }

        public static string GetSPTableLocation(this SPSite site, string tableName)
        {
            return String.Format("[{0}].[dbo].[{1}]", site.GetDatabaseName(), tableName);
        }


        public static void EndSerialization(this StringBuilder builder)
        {
            if (builder.Length > 1)
            {
                builder.Remove(builder.Length - 1, 1);
            }
            builder.Append(builder[0] == '{' ? "}" : "]");
        }

        public static string GetStrValue(this DataRow row, string column)
        {
            try
            {
                return row[column] != null ? row[column].ToString() : null;
            }
            catch
            {
            }
            return null;
        }

        public static string GetStrValue(this SqlDataReader reader, int index)
        {
            try
            {
                return reader.GetValue(index).ToString();
            }
            catch
            {
            }
            return null;
        }

        public static string GetStrValue(this SPListItem item, string field)
        {
            try
            {
                return item[field] != null ? item[field].ToString() : null;
            }
            catch
            {
            }
            return null;
        }

        public static int GetIntValue(this DataRow row, string column)
        {
            try
            {
                string strValue = row[column] != null ? row[column].ToString() : null;
                if (!String.IsNullOrWhiteSpace(strValue))
                {
                    return Int32.Parse(strValue);
                }
            }
            catch
            {
            }
            return 0;
        }

        public static string ShiftBar(this string str)
        {
            if (str.Contains("|"))
            {
                str = str.Substring(str.IndexOf("|") + 1);
            }
            return str;
        }

        public static string GetStripIDValue(this DataRow row, string column)
        {
            try
            {
                object value = row[column];
                if (value == null)
                {
                    return null;
                }
                string strValue = value.ToString();
                int poundIndex = strValue.IndexOf("#");
                if (poundIndex != -1 && poundIndex <= strValue.Length)
                {
                    strValue = strValue.Substring(poundIndex + 1);
                }
                return strValue;
            }
            catch
            {
                return null;
            }
        }

        static DateTime defDate = default(DateTime);
        public static DateTime GetDateValue(this DataRow row, string column)
        {
            try
            {
                object val = row[column];
                if (val != null)
                {
                    DateTime time;
                    if (DateTime.TryParse(val.ToString(), out time))
                    {
                        return time;
                    }
                }
                return defDate;
            }
            catch
            {
            }
            return defDate;
        }

        public static UInt16 ToDosDate(this DateTime time)
        {
            return (UInt16)(((time.Year - 1980) << 9) | (time.Month << 5) | time.Day);
        }

        public static UInt16 ToDosTime(this DateTime time)
        {
            return (UInt16)((time.Hour << 11) | (time.Minute << 6) | ((time.Second << 1) / 2));
        }

        public static void AppendAttributes(this XmlElement element, params XmlAttribute[] attributes)
        {
            foreach (XmlAttribute attribute in attributes)
            {
                element.Attributes.Append(attribute);
            }
        }
    }

    public enum RulePredicate { Equals, NotEqual, Contains, StartsWith, EndsWith, LessThan, HigherThan }

    public static class StrExt
    {
        public static string StrJoin(this IEnumerable<string> e, string connector = null)
        {
            StringBuilder strBuilder = new StringBuilder();
            string _connector = connector == null ? String.Empty : connector;
            var en = e.GetEnumerator();
            en.MoveNext();
            strBuilder.Append(en.Current);
            if (connector == null)
            {
                while (en.MoveNext())
                {
                    strBuilder.Append(en.Current);
                }
            }
            else
            {
                while (en.MoveNext())
                {
                    strBuilder.Append(_connector);
                    strBuilder.Append(en.Current);
                }
            }
            return strBuilder.ToString();
            /*var array = e.ToArray();
            return connector != null ? String.Join(connector, array) : String.Concat(array);*/
        }


        public static string AddAttribute(this string str, string name, string value)
        {
            try
            {
                var element = XElement.Parse(str);
                element.Add(new XAttribute(name, value));
                return element.ToString(SaveOptions.DisableFormatting);
            }
            catch
            {
                return str;
            }
        }

        public static string AggregateFormat(this IEnumerable<string> e, string format)
        {
            return e.Aggregate((s, i) => String.Format(format, s, i));
        }

        public static string UseAsFormat(this string str, params object[] objects)
        {
            try
            {
                return String.Format(str, objects);
            }
            catch (FormatException exception)
            {
                return null;
            }
        }
    }

    public enum CAMLValueType
    {
        Text = 0,
        Lookup,
        Number,
        Boolean,
        Guid,
        DateTime,
        UniqueId,
        Integer
    }

    public class FieldDefinition
    {
        public string InternalName;
        public string DisplayName;

        public override string ToString()
        {
            return InternalName;
        }
        public override bool Equals(object obj)
        {
            var fd = obj as FieldDefinition;
            return fd != null && fd.InternalName == InternalName;
        }
        public override int GetHashCode()
        {
            return (InternalName != null ? InternalName : String.Empty).GetHashCode();
        }
    }

    public static class Q
    {
        public const string WhereFormat = "<Where>{0}</Where>";
        public const string AndFormat = "<And>{0}{1}</And>";
        public const string ContainsFormat = "<Contains>{0}{1}</Contains>";
        public const string BeginsWithFormat = "<BeginsWith>{0}{1}</BeginsWith>";
        public const string EndsWithFormat = "<EndsWith>{0}{1}</EndsWith>";
        public const string GreaterThanFormat = "<Gt>{0}{1}</Gt>";
        public const string LessThanFormat = "<Lt>{0}{1}</Lt>";
        public const string InFormat = "<In>{0}{1}</In>";
        public const string OrFormat = "<Or>{0}{1}</Or>";
        public const string OrderByFormat = "<OrderBy>{0}</OrderBy>";
        public const string EqFormat = "<Eq>{0}{1}</Eq>";
        public const string NeqFormat = "<Neq>{0}{1}</Neq>";
        public const string FieldRefIdFormat = "<FieldRef Id=\"{0}\"/>";
        public const string FieldRefIdFormatLookup = "<FieldRef Id=\"{0}\" />";
        public const string FieldRefFormat = "<FieldRef Name=\"{0}\"/>";
        public const string FieldRefFormatLookup = "<FieldRef Name=\"{0}\" LookupId=\"TRUE\"/>";
        public const string ValuesFormat = "<Values>{0}</Values>";
        public const string ValueTypeFormat = "<Value Type=\"{1}\">{0}</Value>";
        public const string ValueFormat = "<Value Type=\"Text\">{0}</Value>";
        public const string LookupValueFormat = "<Value Type=\"Lookup\">{0}</Value>";
        public const string NumberValueFormat = "<Value Type=\"Number\">{0}</Value>";
        public const string BooleanValueFormat = "<Value Type=\"Boolean\">{0}</Value>";
        public const string DateTimeValueFormat = "<Value Type=\"DateTime\" IncludeTimeValue=\"TRUE\">{0}</Value>";
        public const string IsNotNullFormat = "<IsNotNull><FieldRef Name=\"{0}\"/></IsNotNull>";
        public const string IsNullFormat = "<IsNull><FieldRef Name=\"{0}\"/></IsNull>";
        public const string WebsFormat = "<Webs Scope=\"{0}\"/>";
        public const string ViewFormat = "<View{0}>{1}</View>";
        public const string ScopeRecursive = "Scope=\"Recursive\"";
        public const string QueryFormat = "<Query>{0}</Query>";



        public static string View(string query, string viewFields = null, bool recursive = false)
        {
            return String.Format(ViewFormat,
                recursive ? " " + ScopeRecursive : "",
                (String.IsNullOrWhiteSpace(query) ? "" : String.Format(QueryFormat, query)),
                (String.IsNullOrWhiteSpace(viewFields) ? "" : viewFields));
        }

        public static string Value(string value, CAMLValueType type)
        {
            switch (type)
            {
                case CAMLValueType.Integer:
                    return String.Format(ValueTypeFormat, value, type.ToString());
                case CAMLValueType.Number:
                    return NumberValueFormat.UseAsFormat(value);
                case CAMLValueType.Lookup:
                case CAMLValueType.UniqueId:
                    return LookupValueFormat.UseAsFormat(value);
                case CAMLValueType.Boolean:
                    return BooleanValueFormat.UseAsFormat(Convert.ToInt32(ParseBool(value)).ToString());
                case CAMLValueType.DateTime:
                    return BooleanValueFormat.UseAsFormat(Convert.ToInt32(ParseBool(value)).ToString());
                default:
                    return ValueFormat.UseAsFormat(value);
            }
        }

        public static string OrderByID
        {
            get
            {
                return OrderBy(true, "ID");
            }
        }

        public static string Webs(bool recursive)
        {
            return Q.WebsFormat.UseAsFormat(recursive ? "True" : "False");
        }

        public static string In(string fieldName, IEnumerable<string> values, CAMLValueType type = CAMLValueType.Text)
        {
            return InFormat.UseAsFormat((type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup).UseAsFormat(fieldName), Values(values, type));
        }

        public static string Values(IEnumerable<string> values, CAMLValueType type = CAMLValueType.Text)
        {
            if (values == null || !values.Any()) { return String.Empty; }
            return ValuesFormat.UseAsFormat(values.Select(v => Value(v, type)).StrJoin());
        }

        public static string Condition(string field, string value, RulePredicate exclusionPredicate, CAMLValueType type = CAMLValueType.Text)
        {
            switch (exclusionPredicate)
            {
                case RulePredicate.Equals:
                    return Condition(field, value, type);
                case RulePredicate.NotEqual:
                    return NCondition(field, value, type);
                case RulePredicate.Contains:
                    return CCondition(field, value, type);
                case RulePredicate.StartsWith:
                    return BCondition(field, value, type);
                case RulePredicate.EndsWith:
                    return ECondition(field, value, type);
                case RulePredicate.HigherThan:
                    return GCondition(field, value, type);
                case RulePredicate.LessThan:
                    return LCondition(field, value, type);
            }
            return null;
        }

        public static string Condition(FieldDefinition fieldDefinition, string value, RulePredicate exclusionPredicate, CAMLValueType type = CAMLValueType.Text)
        {
            return Condition(fieldDefinition.InternalName, value, exclusionPredicate, type);
        }

        public static string Condition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return Condition(pair.Key, pair.Value.ToString());
        }

        public static string Condition(Guid field, string value, RulePredicate exclusionPredicate, CAMLValueType type = CAMLValueType.Text)
        {
            switch (exclusionPredicate)
            {
                case RulePredicate.Equals:
                    return Condition(field, value, type);
                case RulePredicate.NotEqual:
                    return NCondition(field, value, type);
                case RulePredicate.Contains:
                    return CCondition(field, value, type);
                case RulePredicate.StartsWith:
                    return BCondition(field, value, type);
                case RulePredicate.EndsWith:
                    return ECondition(field, value, type);
                case RulePredicate.HigherThan:
                    return GCondition(field, value, type);
                case RulePredicate.LessThan:
                    return LCondition(field, value, type);
            }
            return null;
        }

        public static string NCondition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return NCondition(pair.Key, pair.Value.ToString(), type);
        }

        public static string NCondition(FieldDefinition fieldDefinition, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return NCondition(fieldDefinition.InternalName, value, type);
        }
        public static string NCondition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(NeqFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string CCondition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return CCondition(pair.Key, pair.Value.ToString(), type);
        }

        public static string CCondition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(ContainsFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string BCondition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return BCondition(pair.Key, pair.Value.ToString(), type);
        }

        public static string BCondition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(BeginsWithFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string ECondition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return ECondition(pair.Key, pair.Value.ToString(), type);
        }

        public static string ECondition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(EndsWithFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string GCondition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return GCondition(pair.Key, pair.Value.ToString(), type);
        }

        public static string GCondition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(GreaterThanFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string LCondition(StrPair pair, CAMLValueType type = CAMLValueType.Text)
        {
            return LCondition(pair.Key, pair.Value.ToString(), type);
        }

        public static string LCondition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(LessThanFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string Condition(string fieldName, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(EqFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefFormat : FieldRefFormatLookup, fieldName), Value(value, type));
        }

        public static string Condition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(EqFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string NCondition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(NeqFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string GCondition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(GreaterThanFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string CCondition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(ContainsFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string LCondition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(LessThanFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string ECondition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(EndsWithFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string BCondition(Guid fieldId, string value, CAMLValueType type = CAMLValueType.Text)
        {
            return String.Format(BeginsWithFormat, String.Format(type != CAMLValueType.Lookup ? FieldRefIdFormat : FieldRefIdFormatLookup, fieldId), Value(value, type));
        }

        public static string OrderBy(bool ascending, params string[] fields)
        {
            return String.Format(OrderByFormat, fields.Select(f => String.Format(FieldRefFormat.AddAttribute("Ascending", "{1}"), f, ascending ? "TRUE" : "FALSE")).ToArray().StrJoin());
        }

        public static string OrderBy(bool ascending, params FieldDefinition[] fields)
        {
            return OrderBy(ascending, fields.Select(f => f.InternalName).ToArray());
        }

        public static string WhereAnd(params string[] conditions)
        {
            return String.Format(WhereFormat, conditions.AggregateFormat(AndFormat));
        }

        public static string WhereOr(params string[] conditions)
        {
            return String.Format(WhereFormat, conditions.AggregateFormat(OrFormat));
        }

        public static string Where(string condition)
        {
            return String.Format(WhereFormat, condition);
        }

        public static string Query(string condition, bool ascending = true, params string[] orderByFields)
        {
            string where = condition == null ? "" : Where(condition);
            if (orderByFields == null || orderByFields.Length == 0)
            {
                return where;
            }
            return OrderBy(ascending, orderByFields) + where;
        }

        public static string IsNotNull(string field)
        {
            return String.Format(IsNotNullFormat, field);
        }

        public static string AreNotNull(params string[] fields)
        {
            return Ands(fields.Select(f => IsNotNull(f)).ToArray());
        }

        public static string IsNull(string field)
        {
            return String.Format(IsNullFormat, field);
        }

        public static string AreNull(params string[] fields)
        {
            return Ands(fields.Select(f => IsNull(f)).ToArray());
        }

        public static string And(string a, string b)
        {
            return String.Format(AndFormat, a, b);
        }

        public static string Ands(params string[] conditions)
        {
            return conditions.Aggregate((a, c) => And(a, c));
        }

        public static string Or(string a, string b)
        {
            return String.Format(OrFormat, a, b);
        }

        public static string Ors(params string[] conditions)
        {
            return conditions.Aggregate((a, c) => Or(a, c));
        }

        public static string AndEqs(params StrPair[] pairs)
        {
            return Eqs(pairs).AggregateFormat(AndFormat);
        }

        public static string OrEqs(params StrPair[] pairs)
        {
            return Eqs(pairs).AggregateFormat(OrFormat);
        }

        public static string OrNeqs(params StrPair[] pairs)
        {
            return Neqs(pairs).AggregateFormat(OrFormat);
        }

        public static string AndNeqs(params StrPair[] pairs)
        {
            return Neqs(pairs).AggregateFormat(AndFormat);
        }

        public static string ViewFields(params FieldDefinition[] viewFields)
        {
            return ViewFields(viewFields.Select(f => f.InternalName).ToArray());
        }
        public static string ViewFields(params string[] viewFields)
        {
            if (viewFields == null || viewFields.Length == 0)
            {
                return null;
            }
            return String.Concat(viewFields.Select(f => String.Format(FieldRefFormat, f)).ToArray());
        }

        public static string ViewFields(params Guid[] viewFields)
        {
            if (viewFields == null || viewFields.Length == 0)
            {
                return null;
            }
            return String.Concat(viewFields.Select(f => String.Format(FieldRefIdFormat, f)).ToArray());
        }

        public static IEnumerable<string> Eqs(params StrPair[] pairs)
        {
            return pairs.Select(p => Condition(p));
        }

        public static IEnumerable<string> Neqs(params StrPair[] pairs)
        {
            return pairs.Select(p => NCondition(p));
        }

        static Regex boolFinder = new Regex(@"^(?i)(Yes|No|True|False|1|0)$");
        static Regex trueFinder = new Regex(@"^(?i)(Yes|True|1)$");

        public static bool IsBool(string value)
        {
            return boolFinder.IsMatch(value);
        }

        public static bool ParseBool(string value)
        {
            return trueFinder.IsMatch(value);
        }
    }

    public class StrPair
    {
        public string Key
        {
            get;
            set;
        }
        public string Value
        {
            get;
            set;
        }

        public StrPair()
        {
        }

        public StrPair(string key)
        {
            Key = key;
        }

        public StrPair(string key, string value)
            : this(key)
        {
            Value = value;
        }

        public static List<StrPair> Pairs(params string[] nameValueNameEtc)
        {
            List<StrPair> pairs = new List<StrPair>();
            for (int i = 0; i < nameValueNameEtc.Length - 1; i = i + 2)
            {
                pairs.Add(new StrPair(nameValueNameEtc[i], nameValueNameEtc[i + 1]));
            }
            if (nameValueNameEtc.Length % 2 != 0)
            {
                pairs.Add(new StrPair(nameValueNameEtc.Last(), String.Empty));
            }
            return pairs;
        }
    }

    public static class B
    {
        public const string methodFormatId = "<Method ID='{1}'>{0}</Method>";
        public const string fieldNameFormat = "<Field Name='{0}' >{1}</Field>";
        public const string updateMethodFormat = "<Method ID='{0}' Cmd='Update'>{1}</Method>";
        public const string batchFormat = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><ows:Batch OnError=\"Continue\">{0}</ows:Batch>";
        public const string methodFormat = "<Method>{0}</Method>";
        public const string listsBatchUpdateFormat = "<Batch OnError=\"Continue\" ListVersion=\"1\">{0}</Batch>";
        public const string setListFormat = "<SetList{1}>{0}</SetList>";
        public const string setVarFormat = "<SetVar Name=\"{0}\">{1}</SetVar>";
        public const string getVarFormat = "<GetVar Name=\"{0}\"/>";
        public const string columnPrefix = "urn:schemas-microsoft-com:office:office#";
        const string scopeRequest = " Scope=\"Request\"";

        public static string ListsUpdateItem(int itemId, params StrPair[] values)
        {
            string fields = fieldNameFormat.UseAsFormat("ID", itemId);
            if (values != null && values.Length > 0)
            {
                fields += values.Select(v => fieldNameFormat.UseAsFormat(v.Key, v.Value)).StrJoin();
            }
            return updateMethodFormat.UseAsFormat(itemId, fields);
        }

        public static string DeleteItems(string listId, params int[] itemIds)
        {
            return batchFormat.UseAsFormat(String.Concat(itemIds.Select(i => DeleteMethod(listId, i)).ToArray()));
        }

        public static string DeleteMethod(string listId, int itemId)
        {
            string setVars =
                   String.Format(setVarFormat, "ID", itemId) +
                   String.Format(setVarFormat, "Cmd", "Delete");
            return String.Format(methodFormat, String.Concat(String.Format(setListFormat, listId.ToString(), scopeRequest), setVars));
        }

        public static string SetVar(string name, string value, bool prefix = true)
        {
            return setVarFormat.UseAsFormat((prefix ? columnPrefix : "") + name, value);
        }

        public static string UpdateMethod(string listId, string itemId, params StrPair[] values)
        {
            string setVars = setVarFormat.UseAsFormat("ID", itemId) +
               String.Format(setVarFormat, "Cmd", "Save") +
             values.Select(s => SetVar(s.Key, s.Value)).StrJoin();
            return String.Format(methodFormat, String.Concat(String.Format(setListFormat, listId, ""), setVars));
        }
        public static string UpdateMethod(string listId, string itemId, Dictionary<string, string> item)
        {
            return UpdateMethod(listId, itemId, item.Select(i => new StrPair(i.Key, i.Value)).ToArray());
        }

        public static string AddMethod(string listId, params StrPair[] values)
        {
            return UpdateMethod(listId, "New", values);
        }

        public static string AddMethod(string listId, Dictionary<string, string> item)
        {
            return UpdateMethod(listId, "New", item.Select(i => new StrPair(i.Key, i.Value)).ToArray());
        }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class StructureID : Attribute
    {
    }

    public class StructureReader<T>
    {
        static Type tType = typeof(T);
        Dictionary<Type, PropertyInfo> propDic;
        Structure<T> global;
        int depth;
        public StructureReader(Structure<T> structure)//, IEnumerable[] enumerables)
        {
            global = structure;
            depth = global.Depth;
            /* if (structure.Depth != enumerables.Length)
             {
                 throw new StructureDepthMismatchException(structure.Depth, enumerables.Length);
             }
             propDic = new Dictionary<Type, PropertyInfo>();
             var ee = enumerables.GetEnumerator();
             while (ee.MoveNext())
             {
                 var e = ee.Current;
                 Type eT = e.GetType();
                 if (!eT.IsGenericType)
                 {
                     throw new EnumerableTypeNonGenericException(eT);
                 }
                 Type genType = eT.GenericTypeArguments[0];
                 PropertyInfo prop = genType.GetProperties(BindingFlags.Public | BindingFlags.Instance).FirstOrDefault(p => p.GetCustomAttribute<StructureID>() != null);
                 if (prop == null)
                 {
                     throw new GenericTypeNonStructurableException(genType);
                 }
                 else if (prop.PropertyType != tType)
                 {
                     throw new GenericTypeMismatchException(prop);
                 }
                 propDic.Add(genType, prop);
             }*/
        }

        Structure<T> currentStruct;
        void SeekOriginal()
        {
            Structure<T> stru = global;
            while (stru.Count > 0)
            {
                stru = stru.First().Value;
                indices.Add(0);
            }
            currentStruct = stru.Parent;
        }

        public bool Read()
        {
            if (currentStruct == null)
            {
                SeekOriginal();
            }
            return true;
        }

        List<int> indices = new List<int>();
        public IEnumerable<T> Current
        {
            get
            {
                List<T> values = new List<T>(depth);
                T t = currentStruct.ElementAt(indices.Last()).Key;
                values.Add(t);
                Structure<T> read = currentStruct.Parent;
                while (read != null)
                {
                    // values.Add(
                }
                return null;
            }
        }

        public class StructureDepthMismatchException : Exception
        {
            public StructureDepthMismatchException(int depth, int enumerableLength)
                : base(String.Format("Cannot read structure of depth {0} with enumerable array length {1}", depth, enumerableLength))
            {
            }
        }

        public class EnumerableTypeNonGenericException : Exception
        {
            public EnumerableTypeNonGenericException(Type type)
                : base(String.Format("Enumerable type \"{0}\" is not generic and therefore cannot be used to build structure",
                type.FullName))
            {
            }
        }

        public class GenericTypeMismatchException : Exception
        {
            public GenericTypeMismatchException(PropertyInfo prop)
                : base(String.Format("Property \"{0}\" of type \"{1}\" for generic type \"{3}\" does not match structure key type \"{2}\"",
                prop.Name,
                prop.PropertyType.FullName,
                tType.FullName,
                prop.DeclaringType.FullName))
            {
            }
        }
        public class GenericTypeNonStructurableException : Exception
        {
            public GenericTypeNonStructurableException(Type type)
                : base(String.Format("Type \"{0}\" is missing required \"Montrium.Navigator.StructureID\" attribute for one of its public properties of type \"{1}\"",
                type.FullName,
                tType.FullName))
            {
            }
        }
    }

    public class Structure<T> : Dictionary<T, Structure<T>>
    {
        public Structure<T> Parent
        {
            get;
            internal set;
        }

        private int depth = 0;
        public int Depth
        {
            get { return depth; }
            internal set
            {
                depth = value;
                if (Parent != null && Parent.Depth <= depth)
                {
                    Parent.Depth = depth + 1;
                }
            }
        }

        Type type;
        MethodInfo addMethod;
        MethodInfo setMethod;
        public Structure()
        {
            type = this.GetType();
            Type[] paramTypes = new[] { typeof(T), typeof(T) };
            addMethod = type.GetMethod("Add", BindingFlags.Public | BindingFlags.Instance, null, paramTypes, null);
            setMethod = type.GetMethod("set_Item", BindingFlags.Public | BindingFlags.Instance, null, paramTypes, null);
        }

        public Structure(Structure<T> parent)
            : this()
        {
            Parent = parent;
        }

        public Structure<T> this[params T[] path]
        {
            get
            {
                var dic = base[path.First()];
                foreach (T t in path.Skip(1))
                {
                    if (dic == null || !dic.ContainsKey(t))
                    {
                        return null;
                    }
                    dic = dic[t];
                }
                return dic;
            }
        }

        public void Add(T key, T value)
        {
            if (!ContainsKey(key))
            {
                Add(key, new Structure<T>(this));
            }
            base[key].Add(value);
        }

        public void Add(T key)
        {
            if (!ContainsKey(key))
            {
                if (Count == 0)
                {
                    Depth++;
                }
                Add(key, new Structure<T>(this));
            }
        }

        public void Add(T val, params T[] path)
        {
            Structure<T> stru = this;
            foreach (T pt in path)
            {
                if (!stru.ContainsKey(pt))
                {
                    if (stru.Count == 0) { stru.Depth++; }
                    stru.Add(pt, new Structure<T>(stru));
                }
                stru = stru[pt];
            }
            stru.Add(val);
        }

        public bool Remove(Structure<T> stru)
        {
            foreach (T t in stru.Keys)
            {
                if (ContainsKey(t))
                {
                    base[t].Remove(stru[t]);
                    if (base[t].Count == 0)
                    {
                        base.Remove(t);
                    }
                }
            }
            return true;
        }

        public bool RemoveByPath(params T[] path)
        {
            Structure<T> stru = this;
            List<Structure<T>> strus = new List<Structure<T>>();
            foreach (T pt in path)
            {
                if (!stru.ContainsKey(pt))
                {
                    return false;
                }
                strus.Add(stru);
                stru = stru[pt];
            }
            strus.Reverse();
            List<T> rpath = path.Reverse().ToList();
            strus[0].Remove(rpath[0]);
            strus.RemoveAt(0);
            rpath.RemoveAt(0);
            for (int i = 0; i < rpath.Count; i++)
            {
                var s = strus[i];
                T p = rpath[i];
                if (s[p].Count == 0)
                {
                    s.Remove(p);
                }
            }
            return true;
        }
    }

    public static class oX
    {
        static Random randy = new Random();
        public static string RandomId()
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < 16; i++)
            {
                string c = alphaNum[randy.Next(0, alphaNum.Length)].ToString();
                builder.Append(c);
            }
            return builder.ToString();
        }
        const string alphaNum = "abcdefghijklmnopqrstuvwxyz0123456789";
        public const string RelsForm = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{0}" +
            @"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"" Target=""/xl/sharedStrings.xml"" Id=""R062dbc20e82e4ed8"" />
  <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""Rb932f6550a814ff9"" />
        </Relationships>";
        public const string RelForm = "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{1}\" Target=\"{2}\" Id=\"{0}\" />";
        public static string GetColumnReference(int index)
        {
            int original = index;
            string reference = "";
            do
            {
                int remainder = index % 26;
                reference = GetColumnChar(remainder) + reference;
                index = (index - remainder) / 26;
            }
            while (index != 0);
            return reference;
        }

        static string GetColumnChar(int index)
        {
            return ((char)((int)'A' + index - 1)).ToString();
        }

        public const string WorkbookSt = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<x:sheets>";

        public const string ContentTypes = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"" />
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
  {0}
  <Override PartName=""/xl/sharedStrings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"" />
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" />
</Types>";

        public const string SheetType = @"<Override PartName=""/xl/worksheets/{0}"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"" />";

        public static string F(this string str, string content)
        {
            return String.Format(str, content);
        }
        public static string F(this string str, IEnumerable<string> content)
        {
            return String.Format(str, content.ToArray());
        }
        public const string Head = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
        public const string xmlns = "xmlns:x";
        public const string xmlnsr = "xmlns:r";
        public const string spmain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public const string op = "<x:";
        public const string cl = "</x:";
        public static readonly string MGs = el("mergeCells", true);
        public static readonly string Cols = el("cols", true);
        static readonly string[][] sheetAttr = buildArr(new[] { xmlnsr, "name", "sheetId", "r:id" });
        public static string Col(int min, int max, int width, bool custom)
        {
            provision(colAttrArr, new[] { min.ToString(), max.ToString(), width.ToString(), custom ? "1" : "0" });
            return el("col", false, colAttrArr);
        }
        public static string MG(int startRow, int endRow, int startCol, int endCol)
        {
            return el("mergeCell", false, "ref",
                GetColumnReference(startCol) + startRow + ":" +
                GetColumnReference(endCol) + endRow);
        }

        public static string MG(int row, int startCol, int endCol)
        {
            return MG(row, row, startCol, endCol);
        }
        static string[][] stateAttr = buildArr(new[] { "state" });
        public static string S(string name, int sheetIndex, string sheetId, bool hidden = false)
        {
            sheetAttr[0][1] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            sheetAttr[1][1] = name;
            sheetAttr[2][1] = sheetIndex.ToString();
            sheetAttr[3][1] = sheetId;
            if (hidden)
            {
                stateAttr[0][1] = "hidden";
                return el("sheet", false, sheetAttr.Concat(stateAttr).ToArray());
            }
            return el("sheet", false, sheetAttr);
        }
        public static readonly string SD = el("sheetData", true);
        public static readonly string WS = el("worksheet", true);
        public static readonly string Row = el("row", true);
        public static readonly string R = el("r", true);
        public static readonly string B = "<x:b />";
        public static readonly string T = el("t", true);
        public static readonly string RP = el("rPr", true);
        public static string C(string v, string t, string r, string s)
        {
            cellAttrArr[0][1] = t;
            cellAttrArr[1][1] = r;
            cellAttrArr[2][1] = s;
            return el("c", true, cellAttrArr).F(el("v", true).F(v));
        }

        static void provision(string[][] arr, string[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                arr[i][1] = values[i];
            }
        }

        static readonly string[][] cellAttrArr = buildCellAttrArr();
        static readonly string[][] colAttrArr = buildArr(new[] { "min", "max", "width", "custom" });

        static string[][] buildCellAttrArr()
        {
            string[][] arr = attrArr(3, 2);
            arr[0][0] = "t";
            arr[1][0] = "r";
            arr[2][0] = "s";
            return arr;
        }

        public static string[][] buildArr(string[] values)
        {
            string[][] arr = attrArr(values.Length, 2);
            for (int i = 0; i < values.Length; i++)
            {
                arr[i][0] = values[i];
            }
            return arr;
        }

        static string[][] attrArr(int i, int j)
        {
            string[][] arr = new string[i][];
            for (int z = 0; z < i; z++)
            {
                arr[z] = new string[j];
            }
            return arr;
        }
        public static string Color(string color)
        {
            return el("color", false, "rgb", color);
        }
        public static readonly string SI = el("si", true);
        public static readonly string SST = el("sst", true, "smlns:x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        public static string C(string type, string refer, int styleIndex)
        {
            provision(cellAttrArr, new[] { type, refer, styleIndex != -1 ? styleIndex.ToString() : null });
            return el("c", true, cellAttrArr);
        }
        public static string el(string t, bool child)
        {
            return el(t, child, null);
        }

        public static string elSt(string t, string[][] attrs)
        {
            return op + t + getAttrStr(attrs) + ">";
        }

        public static string elSt(string t, string prop, string value)
        {
            return op + t + " " + prop + "=\"" + value + "\">";
        }

        public static string elSt(string t)
        {
            return elSt(t, null);
        }

        public static string elCl(string t)
        {
            return cl + t + ">";
        }

        static string getAttrStr(string[][] attrs)
        {
            string attrStr = null;
            if (attrs != null)
            {
                StringBuilder builder = new StringBuilder();
                foreach (var att in attrs)
                {
                    if (att[1] == null) { continue; }
                    builder.Append(" " + att[0] + "=\"" + att[1] + "\"");
                    attrStr = builder.ToString();
                }
            }
            return attrStr;
        }
        public static string el(string t, bool child, string[][] attrs)
        {
            string attrStr = null;
            if (attrs != null)
            {
                StringBuilder builder = new StringBuilder();
                foreach (var att in attrs)
                {
                    if (att[1] == null) { continue; }
                    builder.Append(" " + att[0] + "=\"" + att[1] + "\"");
                    attrStr = builder.ToString();
                }
            }
            return op + t + attrStr + (child ? ">{0}" + cl + t + ">" : " />");
        }
        public static string el(string t, bool child, string prop, string value)
        {
            return op + t + " " + prop + "=\"" + value + "\"" + (child ? ">{0}" + cl + t + ">" : "/>");
        }
    }

    public class FieldMap : IEnumerable<KeyValuePair<string, string>>, IEnumerable
    {
        /* static */
        Dictionary<Guid, Dictionary<string, string>> cache = new Dictionary<Guid, Dictionary<string, string>>();

        static FieldInfo mapField = typeof(SPListItemCollection).GetField("m_mapFields", BindingFlags.NonPublic | BindingFlags.Instance);
        static MethodInfo EnsureFieldMap = typeof(SPListItemCollection).GetMethod("EnsureFieldMap", BindingFlags.NonPublic | BindingFlags.Instance);
        static FieldInfo fieldNames = null;
        static MethodInfo GetColNumber = null;
        Dictionary<string, string> dic;

        public bool ContainsField(string fieldName) { return dic.ContainsKey(fieldName); }

        public bool ContainsField(Guid guid) { return dic.Keys.Any(k => new Guid(GetFieldId(k)) == guid); }

        public string this[string fieldName]
        {
            get
            {
                return dic.ContainsKey(fieldName) ? dic[fieldName] : null;
            }
        }

        Dictionary<Guid, string> nameDic = new Dictionary<Guid, string>();
        public string GetFieldNameById(Guid id)
        {
            string strId = id.ToString().ToLower();
            if (nameDic.ContainsKey(id)) { return nameDic[id]; }
            var e = dic.Keys.GetEnumerator();
            while (e.MoveNext())
            {
                string idStr = GetFieldId(e.Current);
                if (idStr == null) { continue; }
                Guid guid = new Guid(idStr);
                if (nameDic.ContainsKey(guid)) { continue; }
                nameDic.Add(guid, e.Current);
                if (guid == id)
                {
                    return e.Current;
                }
            }
            return null;
        }

        static readonly string[] lookupTypes = { "Lookup", "ArtfulBits.CascadedLookup", "User", "LookupMulti" };
        List<string> lookupFields;
        static readonly string[] exemptLookupFields = { "FileRef" };
        public List<string> LookupFields
        {
            get
            {
                if (lookupFields == null)
                {
                    lookupFields = ParsedXml.Element("Fields")
                        .Elements("Field").Where(f => lookupTypes.Contains(f.Attribute("Type").Value.ToString()))
                        .Select(f => f.Attribute("Name").Value.ToString()).Where(n => !exemptLookupFields.Contains(n)).ToList();
                }
                return lookupFields;
            }
        }

        XElement parsedXml;
        XElement ParsedXml
        {
            get
            {
                if (parsedXml == null)
                {
                    parsedXml = XElement.Parse(Xml);
                }
                return parsedXml;
            }
        }

        public Guid ListId
        {
            get
            {
                return new Guid(ParsedXml.Attribute("ID").Value);
            }
        }

        public Guid WebId
        {
            get
            {
                return new Guid(ParsedXml.Attribute("WebId").Value);
            }
        }

        public string Xml
        {
            get;
            private set;
        }

        public FieldMap(Guid listId, Stream deflatedStream) : this(listId, default(Guid), deflatedStream) { }
        public FieldMap(Guid listId, Guid WebId, Stream deflatedStream)
        {
            StringBuilder builder = new StringBuilder();
            using (MemoryStream stream = new MemoryStream())
            {
                deflatedStream.Seek(14, SeekOrigin.Begin);
                using (DeflateStream decoder = new DeflateStream(deflatedStream, CompressionMode.Decompress))
                {
                    decoder.CopyTo(stream);
                }
                stream.Seek(0, SeekOrigin.Begin);
                while (stream.Position != stream.Length)
                {
                    builder.Append((char)stream.ReadByte());
                }
            }
            deflatedStream.Dispose();
            string xmlStr = builder.ToString();
            BuildDic(listId, String.Format("<List ID=\"{0}\" WebId=\"{1}\"><Fields>{2}</Fields></List>", listId, WebId, xmlStr.Substring(xmlStr.IndexOf("<"))));
            dic = cache[listId];
        }

        public IEnumerable<string> MappedFields
        {
            get { return dic.Keys.Where(f => !LookupFields.Contains(f)); }
        }

        public FieldMap(SPList list)
        {
            if (!cache.ContainsKey(list.ID))
            {
                BuildDic(list.ID, list.SchemaXml);
            }
            dic = cache[list.ID];
        }

        Dictionary<Guid, string> multiFieldDictionary;
        public Dictionary<Guid, string> MultiFieldDictionary
        {
            get
            {
                if (multiFieldDictionary == null)
                {
                    multiFieldDictionary = XElement.Parse(Xml).Element("Fields").Elements("Field")
                        .Where(f => IsMulti(f)).ToDictionary(f => new Guid(f.Attribute("ID").Value), f => f.Attribute("Name").Value);
                }
                return multiFieldDictionary;
            }
        }

        bool IsMulti(XElement fieldElement)
        {
            XAttribute type = fieldElement.Attribute("Type");
            if (type == null) { return false; }
            string value = type.Value;
            return value.IndexOf("Lookup") != -1 && value.IndexOf("Multi") != -1;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return dic.GetEnumerator();
        }

        public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            return dic.GetEnumerator();
        }


        public string GetFieldAttribute(string fieldName, string fieldAttribute)
        {
            try
            {
                return GetFieldElement(fieldName).Attribute(fieldAttribute).Value.ToString();
            }
            catch
            {
            }
            return null;
        }

        public string GetFieldId(string fieldName)
        {
            try
            {
                return GetFieldElement(fieldName).Attribute("ID").Value.ToString();
            }
            catch
            {
            }
            return null;
        }

        Dictionary<string, XElement> fieldElementDic = new Dictionary<string, XElement>();
        XElement GetFieldElement(string fieldName)
        {
            try
            {
                if (!fieldElementDic.ContainsKey(fieldName))
                {
                    fieldElementDic.Add(fieldName, ParsedXml.Element("Fields").Elements("Field").First(f => f.Attribute("Name").Value.ToString() == fieldName));
                }
                return fieldElementDic[fieldName];
            }
            catch
            {
            }
            return null;
        }

        public string GetFieldTypeAsString(string fieldName)
        {
            try
            {
                return GetFieldElement(fieldName).Attribute("Type").Value.ToString();
            }
            catch
            {
            }
            return null;
        }

        void BuildDic(Guid listId, string xml)
        {
            Xml = xml;
            parsedXml = null;
            cache.Add(listId, ParsedXml.Element("Fields").Elements().Where(f => f.Attribute("ColName") != null).ToDictionary(f => f.Attribute("Name").Value, f => f.Attribute("ColName").Value));
        }

        void EnsureMapMembers(object map)
        {
            if (fieldNames == null)
            {
                Type mapType = map.GetType();
                fieldNames = mapType.GetField("m_scFieldNames", BindingFlags.NonPublic | BindingFlags.Instance);
                GetColNumber = mapType.GetMethod("GetColumnNumber", BindingFlags.Public | BindingFlags.Instance);
            }
        }
    }

    public class BatchXml : XmlDocument
    {
        public XmlElement RootElement
        {
            get;
            private set;
        }

        public int ItemCount
        {
            get
            {
                return RootElement != null ? RootElement.ChildNodes.Count : 0;
            }
        }

        public BatchXml()
        {
            var root = CreateElement("Batch");
            root.Attributes.Append(GetAttribute("OnError", "Continue"));
            root.Attributes.Append(GetAttribute("ListVersion", "1"));
            AppendChild(root);
            RootElement = root;
        }
        int methodIdInt = 1;
        public void DeleteItem(int id)
        {
            XmlElement method = CreateElement("Method");
            method.AppendAttributes(GetAttribute("ID", methodIdInt++), GetAttribute("Cmd", "Delete"));
            XmlElement field = CreateElement("Field");
            field.InnerText = id.ToString();
            field.AppendAttributes(GetAttribute("Name", "ID"));
            method.AppendChild(field);
            RootElement.AppendChild(method);
        }

        public void SubmitItem(Dictionary<string, object> values)
        {
            XmlElement method = CreateElement("Method");
            int o;
            method.AppendAttributes(GetAttribute("ID", methodIdInt++), GetAttribute("Cmd", values.ContainsKey("ID") && Int32.TryParse(values["ID"].ToString(), out o) ? "Update" : "New"));
            foreach (var pair in values)
            {
                XmlElement field = CreateElement("Field");
                field.AppendAttributes(GetAttribute("Name", pair.Key));
                field.InnerText = pair.Value != null ? pair.Value.ToString() : String.Empty;
                method.AppendChild(field);
            }
            RootElement.AppendChild(method);
        }



        XmlAttribute GetAttribute(string name, object value)
        {
            var attribute = CreateAttribute(name);
            attribute.Value = value != null ? value.ToString() : null;
            return attribute;
        }
    }


    [AttributeUsage(AttributeTargets.Property)]
    public class NonSerializable : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Class)]
    public class SerializeMethod : Attribute
    {
        string name;
        public string Name
        {
            get { return name; }
            private set { name = value; }
        }

        public SerializeMethod() : this("Serialize") { }
        public SerializeMethod(string methodName) { name = methodName; }
    }

    /*public partial class BaselineMethods
    {
        /// <summary>
        /// This super-function handles the "publishing" of a InfoPath form template onto a SharePoint list.
        /// </summary>
        /// <param name="web">SPWeb object</param>
        /// <param name="list">List that needs to be customized with the InfoPath template</param>
        /// <param name="formTemplateFileName">The InfoPath template's file name</param>
        /// <param name="templateData">The InfoPath template file as a byte array</param>
        /// <returns></returns>
        public bool PublishListFormTemplate(SPWeb web, SPList list, string formTemplateFileName, Byte[] templateData)
        {
            bool result = true;
            string tempFolderPath = null;
            try
            {
                SPFile newTemplate;

                tempFolderPath = MTMUtility.CreateTempFolder();
                string templateLocalPath = Path.Combine(tempFolderPath, formTemplateFileName);
                // this will overwrite an already existing file
                using (FileStream outStream = new FileStream(templateLocalPath, FileMode.Create, FileAccess.ReadWrite))
                {
                    outStream.Write(templateData, 0, templateData.Length);
                }

                // modify the form .xsn template to fit the list, data connetions, links, etc.
                if (ProcessListFormTemplate(web, list, templateLocalPath) == false) return false;

                byte[] hash;
                using (FileStream stream = IOFile.Open(templateLocalPath, FileMode.Open))
                {
                    System.Security.Cryptography.HashAlgorithm hashAlgorithm = new System.Security.Cryptography.SHA256CryptoServiceProvider();
                    hash = hashAlgorithm.ComputeHash(stream);

                    newTemplate = web.Files.Add(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/" + formTemplateFileName, stream, true);
                }

                /////////////////////// PROPERTY BAG OF THE INFOPATH TEMPLATE FILE ==============================

                if (newTemplate.Properties.ContainsKey("ipfs_listform")) newTemplate.Properties["ipfs_listform"] = "true";
                else newTemplate.Properties.Add("ipfs_listform", "true");

                if (newTemplate.Properties.ContainsKey("ipfs_streamhash")) newTemplate.Properties["ipfs_streamhash"] = Convert.ToBase64String(hash);
                else newTemplate.Properties.Add("ipfs_streamhash", Convert.ToBase64String(hash));

                newTemplate.Update();
                /////////////////////////////////////////////////////////////////////////////////////////////////



                /*
                SPFarm localFarm = SPFarm.Local;
                FormsService localFormsService = localFarm.Services.GetValue<FormsService>(FormsService.ServiceName);


                if (localFormsService.IsUserFormTemplateBrowserEnabled(newTemplate) == true)
                {
                    LogInfo("AddInfoPathTemplate :: YES, form is browser enabled: " + newTemplate.Name);

                }
                else
                {
                    LogInfo("AddInfoPathTemplate :: form is NOT browser enabled: " + newTemplate.Name);

                    ConverterMessageCollection response = localFormsService.BrowserEnableUserFormTemplate(newTemplate);

                    StreamWriter sw = new StreamWriter("C:\\TEMP\\INFOPATH_IFS.log", true);

                    foreach (ConverterMessage convMessage in response)
                    {
                        sw.WriteLine(convMessage.ShortMessage.ToString() + ": " + convMessage.DetailedMessage.ToString());
                    }

                    sw.Flush();
                    sw.Close();

                    if (localFormsService.IsUserFormTemplateBrowserEnabled(newTemplate))
                        LogInfo("AddInfoPathTemplate :: forms service says that it broswer-enabled the form: " + newTemplate.Name);
                }

            }
            catch (Exception ex)
            {

                return false;
            }
            finally
            {
                // clean up used resources
                if (Directory.Exists(tempFolderPath))
                {
                    MTMUtility.DeleteFileSystemInfo(new DirectoryInfo(tempFolderPath));
                }
            }

            return result;
        }


        /// <summary>
        /// This function extracts the InfoPath template's source files,
        /// and modifies various references in the 'manifest.xsf' file to match the objects (list IDs/URLs/content-type IDs) of the current SharePoint site.
        /// </summary>
        /// <param name="web">SPWeb object</param>
        /// <param name="targetList">List that needs to be customized with the InfoPath template</param>
        /// <param name="templateLocalPath">Local absolute path of the .xsn template</param>
        /// <returns></returns>
        private bool ProcessListFormTemplate(SPWeb web, SPList targetList, string templateLocalPath)
        {
            try
            {
                // create a instance of Microsoft.Deployment.Compression.Cab.CabInfo
                // which provides file-based operations on the cabinet file
                CabInfo cab = new CabInfo(templateLocalPath);

                string extractFolderPath = Path.GetDirectoryName(templateLocalPath) + @"\TEMPLATE_Extracted";
                // unpack the .xsn template
                cab.Unpack(extractFolderPath);

                // create XML document to load and process the form data
                XmlDocument templateXml = new XmlDocument();
                templateXml.PreserveWhitespace = true;

                using (FileStream stream = IOFile.Open(extractFolderPath + @"\manifest.xsf", FileMode.Open))
                {
                    templateXml.Load(stream);
                }

                // set up namespace manager for xPath
                XmlNamespaceManager ns = new XmlNamespaceManager(templateXml.NameTable);
                ns.AddNamespace("xsf", templateXml.DocumentElement.GetNamespaceOfPrefix("xsf"));
                ns.AddNamespace("xsf2", templateXml.DocumentElement.GetNamespaceOfPrefix("xsf2"));
                ns.AddNamespace("xsf3", templateXml.DocumentElement.GetNamespaceOfPrefix("xsf3"));

                XmlNodeList listAdapterNodes = templateXml.SelectNodes("//xsf:sharepointListAdapterRW", ns);
                foreach (XmlNode node in listAdapterNodes)
                {
                    if (node.Attributes["relativeListUrl"] == null)
                    {
                        return false;
                    }

                    // this is the main list adapter
                    if (node.Attributes["submitAllowed"].Value == "yes")
                    {
                        node.Attributes["sharePointListID"].Value = targetList.ID.ToString("B").ToUpper();
                        if (node.Attributes["contentTypeID"] != null)
                        {
                            node.Attributes["contentTypeID"].Value = targetList.ContentTypes.BestMatch(SPBuiltInContentTypeId.Item).ToString();
                        }

                        // adjust URL values, unless they are indicated as relative to current list location
                        if (node.Attributes["siteURL"].Value.Contains("../") == false)
                        {
                            node.Attributes["siteURL"].Value = web.ServerRelativeUrl;
                            node.Attributes["relativeListUrl"].Value = MTMUtility.UrlEncode(targetList.RootFolder.ServerRelativeUrl);
                        }
                    }
                    // this is a list data connection adapter
                    else
                    {
                        SPFolder listRootFolder;
                        string listURL = node.Attributes["relativeListUrl"].Value.TrimEnd('/');
                        // this is a list connection onto the list itself
                        if (listURL == "..")
                        {
                            listRootFolder = targetList.RootFolder;
                        }
                        else
                        {
                            listURL = StringFactory.SubstringAfter(listURL, "/", false);

                            listRootFolder = web.GetFolder(web.ServerRelativeUrl.TrimEnd('/') + "/Lists/" + listURL);
                            if (listRootFolder.Exists == false)
                            {
                                // this could be a link to a document library
                                listRootFolder = web.GetFolder(web.ServerRelativeUrl.TrimEnd('/') + "/" + listURL);
                                if (listRootFolder.Exists == false && web.IsRootWeb == false)
                                {
                                    SPWeb rootWeb = web.Site.RootWeb;
                                    listRootFolder = rootWeb.GetFolder(rootWeb.ServerRelativeUrl.TrimEnd('/') + "/Lists/" + listURL);
                                    if (listRootFolder.Exists == false)
                                    {
                                        // this could be a link to a document library
                                        listRootFolder = rootWeb.GetFolder(rootWeb.ServerRelativeUrl.TrimEnd('/') + "/" + listURL);
                                    }
                                }
                            }

                            // special case: data connections to product/study specific lists like "Product A - Health Authority"
                            if (listRootFolder.Exists == false && listURL.Contains('-'))
                            {
                                string partURL = MTMUtility.UrlEncode("-" + StringFactory.SubstringAfter(listURL, "-", false));

                                SPList list = web.Lists.Cast<SPList>().FirstOrDefault(l => MTMUtility.UrlEncode(l.RootFolder.Url).EndsWith(partURL, StringComparison.OrdinalIgnoreCase));
                                if (list != null) listRootFolder = list.RootFolder;
                            }


                            if (listRootFolder.Exists == false)
                            {

                                return false;
                            }
                        }

                        node.Attributes["sharePointListID"].Value = listRootFolder.ParentListId.ToString("B").ToUpper();

                        // when this attribute exists and has a value, reset it to match the "Item" content type of this particular list
                        if (node.Attributes["contentTypeID"] != null && !String.IsNullOrEmpty(node.Attributes["contentTypeID"].Value))
                        {
                            SPList list = web.Lists[listRootFolder.ParentListId];
                            node.Attributes["contentTypeID"].Value = list.ContentTypes.BestMatch(SPBuiltInContentTypeId.Item).ToString();
                        }

                        // adjust URL values, unless they are indicated as relative to current list location
                        if (node.Attributes["siteURL"].Value.Contains("../") == false)
                        {
                            // adding a trailing slash, so that when users choose to customize the template manually,
                            // modifying the data connections will not lose the list binding
                            node.Attributes["siteURL"].Value = web.ServerRelativeUrl.TrimEnd('/') + '/';
                            node.Attributes["relativeListUrl"].Value = MTMUtility.UrlEncode(listRootFolder.ServerRelativeUrl);
                        }
                    }
                }

                // reset the 'siteCollection' attribute on all data connections
                XmlNodeList dataConnectionNodes = templateXml.SelectNodes("//xsf2:connectoid", ns);
                foreach (XmlNode node in dataConnectionNodes)
                {
                    if (node.Attributes["siteCollection"] != null)
                    {
                        node.Attributes["siteCollection"].Value = web.Url.TrimEnd('/') + '/';
                    }
                }

                // removing this attribute (if it exists) will indicate that the form is published
                XmlNode docClassNode = templateXml.SelectSingleNode("/xsf:xDocumentClass", ns);
                if (docClassNode.Attributes["publishUrl"] != null)
                    docClassNode.Attributes.Remove(docClassNode.Attributes["publishUrl"]);

                // reset various other attributes
                XmlNode specificNode = MTMUtility.GetNodeWithAttribute(templateXml.DocumentElement, "runtimeCompatibilityURL");
                if (specificNode != null)
                    specificNode.Attributes["runtimeCompatibilityURL"].Value = "../../../_vti_bin/FormsServices.asmx";

                specificNode = MTMUtility.GetNodeWithAttribute(templateXml.DocumentElement, "originalPublishUrl");
                if (specificNode != null)
                    specificNode.Attributes["originalPublishUrl"].Value = "../../../";

                specificNode = MTMUtility.GetNodeWithAttribute(templateXml.DocumentElement, "path");
                if (specificNode != null)
                    specificNode.Attributes["path"].Value = "../../../";

                specificNode = MTMUtility.GetNodeWithAttribute(templateXml.DocumentElement, "relativeUrlBase");
                if (specificNode != null)
                {
                    specificNode.Attributes["relativeUrlBase"].Value = MTMUtility.UrlEncode(web.Url.TrimEnd('/') + '/' + targetList.RootFolder.Url.TrimEnd('/') + "/Item/");
                }
                else
                {
                    XmlNode parentNode = templateXml.DocumentElement.SelectSingleNode("/xsf:xDocumentClass/xsf:extensions/xsf:extension/xsf3:solutionDefinition", ns);

                    XmlNode child = templateXml.CreateNode(XmlNodeType.Element, "xsf3", "baseUrl", templateXml.DocumentElement.GetNamespaceOfPrefix("xsf3"));
                    child.InnerText = String.Empty;

                    XmlAttribute relBaseAttribute = templateXml.CreateAttribute("relativeUrlBase");
                    relBaseAttribute.Value = MTMUtility.UrlEncode(web.Url.TrimEnd('/') + '/' + targetList.RootFolder.Url.TrimEnd('/') + "/Item/");

                    child.Attributes.Append(relBaseAttribute);
                    parentNode.AppendChild(child);
                }

                // save the updated manifest file
                templateXml.Save(extractFolderPath + @"\manifest.xsf");

                // update static default value in the 'template.xml' instance
                ProcessListFormDefaultValues(targetList, extractFolderPath);

                // update static hyperlinks that may exist in the template's views
                ProcessListFormHyperlinks(web, extractFolderPath);

                // re-compress and overwrite the original .xsn file
                cab.Pack(extractFolderPath);

                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }


        private bool ProcessListFormDefaultValues(SPList targetList, string extractFolderPath)
        {
            bool saveRequired = false;
            try
            {
                // create XML document to load and process the form data
                XmlDocument templateXml = new XmlDocument();
                templateXml.PreserveWhitespace = true;

                using (FileStream stream = IOFile.Open(extractFolderPath + @"\template.xml", FileMode.Open))
                {
                    templateXml.Load(stream);
                }

                // set up namespace manager for xPath
                XmlNamespaceManager ns = new XmlNamespaceManager(templateXml.NameTable);
                ns.AddNamespace("dfs", templateXml.DocumentElement.GetNamespaceOfPrefix("dfs"));
                ns.AddNamespace("q", templateXml.DocumentElement.GetNamespaceOfPrefix("q"));
                ns.AddNamespace("my", templateXml.DocumentElement.GetNamespaceOfPrefix("my"));

                List<string> internalColumnNames = new List<string>();
                XmlNode queryFields = templateXml.SelectSingleNode("/dfs:myFields/dfs:queryFields/q:SharePointListItem_RW", ns);
                // get all the child nodes (excluding "#whitespace" "nodes")
                foreach (XmlNode childNode in queryFields.SelectNodes("*"))
                {
                    internalColumnNames.Add(childNode.LocalName);
                }

                foreach (string internalName in internalColumnNames)
                {
                    SPField field = targetList.Fields.Cast<SPField>().FirstOrDefault(f => f.InternalName == internalName);
                    if (field == null) continue;

                    if (field.Type == SPFieldType.Text && !String.IsNullOrEmpty(field.DefaultValue))
                    {
                        XmlNode targetNode = templateXml.SelectSingleNode("//my:" + internalName, ns);
                        if (targetNode != null)
                        {
                            targetNode.InnerText = field.DefaultValue;
                            saveRequired = true;

                        }
                    }
                }

                if (saveRequired)
                {
                    // save the updated instance file
                    templateXml.Save(extractFolderPath + @"\template.xml");
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        /// <summary>
        /// This function searches for static hyperlinks in the InfoPath template views,
        /// and updates these for the current SharePoint site.
        /// </summary>
        /// <param name="web">SPWeb object</param>
        /// <param name="tempFolderPath">Local direcory path of extracted InfoPath source files</param>
        /// <returns></returns>
        private bool ProcessListFormHyperlinks(SPWeb web, string tempFolderPath)
        {
            bool result = true;

            try
            {
                System.IO.FileInfo[] files = null;
                System.IO.DirectoryInfo tempFolder = new DirectoryInfo(tempFolderPath);
                // hyperlinks are found in the stylesheet files responsible for view rendering
                files = tempFolder.GetFiles("*.xsl");

                foreach (System.IO.FileInfo file in files)
                {
                    // create XML document to load and process the form data
                    XmlDocument templateXml = new XmlDocument();
                    templateXml.PreserveWhitespace = true;

                    using (FileStream stream = file.Open(FileMode.Open))
                    {
                        templateXml.Load(stream);
                    }

                    XmlNodeList linkNodes = templateXml.SelectNodes("//*[@href]");
                    foreach (XmlNode node in linkNodes)
                    {
                        string linkUrl = node.Attributes["href"].Value;

                        try
                        {
                            Uri uri = new Uri(linkUrl);
                            string absolutePath = uri.AbsolutePath.TrimStart('/');
                            string queryString = linkUrl.Substring(uri.GetLeftPart(UriPartial.Path).Length);

                            string aspxPage = String.Empty;
                            if (absolutePath.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                            {
                                aspxPage = StringFactory.SubstringAfter(absolutePath, "/", false);
                                absolutePath = absolutePath.Replace('/' + aspxPage, "");
                                if (absolutePath.IndexOf("/Lists/", StringComparison.OrdinalIgnoreCase) == -1)
                                    absolutePath = StringFactory.SubstringAfter(absolutePath, "/", true);
                                else
                                    absolutePath = "Lists/" + StringFactory.SubstringAfter(absolutePath, "/", true);
                            }

                            while (true)
                            {
                                SPFolder folder = web.GetFolder(web.ServerRelativeUrl + '/' + absolutePath);

                                if (folder.Exists)
                                {
                                    linkUrl = web.Url.TrimEnd('/') + '/' + folder.Url.TrimEnd('/') + '/' + aspxPage + queryString;
                                    node.Attributes["href"].Value = MTMUtility.UrlEncode(linkUrl);
                                    break;
                                }

                                SPWeb rootWeb = web.Site.RootWeb;
                                folder = rootWeb.GetFolder(rootWeb.ServerRelativeUrl + '/' + absolutePath);
                                if (folder.Exists)
                                {
                                    linkUrl = rootWeb.Url.TrimEnd('/') + '/' + folder.Url.TrimEnd('/') + '/' + aspxPage + queryString;
                                    node.Attributes["href"].Value = MTMUtility.UrlEncode(linkUrl);
                                    break;
                                }

                                int index = absolutePath.IndexOf('/');
                                if (index != -1)
                                    absolutePath = absolutePath.Substring(index + 1);
                                else
                                {

                                    result = false;
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                            result = false;
                        }

                    }

                    templateXml.Save(file.FullName);
                }
            }
            catch (Exception ex)
            {

                return false;
            }

            return result;
        }


        /// <summary>
        /// Creates the NewForm, EditForm and DisplayForm aspx pages used by the list InfoPath forms
        /// </summary>
        /// <param name="web"></param>
        /// <param name="list"></param>
        /// <param name="formTemplateFileName"></param>
        /// <returns></returns>
        public bool ProcessASPXforms(SPWeb web, SPList list, string formTemplateFileName)
        {
            bool result = true;
            try
            {
                SPFile aspxSource;
                SPFile aspxPage;

                aspxSource = web.GetFile(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/DispForm.aspx");
                aspxSource.CopyTo(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/displayifs.aspx", true);
                aspxPage = web.GetFile(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/displayifs.aspx");
                AddBrowserFormWebPartToListPage(web, list, aspxPage, PAGETYPE.PAGE_DISPLAYFORM, formTemplateFileName);

                aspxSource = web.GetFile(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/EditForm.aspx");
                aspxSource.CopyTo(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/editifs.aspx", true);
                aspxPage = web.GetFile(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/editifs.aspx");
                AddBrowserFormWebPartToListPage(web, list, aspxPage, PAGETYPE.PAGE_EDITFORM, formTemplateFileName);

                aspxSource = web.GetFile(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/NewForm.aspx");
                aspxSource.CopyTo(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/newifs.aspx", true);
                aspxPage = web.GetFile(list.RootFolder.ServerRelativeUrl.TrimEnd('/') + "/Item/newifs.aspx");
                AddBrowserFormWebPartToListPage(web, list, aspxPage, PAGETYPE.PAGE_NEWFORM, formTemplateFileName);

                SPContentType itemContentType = list.ContentTypes[list.ContentTypes.BestMatch(SPBuiltInContentTypeId.Item)];
                itemContentType.DisplayFormUrl = "~list/Item/displayifs.aspx";
                itemContentType.NewFormUrl = "~list/Item/newifs.aspx";
                itemContentType.EditFormUrl = "~list/Item/editifs.aspx";

                itemContentType.Update();
            }
            catch (Exception ex)
            {
                return false;
            }
            return result;
        }

        /// <summary>
        /// Adds a Browser Form Web Part to an indicated aspx page
        /// </summary>
        /// <param name="web"></param>
        /// <param name="list"></param>
        /// <param name="aspxPage"></param>
        /// <param name="pageType"></param>
        /// <param name="formTemplateFileName"></param>
        private static void AddBrowserFormWebPartToListPage(SPWeb web, SPList list, SPFile aspxPage, PAGETYPE pageType, string formTemplateFileName)
        {
            using (SPLimitedWebPartManager webPartManager = aspxPage.GetLimitedWebPartManager(System.Web.UI.WebControls.WebParts.PersonalizationScope.Shared))
            {
                List<AspWebPart> partsToDelete = new List<AspWebPart>();
                foreach (AspWebPart wPart in webPartManager.WebParts) partsToDelete.Add(wPart);
                foreach (AspWebPart wPart in partsToDelete) webPartManager.DeleteWebPart(wPart);

                // instantiate a new browser-form web part object
                BrowserFormWebPart webPart = MTMUtility.GetWebPart(web.Site, "Microsoft.Office.InfoPath.Server.BrowserForm.webpart") as BrowserFormWebPart;

                webPart.Title = "MTM List Form";
                webPart.FormLocation = "~list/Item/" + formTemplateFileName;
                webPart.ContentTypeId = list.ContentTypes.BestMatch(SPBuiltInContentTypeId.Item).ToString();

                webPart.SubmitBehavior = SubmitBehavior.FormDefault;
                webPart.ChromeType = System.Web.UI.WebControls.WebParts.PartChromeType.None;

                // "Editable" for newifs.aspx and editifs.aspx and "ReadOnly" for displayifs.aspx
                webPart.ListFormMode = (pageType == PAGETYPE.PAGE_DISPLAYFORM) ? ListFormMode.ReadOnly : ListFormMode.Editable;

                // set the IListWebPart properties required for the form
                IListWebPart part = webPart as IListWebPart;
                part.ListId = list.ID;
                part.PageType = pageType;

                webPartManager.AddWebPart(webPart, "Main", 0);
                webPartManager.SaveChanges(webPart);
            }
        }
    }*/

    public class Serialization
    {
        Type type;
        IEnumerable<MethodInfo> getMethods;
        object obj;
        public Serialization(object obj)
        {
            this.obj = obj;
            if (obj == null)
            {
                serialization = "null";
                return;
            }
            type = obj.GetType();
            getMethods = from p in type.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                         where p.GetCustomAttribute<NonSerializable>() == null
                         select p.GetMethod;
        }

        public static implicit operator string(Serialization serialization)
        {
            return serialization != null ? serialization.ToString() : null;
        }

        string TrySerializeDataEntity()
        {
            if (!type.FullName.StartsWith("System.Data.Data") || type.FullName.EndsWith("Collection"))
            {
                return null;
            }
            if (obj is DataRow)
            {
                DataRow row = obj as DataRow;
                StringBuilder dicBuilder = new StringBuilder("{");
                foreach (DataColumn column in row.Table.Columns)
                {
                    dicBuilder.AppendFormat("\"{0}\":{1},", column.ColumnName, new Serialization(row[column.ColumnName]));
                }
                dicBuilder.EndSerialization();
                return dicBuilder.ToString();
            }
            else if (obj is DataTable)
            {
                return new Serialization((obj as DataTable).Rows);
            }
            return null;
        }

        public static string SerializeType(Type type)
        {
            if (type.IsEnum)
            {
                return String.Format("{{{0}}}", String.Join(",", Enum.GetNames(type).Select(n => String.Format("\"{0}\":{1}", n, (int)Enum.Parse(type, n)))));
            }
            StringBuilder typeSerializer = new StringBuilder("{");
            BindingFlags allFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.Static | BindingFlags.NonPublic;
            typeSerializer.AppendFormat("\"{0}\":{1},", "Name", new Serialization(type.Name));
            typeSerializer.AppendFormat("\"{0}\":{1},", "FullName", new Serialization(type.FullName));
            typeSerializer.AppendFormat("\"{0}\":{1},", "QualifiedName", new Serialization(type.AssemblyQualifiedName));
            typeSerializer.AppendFormat("\"{0}\":{1},", "Properties", new Serialization(type.GetProperties(allFlags).Select(p => new Dictionary<string, string>
                {
                    { "Name", p.Name },
                    { "PropertyType", p.PropertyType.FullName },
                })));
            typeSerializer.AppendFormat("\"{0}\":{1},", "Methods", new Serialization(type.GetMethods(allFlags).Select(p => new Dictionary<string, string>
                {
                    { "Name", p.Name },
                    { "ReturnType", p.ReturnType.FullName },
                })));
            typeSerializer.AppendFormat("\"{0}\":{1}", "Fields", new Serialization(type.GetFields(allFlags).Select(p => new Dictionary<string, string>
                {
                    { "Name", p.Name },
                    { "FieldType", p.FieldType.FullName },
                })));
            typeSerializer.Append("}");
            return typeSerializer.ToString();
        }

        string serialization;
        public override string ToString()
        {
            if (serialization != null) { return serialization; }
            switch (type.FullName)
            {
                case "System.Object":
                    return "{}";
                case "System.DBNull":
                    return new Serialization(null);
                case "System.Boolean":
                case "System.Double":
                case "System.Int32":
                    return serialization = obj.ToString().ToLower();
                case "System.Char":
                    return serialization = "'" + obj.ToString() + "'";
                case "System.DateTime":
                    return serialization = ((DateTime)obj) != default(DateTime) ? obj.ToString() : null;
                case "System.String":
                case "System.Guid":
                    return serialization = obj != null ? String.Format("\"{0}\"", obj.ToString().Replace("\"", "\\\"")) : "null";
            }
            if (obj is Type)
            {
                return SerializeType(obj as Type);
            }
            serialization = TrySerializeDataEntity();
            if (serialization != null) { return serialization; }
            if (type.IsEnum)
            {
                return serialization = (obj is string ? (int)Enum.Parse(type, obj.ToString()) : (int)obj).ToString();
            }
            else if (obj is IDictionary)
            {
                var dic = obj as IDictionary;
                var e = dic.Keys.GetEnumerator();
                StringBuilder dicBuilder = new StringBuilder("{");
                while (e.MoveNext())
                {
                    dicBuilder.AppendFormat("{0}:{1},", new Serialization(e.Current.ToString()), new Serialization(dic[e.Current]));
                }
                dicBuilder.EndSerialization();
                return serialization = dicBuilder.ToString();
            }
            else if (obj is IEnumerable)
            {
                var e = (obj as IEnumerable).GetEnumerator();
                StringBuilder arrayBuilder = new StringBuilder("[");
                while (e.MoveNext())
                {
                    arrayBuilder.AppendFormat("{0},", new Serialization(e.Current));
                }
                arrayBuilder.EndSerialization();
                return serialization = arrayBuilder.ToString();
            }
            var serializeMethodAttribute = type.GetCustomAttribute<SerializeMethod>();
            if (serializeMethodAttribute != null)
            {
                return type.GetMethod(serializeMethodAttribute.Name).Invoke(obj, new object[0]) as string;
            }
            StringBuilder builder = new StringBuilder("{");
            foreach (MethodInfo getMethod in getMethods)
            {
                if (getMethod.GetParameters().Length > 0) { continue; }
                string name = getMethod.Name.Substring(getMethod.Name.IndexOf("_") + 1);
                object value = getMethod.Invoke(obj, new object[0]);
                if (value == null)
                {
                    builder.AppendFormat("\"{0}\":{1},", name, "null");
                    continue;
                }
                string strValue = value.ToString();
                switch (getMethod.ReturnType.FullName)
                {
                    case "System.Boolean":
                    case "System.Double":
                    case "System.Int32":
                        builder.AppendFormat("\"{0}\":{1},", name, strValue.ToLower());
                        break;
                    case "System.DateTime":
                        builder.AppendFormat("\"{0}\":{1},", name, ((DateTime)value) != default(DateTime) ? new Serialization(strValue) : "null");
                        break;
                    case "System.String":
                        builder.AppendFormat("\"{0}\":\"{1}\",", name, strValue.Replace("\"", "\\\""));
                        break;
                    default:
                        builder.AppendFormat("\"{0}\":{1},", name, new Serialization(value));
                        break;
                }
            }
            builder.EndSerialization();
            return serialization = builder.ToString();
        }
    }

}
