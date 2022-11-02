using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Web;
using System.Text;

namespace DKOutlookAddIn
{
    [ComVisible(true)]
    public interface IAddInUtilities
    {
        string GetAppointmentArray();
        void SetAppointmentArray(string json);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        ThisAddIn addIn;
        public void SetAddIn(ThisAddIn a) => addIn = a;

        public string GetAppointmentArray()
        {
            Outlook.MAPIFolder calendar = addIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            Outlook.Items calendarItems = calendar.Items;

            JObject jRoot = new JObject();
            JArray jArray = new JArray();

            Outlook.AppointmentItem item = calendarItems.GetFirst() as Outlook.AppointmentItem;

            while (item != null)
            {
                JObject obj = new JObject();
                //obj["EntryID"] = item.EntryID;
                obj["Importance"] = item.Importance.ToString();
                obj["Subject"] = item.Subject;
                obj["Location"] = item.Location;
                obj["Body"] = item.Body;
                obj["Start"] = item.Start.ToString();
                obj["End"] = item.End.ToString();
                obj["AllDayEvent"] = item.AllDayEvent;

                jArray.Add(obj);

                item = calendarItems.GetNext();
            }

            jRoot["appointmentArray"] = jArray;
            return jRoot.ToString();
        }

        public void SetAppointmentArray(string json)
        {
            Outlook.MAPIFolder calendar = addIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            Outlook.Items calendarItems = calendar.Items;

            try
            {
                JObject jRoot = JObject.Parse(json);
                JArray jArray = jRoot["appointmentArray"].ToObject<JArray>();
                foreach (var item in jArray)
                {
                    JObject obj = (JObject)item;
                    string Subject = obj["Subject"].ToString();

                    // Outlook 不支持通过EntryID查找？
                    Outlook.AppointmentItem matchItem = calendarItems.Find("[Subject] = '" + Subject + "'");
                    if (matchItem != null)
                    {
                        if (obj["Start"] != null)       matchItem.Start = DateTime.Parse(obj["Start"].ToString());
                        if (obj["End"] != null)         matchItem.End = DateTime.Parse(obj["End"].ToString());
                        if (obj["Location"] != null)    matchItem.Location = UrlDecode(obj["Location"].ToString(), Encoding.UTF8);
                        if (obj["Body"] != null)        matchItem.Body = UrlDecode(obj["Body"].ToString(), Encoding.UTF8);
                        if (obj["Subject"] != null)     matchItem.Subject = UrlDecode(obj["Subject"].ToString(), Encoding.UTF8);
                        if (obj["AllDayEvent"] != null) matchItem.AllDayEvent = obj["AllDayEvent"].ToObject<bool>();
                        matchItem.Save();
                        //matchItem.Display(true);
                    }
                    else
                    {
                        Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)addIn.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                        if (obj["Start"] != null)       newAppointment.Start = DateTime.Parse(obj["Start"].ToString());
                        if (obj["End"] != null)         newAppointment.End = DateTime.Parse(obj["End"].ToString());
                        if (obj["Location"] != null)    newAppointment.Location = UrlDecode(obj["Location"].ToString(), Encoding.UTF8);
                        if (obj["Body"] != null)        newAppointment.Body = UrlDecode(obj["Body"].ToString(), Encoding.UTF8);
                        if (obj["Subject"] != null)     newAppointment.Subject = UrlDecode(obj["Subject"].ToString(), Encoding.UTF8);
                        if (obj["AllDayEvent"] != null) newAppointment.AllDayEvent = obj["AllDayEvent"].ToObject<bool>();
                        newAppointment.Save();
                        //newAppointment.Display(true);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("err: " + ex.ToString());
            }
        }
        public static string UrlEncode(string text, Encoding encod, bool cap = true)
        {
            if (cap)
            {
                StringBuilder builder = new StringBuilder();
                foreach (char c in text)
                {
                    if (System.Web.HttpUtility.UrlEncode(c.ToString(), encod).Length > 1)
                    {
                        builder.Append(System.Web.HttpUtility.UrlEncode(c.ToString(), encod).ToUpper());
                    }
                    else
                    {
                        builder.Append(c);
                    }
                }
                return builder.ToString();
            }
            else
            {
                string encodString = System.Web.HttpUtility.UrlEncode(text, encod);
                return encodString;
            }
        }

        public static string UrlDecode(string encodString, Encoding encod)
        {
            string text = System.Web.HttpUtility.UrlDecode(encodString, encod);
            return text;
        }
    }
}
