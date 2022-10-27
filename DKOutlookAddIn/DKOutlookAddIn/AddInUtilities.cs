using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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

            JArray jArray = new JArray();

            Outlook.AppointmentItem item = calendarItems.GetFirst() as Outlook.AppointmentItem;

            while (item != null)
            {
                JObject obj = new JObject();
                obj["EntryID"] = item.EntryID;
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

            return jArray.ToString();
        }

        public void SetAppointmentArray(string json)
        {
            Outlook.MAPIFolder calendar = addIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            Outlook.Items calendarItems = calendar.Items;

            try
            {
                JArray jArray = JArray.Parse(json);
                foreach (var item in jArray)
                {
                    JObject obj = (JObject)item;
                    string Subject = obj["Subject"].ToString();

                    // Outlook 不支持通过EntryID查找？
                    Outlook.AppointmentItem matchItem = calendarItems.Find("[Subject] = '" + Subject + "'");
                    if (matchItem != null)
                    {
                        matchItem.Start = DateTime.Parse(obj["Start"].ToString());
                        matchItem.End = DateTime.Parse(obj["End"].ToString());
                        matchItem.Location = obj["Location"].ToString();
                        matchItem.Body = obj["Body"].ToString();
                        matchItem.Subject = obj["Subject"].ToString();
                        matchItem.Save();
                        //matchItem.Display(true);
                    }
                    else
                    {
                        Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)addIn.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                        newAppointment.Start = DateTime.Parse(obj["Start"].ToString());
                        newAppointment.End = DateTime.Parse(obj["End"].ToString());
                        newAppointment.Location = obj["Location"].ToString();
                        newAppointment.Body = obj["Body"].ToString();
                        newAppointment.Subject = obj["Subject"].ToString();
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
    }
}
