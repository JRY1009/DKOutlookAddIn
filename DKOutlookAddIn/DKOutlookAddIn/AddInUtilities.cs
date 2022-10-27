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

                jArray.Add(obj);

                item = calendarItems.GetNext();
            }

            return jArray.ToString();
        }
    }
}
