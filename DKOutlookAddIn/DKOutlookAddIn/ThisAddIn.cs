using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace DKOutlookAddIn
{
    public partial class ThisAddIn
    {
        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern IntPtr SendMessage(int hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern int FindWindow(string lpClassName, string lpWindowName);
        public const int UWM_MESSAGE_NOTIFY_OUTLOOKCHANGED = 0x0400 + 0x2201;
        public const int UWM_MESSAGE_NOTIFY_OUTLOOKREMOVED = 0x0400 + 0x2202;

        private AddInUtilities utilities;

        protected override object RequestComAddInAutomationService()
        {
            if (utilities == null)
                utilities = new AddInUtilities();

            return utilities;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            utilities.SetAddIn(this);

            Outlook.MAPIFolder calendar = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);

            Outlook.Items calendarItems = calendar.Items;

            calendarItems.ItemAdd += OnEventItemChanged;
            calendarItems.ItemChange += OnEventItemChanged;
            calendarItems.ItemRemove += OnEventItemRemove;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        void OnEventItemChanged(object item)
        {
            int hWnd = FindWindow("DKPluginOutlookDelegateWnd", "DKPluginOutlookDelegateWndText");
            if (hWnd != 0)
            {
                SendMessage(hWnd, UWM_MESSAGE_NOTIFY_OUTLOOKCHANGED, new IntPtr(0), new IntPtr(0));
            }

        }
        void OnEventItemRemove()
        {
            int hWnd = FindWindow("DKPluginOutlookDelegateWnd", "DKPluginOutlookDelegateWndText");
            if (hWnd != 0)
            {
                SendMessage(hWnd, UWM_MESSAGE_NOTIFY_OUTLOOKREMOVED, new IntPtr(0), new IntPtr(0));
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
