using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace MeetingRequestHelper
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private string LOCATION_PATH = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MeetingRequestHelper\\Location.txt");
        private string TEMPLATE_PATH = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MeetingRequestHelper\\MeetingTemplate.rtf");

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MeetingRequestHelper.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void AttachICal(Office.IRibbonControl control)
        {
            try
            {
                var appointment = GetCurrentAppointmentItem();
                if (appointment != null)
                {
                    string filename = string.Format("{0}\\iCalendar.ics",
                        Environment.GetFolderPath(Environment.SpecialFolder.InternetCache));

                    //export appointment as iCal format
                    appointment.SaveAs(filename, Outlook.OlSaveAsType.olICal);

                    //then attach exported iCal at the end of the body
                    appointment.Attachments.Add(filename, Outlook.OlAttachmentType.olByValue,
                        (appointment.Body == null) ? 1 : appointment.Body.Length + 1);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void ApplyTemplate(Office.IRibbonControl control)
        {
            try
            {
                var appointment = GetCurrentAppointmentItem();
                if (appointment != null)
                {
                    if (File.Exists(LOCATION_PATH))
                        appointment.Location = File.ReadAllText(LOCATION_PATH);

                    if (File.Exists(TEMPLATE_PATH))
                        appointment.RTFBody = File.ReadAllBytes(TEMPLATE_PATH);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void SaveTemplate(Office.IRibbonControl control)
        {
            try
            {
                var appointment = GetCurrentAppointmentItem();
                if (appointment != null)
                {
                    //ensure our directory exists
                    (new FileInfo(LOCATION_PATH)).Directory.Create();

                    if (appointment.Location != null)
                        File.WriteAllText(LOCATION_PATH, appointment.Location);

                    if (appointment.RTFBody != null)
                        File.WriteAllBytes(TEMPLATE_PATH, appointment.RTFBody);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private Outlook.AppointmentItem GetCurrentAppointmentItem()
        {
            var outlook = new Outlook.Application();
            var inspector = outlook.ActiveInspector();
            return inspector.CurrentItem as Outlook.AppointmentItem;
        }

        #endregion
    }
}
