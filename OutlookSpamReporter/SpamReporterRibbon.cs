using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using OutlookSpamReporter.Utilities;
using Microsoft.Win32;

namespace OutlookSpamReporter
{
    public partial class SpamReporterRibbon
    {
        string email = "";
        private void SpamReporterRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            string KeyPath = @"Software\Microsoft\Office\Outlook\Addins\OutlookSpamReporter";
            email = ReadRegistryValue(KeyPath, "AdminEmail");
        }

        private void btnReportSpam_Click(object sender, RibbonControlEventArgs e)
        {
            string KeyPath = @"Software\Microsoft\Office\Outlook\Addins\OutlookSpamReporter";
            string labelValue = ReadRegistryValue(KeyPath, "ButtonLabel");
            if (!string.IsNullOrEmpty(labelValue))
            {
                btnReportSpam.Label = labelValue;
            }
            else
            { btnReportSpam.Label = "Report Spam"; }
            try
            {
                var application = Globals.ThisAddIn.Application;
                FileLogger.Info("Report Spam clicked");
                Outlook.Selection selection = application.ActiveExplorer().Selection;
                if (selection == null || selection.Count == 0)
                {
                    FileLogger.Info("No items selected");
                    System.Windows.Forms.MessageBox.Show("Please select at least one email.");
                    return;
                }


                Outlook.MailItem forwardMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                forwardMail.Subject = "Spam Report";
                forwardMail.To = email;
                forwardMail.Body = "Please review the attached spam emails for investigation.";

                for (int i = 1; i <= selection.Count; i++)
                {
                    object item = selection[i];
                    var mail = item as Outlook.MailItem;
                    if (mail == null)
                    {
                        FileLogger.Info("Skipped non-mail selection item");
                        continue;
                    }

                    string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".msg");
                    mail.SaveAs(tempPath, Outlook.OlSaveAsType.olMSGUnicode);
                    forwardMail.Attachments.Add(tempPath, Outlook.OlAttachmentType.olByValue, Type.Missing, mail.Subject);
                    FileLogger.Info("Attached message: " + tempPath);
                }

                forwardMail.Send();
                FileLogger.Info("Spam report email sent automatically");
            }
            catch (Exception ex)
            {
                FileLogger.Error("Error creating spam report", ex);
            }
        }

        private void btnReportSpamRead_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                FileLogger.Info("Report Spam (Read) clicked");

                // Get the currently open mail item
                Outlook.Inspector inspector = application.ActiveInspector();
                if (inspector == null)
                {
                    FileLogger.Info("No mail item open");
                    System.Windows.Forms.MessageBox.Show("Please open an email to report as spam.");
                    return;
                }

                Outlook.MailItem currentMail = inspector.CurrentItem as Outlook.MailItem;
                if (currentMail == null)
                {
                    FileLogger.Info("Current item is not a mail item");
                    System.Windows.Forms.MessageBox.Show("Please open an email to report as spam.");
                    return;
                }

                Outlook.MailItem forwardMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                forwardMail.Subject = "Spam Report";
                forwardMail.To = email;
                forwardMail.Body = "Please review the attached spam email for investigation.";

                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".msg");
                currentMail.SaveAs(tempPath, Outlook.OlSaveAsType.olMSGUnicode);
                forwardMail.Attachments.Add(tempPath, Outlook.OlAttachmentType.olByValue, Type.Missing, currentMail.Subject);
                FileLogger.Info("Attached message: " + tempPath);

                forwardMail.Send();
                FileLogger.Info("Spam report email sent automatically");
            }
            catch (Exception ex)
            {
                FileLogger.Error("Error creating spam report", ex);
            }
        }
        private string ReadRegistryValue(string keyName, string valueName)
        { try 
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyName)) 
                { 
                    if (key != null) 
                    {
                        object value = key.GetValue(valueName); if (value != null) 
                        {
                            return value.ToString(); 
                        } 
                    }
                }
            } catch (Exception ex) 
            {
                System.Windows.Forms.MessageBox.Show($"Error reading registry: {ex.Message}");
            } return null; 
        }
    }
}
