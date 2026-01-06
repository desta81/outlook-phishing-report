using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using OutlookPhishingReporter.Utilities;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Net.Mail;

namespace OutlookPhishingReporter
{
    public partial class PhishingReporterRibbon
    {
        private const string REGISTRY_KEY_PATH = @"Software\Microsoft\Office\Outlook\Addins\OutlookPhishingReporter";
        private const string ADMIN_EMAIL_VALUE = "AdminEmail";
        private const string CUSTOM_ICON_VALUE = "CustomIconPath";
        private const string CUSTOM_LABEL_VALUE = "ButtonLabel";

        string email = "";
        string customIconPath = "";
        string customButtonLabel = "";

        private void PhishingReporterRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            email = ReadRegistryValue(REGISTRY_KEY_PATH, ADMIN_EMAIL_VALUE);
            customIconPath = ReadRegistryValue(REGISTRY_KEY_PATH, CUSTOM_ICON_VALUE);
            customButtonLabel = ReadRegistryValue(REGISTRY_KEY_PATH, CUSTOM_LABEL_VALUE);

            // If email not configured in registry, use default for debug/testing
            if (string.IsNullOrEmpty(email))
            {
                #if DEBUG
                    email = "yosi.desta@outlook.co.il"; // Default email for debug mode
                    FileLogger.Info("Using default email for debug mode: " + email);
                #else
                    FileLogger.Info("Admin email not configured in registry");
                #endif
            }
            else if (!IsValidEmail(email))
            {
                FileLogger.Info("Admin email configured but invalid format: " + email);
            }

            // Apply custom icon if configured
            if (!string.IsNullOrEmpty(customIconPath) && File.Exists(customIconPath))
            {
                try
                {
                    System.Drawing.Image customIcon = System.Drawing.Image.FromFile(customIconPath);
                    btnReportPhishing.Image = customIcon;
                    btnReportPhishingRead.Image = customIcon;
                    FileLogger.Info("Custom icon loaded from: " + customIconPath);
                }
                catch (Exception ex)
                {
                    FileLogger.Error("Failed to load custom icon from: " + customIconPath, ex);
                }
            }

            // Apply custom button label if configured
            if (!string.IsNullOrEmpty(customButtonLabel))
            {
                btnReportPhishing.Label = customButtonLabel;
                btnReportPhishingRead.Label = customButtonLabel;
                FileLogger.Info("Custom button label applied: " + customButtonLabel);
            }
        }

        private void btnReportPhishing_Click(object sender, RibbonControlEventArgs e)
        {
            List<string> tempFiles = new List<string>();
            Outlook.MailItem mailToDelete = null;

            try
            {
                var application = Globals.ThisAddIn.Application;
                FileLogger.Info("Report phishing clicked");

                Outlook.Selection selection = application.ActiveExplorer().Selection;
                if (selection == null || selection.Count == 0)
                {
                    FileLogger.Info("No items selected");
                    MessageBox.Show("Please select at least one email.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (selection.Count > 1)
                {
                    FileLogger.Info("Multiple items selected");
                    MessageBox.Show("Only one email at a time is allowed.", "Multiple Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Validate email address
                if (string.IsNullOrEmpty(email) || !IsValidEmail(email))
                {
                    FileLogger.Error("Invalid or missing admin email address: " + email, null);
                    MessageBox.Show("Configuration error: Invalid admin email address. Please contact your administrator.", "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Outlook.MailItem forwardMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                forwardMail.Subject = "Phish Report";
                forwardMail.To = email;
                forwardMail.Body = "Please review the attached phish email for investigation.";

                for (int i = 1; i <= selection.Count; i++)
                {
                    object item = selection[i];
                    var mail = item as Outlook.MailItem;
                    if (mail == null)
                    {
                        FileLogger.Info("Skipped non-mail selection item");
                        continue;
                    }

                    // Store reference to mail item for deletion
                    mailToDelete = mail;

                    string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".msg");
                    tempFiles.Add(tempPath);
                    mail.SaveAs(tempPath, Outlook.OlSaveAsType.olMSGUnicode);
                    forwardMail.Attachments.Add(tempPath, Outlook.OlAttachmentType.olByValue, Type.Missing, mail.Subject);
                    FileLogger.Info("Attached message: " + tempPath);
                }

                DialogResult confirmResult = MessageBox.Show(
                    "You are about to report this email as phishing. The email will be forwarded to your security team for review. If you believe this email is legitimate or was selected by mistake, click Cancel.",
                    "Report Phishing Email",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question);

                if (confirmResult == DialogResult.OK)
                {
                    forwardMail.Send();
                    FileLogger.Info("Phish report email sent successfully");

                    // Move reported email to Deleted Items
                    if (mailToDelete != null)
                    {
                        MoveToDeletedItems(mailToDelete);
                    }

                    // Updated success message with custom title and message text
                    MessageBox.Show(
                        "Thank you for reporting this email as phishing. It has been moved to the Deleted Items folder.",
                        "Report phishing",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    FileLogger.Info("User canceled sending phish report");
                    MessageBox.Show("Phishing report canceled.", "Canceled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("Error creating phish report", ex);
                MessageBox.Show("Failed to report the email. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up temporary files
                if (tempFiles != null && tempFiles.Count > 0)
                {
                    CleanupTempFiles(tempFiles);
                }
            }
        }

        private void btnReportPhishingRead_Click(object sender, RibbonControlEventArgs e)
        {
            string tempPath = null;
            Outlook.MailItem mailToDelete = null;

            try
            {
                var application = Globals.ThisAddIn.Application;
                FileLogger.Info("Report Phish (Read) clicked");

                // Get the currently open mail item
                Outlook.Inspector inspector = application.ActiveInspector();
                if (inspector == null)
                {
                    FileLogger.Info("No mail item open");
                    MessageBox.Show("Please open an email to report as phishing.", "No Email Open", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Outlook.MailItem currentMail = inspector.CurrentItem as Outlook.MailItem;
                if (currentMail == null)
                {
                    FileLogger.Info("Current item is not a mail item");
                    MessageBox.Show("Please open an email to report as phishing.", "Not an Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Store reference to mail item for deletion
                mailToDelete = currentMail;

                // Validate email address
                if (string.IsNullOrEmpty(email) || !IsValidEmail(email))
                {
                    FileLogger.Error("Invalid or missing admin email address: " + email, null);
                    MessageBox.Show("Configuration error: Invalid admin email address. Please contact your administrator.", "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Add confirmation dialog to Read mode
                DialogResult confirmResult = MessageBox.Show(
                    "You are about to report this email as phishing. The email will be forwarded to your security team for review. If you believe this email is legitimate or was selected by mistake, click Cancel.",
                    "Report Phishing Email",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question);

                if (confirmResult != DialogResult.OK)
                {
                    MessageBox.Show("Phishing report canceled.", "Canceled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    FileLogger.Info("Phish report canceled by user");
                    return;
                }

                Outlook.MailItem forwardMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                forwardMail.Subject = "Phish Report";
                forwardMail.To = email;
                forwardMail.Body = "Please review the attached phish email for investigation.";

                tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".msg");
                currentMail.SaveAs(tempPath, Outlook.OlSaveAsType.olMSGUnicode);
                forwardMail.Attachments.Add(tempPath, Outlook.OlAttachmentType.olByValue, Type.Missing, currentMail.Subject);
                FileLogger.Info("Attached message: " + tempPath);

                forwardMail.Send();
                FileLogger.Info("Phish report email sent successfully");

                // Move reported email to Deleted Items
                if (mailToDelete != null)
                {
                    MoveToDeletedItems(mailToDelete);
                }

                // Close the inspector window
                try
                {
                    inspector.Close(Outlook.OlInspectorClose.olDiscard);
                }
                catch (Exception closeEx)
                {
                    FileLogger.Error("Failed to close inspector window", closeEx);
                }

                // Updated success message with consistent title and message
                MessageBox.Show(
                    "Thank you for reporting this email as phishing. It has been moved to the Deleted Items folder.",
                    "Report phishing",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                FileLogger.Error("Error creating phish report", ex);
                MessageBox.Show("Failed to report the email. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up temporary file
                if (!string.IsNullOrEmpty(tempPath))
                {
                    CleanupTempFiles(new List<string> { tempPath });
                }
            }
        }

        private string ReadRegistryValue(string keyName, string valueName)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyName))
                {
                    if (key != null)
                    {
                        object value = key.GetValue(valueName);
                        if (value != null)
                        {
                            return value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("Error reading registry key: " + keyName + ", value: " + valueName, ex);
            }
            return null;
        }

        private bool IsValidEmail(string emailAddress)
        {
            if (string.IsNullOrWhiteSpace(emailAddress))
                return false;

            try
            {
                MailAddress mailAddress = new MailAddress(emailAddress);
                return mailAddress.Address == emailAddress;
            }
            catch
            {
                return false;
            }
        }

        private void CleanupTempFiles(List<string> tempFiles)
        {
            if (tempFiles == null || tempFiles.Count == 0)
                return;

            foreach (string tempFile in tempFiles)
            {
                try
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                        FileLogger.Info("Deleted temporary file: " + tempFile);
                    }
                }
                catch (Exception ex)
                {
                    FileLogger.Error("Failed to delete temporary file: " + tempFile, ex);
                }
            }
        }

        private void MoveToDeletedItems(Outlook.MailItem mail)
        {
            try
            {
                if (mail != null)
                {
                    var application = Globals.ThisAddIn.Application;
                    Outlook.MAPIFolder deletedItemsFolder = application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);

                    mail.Move(deletedItemsFolder);
                    FileLogger.Info("Moved email to Deleted Items folder: " + mail.Subject);
                }
            }
            catch (Exception ex)
            {
                FileLogger.Error("Failed to move email to Deleted Items", ex);
                // Don't throw - this is not critical, the report was already sent
            }
        }
    }
}
