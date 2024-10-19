using Microsoft.Office.Tools.Ribbon;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PhishingReportAddin
{
    public partial class PhishingRibbon
    {
        private void PhishingRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("PhishingRibbon_Load called");
            this.tab1.Visible = true;
        }

        private void btnReportPhishing_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var explorer = app.ActiveExplorer();
            if (explorer.Selection.Count > 0)
            {
                var selectedItem = explorer.Selection[1] as Outlook.MailItem;
                if (selectedItem != null)
                {
                    // Confirm with the user before forwarding
                    DialogResult result = MessageBox.Show(
                        $"Are you sure you want to report this email as phishing?\n\nSubject: {selectedItem.Subject}\nFrom: {selectedItem.SenderName} ({selectedItem.SenderEmailAddress})",
                        "Confirm Phishing Report",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.Yes)
                    {
                        var forwardedMail = selectedItem.Forward();
                        forwardedMail.To = "golemgrumpy@gmail.com";
                        forwardedMail.Subject = "Potential Phishing Email - " + ObfuscateDomains(selectedItem.Subject);
                        string obfuscatedBody = ObfuscateDomains(selectedItem.Body);
                        forwardedMail.HTMLBody = "This email has been reported as a potential phishing attempt. Domains have been obfuscated.<br><br>" + obfuscatedBody;

                        // Copy attachments
                        foreach (Outlook.Attachment attachment in selectedItem.Attachments)
                        {
                            forwardedMail.Attachments.Add(attachment);
                        }

                        forwardedMail.Send();
                        MessageBox.Show("Email reported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private string ObfuscateDomains(string text)
        {
            string pattern = @"(https?:\/\/)?[\w\-]+(\.[\w\-]+)+\.?(:\d+)?(\/\S*)?";
            return Regex.Replace(text, pattern, match =>
            {
                return match.Value.Replace(".", "[.]");
            });
        }
    }
}