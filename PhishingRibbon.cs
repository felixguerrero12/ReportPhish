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
            try
            {
                var app = Globals.ThisAddIn.Application;
                var explorer = app.ActiveExplorer();

                // Check if we have a valid selection
                if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email first.", "No Email Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Get the selected item (using index 1 as Outlook uses 1-based indexing)
                var selectedItem = explorer.Selection[1];

                // Check if it's actually an email
                if (!(selectedItem is Outlook.MailItem))
                {
                    MessageBox.Show("Please select an email message.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var mailItem = selectedItem as Outlook.MailItem;

                // Confirm with the user before forwarding
                DialogResult result = MessageBox.Show(
                    $"Are you sure you want to report this email as phishing?\n\nSubject: {mailItem.Subject}\nFrom: {mailItem.SenderName} ({mailItem.SenderEmailAddress})",
                    "Confirm Phishing Report",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    // Create and show input dialog for description
                    Form descriptionForm = new Form()
                    {
                        Width = 500,
                        Height = 200,
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        Text = "Provide Description",
                        StartPosition = FormStartPosition.CenterScreen
                    };

                    TextBox textBox = new TextBox()
                    {
                        Multiline = true,
                        ScrollBars = ScrollBars.Vertical,
                        Size = new System.Drawing.Size(450, 80),
                        Location = new System.Drawing.Point(20, 20)
                    };

                    Button confirmation = new Button()
                    {
                        Text = "Submit",
                        DialogResult = DialogResult.OK,
                        Location = new System.Drawing.Point(360, 120)
                    };

                    Label label = new Label()
                    {
                        Text = "Please provide a brief description of why this email appears suspicious:",
                        Location = new System.Drawing.Point(20, 5),
                        AutoSize = true
                    };

                    descriptionForm.Controls.Add(textBox);
                    descriptionForm.Controls.Add(confirmation);
                    descriptionForm.Controls.Add(label);
                    descriptionForm.AcceptButton = confirmation;

                    if (descriptionForm.ShowDialog() == DialogResult.OK)
                    {
                        var forwardedMail = mailItem.Forward();
                        forwardedMail.To = "golemgrumpy@gmail.com";
                        forwardedMail.Subject = "Potential Phishing Email - " + ObfuscateDomains(mailItem.Subject);

                        string userDescription = textBox.Text.Trim();
                        string obfuscatedBody = ObfuscateDomains(mailItem.Body);

                        forwardedMail.HTMLBody =
                            "This email has been reported as a potential phishing attempt.<br><br>" +
                            "<b>Reporter's Description:</b><br>" +
                            userDescription + "<br><br>" +
                            "<b>Original Email (with obfuscated domains):</b><br><br>" +
                            obfuscatedBody;

                        // Copy attachments
                        foreach (Outlook.Attachment attachment in mailItem.Attachments)
                        {
                            forwardedMail.Attachments.Add(attachment);
                        }

                        forwardedMail.Send();
                        MessageBox.Show("Email reported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine($"Error in btnReportPhishing_Click: {ex}");
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