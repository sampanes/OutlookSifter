using System;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Text; // This fixes the StringBuilder error
using System.Text.RegularExpressions;

namespace OutlookToClipboard;

static class Program
{
    [STAThread]
    static void Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ApplicationConfiguration.Initialize();
        Form mainForm = new Form { Text = "Universal Email Sifter", Size = new System.Drawing.Size(600, 500), TopMost = true };
        TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2 };
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 25f));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 75f));

        Label dropLabel = new Label { Text = "DRAG EMAIL HERE", TextAlign = System.Drawing.ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, AllowDrop = true, Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold), BorderStyle = BorderStyle.FixedSingle };
        TextBox previewBox = new TextBox { Multiline = true, ReadOnly = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, Font = new System.Drawing.Font("Consolas", 9) };

        dropLabel.DragEnter += (s, e) => e.Effect = DragDropEffects.Copy;

        dropLabel.DragDrop += (s, e) =>
        {
            try
            {
                if (e.Data == null) return;
                string rawData = "";

                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
                    if (files.Length > 0) rawData = File.ReadAllText(files[0]);
                }
                else if (e.Data.GetDataPresent(DataFormats.UnicodeText))
                {
                    rawData = (string)e.Data.GetData(DataFormats.UnicodeText)!;
                }

                if (!string.IsNullOrEmpty(rawData))
                {
                    // 1. EXTRACTION: Pluck the Headers
                    string subject = Regex.Match(rawData, @"^Subject:\s*(.*)", RegexOptions.Multiline).Groups[1].Value.Trim();
                    string from = Regex.Match(rawData, @"^From:\s*(.*)", RegexOptions.Multiline).Groups[1].Value.Trim();
                    string date = Regex.Match(rawData, @"^Date:\s*(.*)", RegexOptions.Multiline).Groups[1].Value.Trim();
                    string to = Regex.Match(rawData, @"^To:\s*(.*)", RegexOptions.Multiline).Groups[1].Value.Trim();

                    // 2. WASHING: Clean the Body
                    string body = rawData;
                    var bodyMatch = Regex.Match(rawData, @"<body[\s\S]*?>([\s\S]*)</body>", RegexOptions.IgnoreCase);
                    if (bodyMatch.Success)
                    {
                        body = bodyMatch.Groups[1].Value;
                    }
                    else
                    {
                        // Fallback: If no HTML body, find the double-newline after headers
                        var m = Regex.Match(rawData, @"\r?\n\r?\n([\s\S]*)");
                        if (m.Success) body = m.Groups[1].Value;
                    }

                    // Decode quoted-printable encoding (=XX hex sequences, soft line breaks)
                    body = Regex.Replace(body, @"=\r?\n", "");
                    body = Regex.Replace(body, @"=([0-9A-Fa-f]{2})", m => Encoding.GetEncoding(1252).GetString(new byte[] { Convert.ToByte(m.Groups[1].Value, 16) }));

                    // Strip HTML tags and decode entities
                    body = Regex.Replace(body, @"<[^>]*>", string.Empty);
                    body = System.Net.WebUtility.HtmlDecode(body);

                    // Normalize all Unicode dashes/hyphens to plain hyphen
                    body = Regex.Replace(body, @"[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]", "-");

                    // Remove Base64 "Garbage" (images/attachments)
                    body = Regex.Replace(body, @"[A-Za-z0-9+/]{100,}", "");

                    // 3. ASSEMBLY: Format the final output
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(subject);
                    sb.AppendLine($"From {from}");
                    sb.AppendLine($"Date {date}");
                    sb.AppendLine($"To {to}");
                    sb.AppendLine();
                    sb.AppendLine(body.Trim());

                    string finalOutput = sb.ToString();

                    previewBox.Text = finalOutput;
                    Clipboard.SetText(finalOutput);

                    dropLabel.Text = "CLEAN DATA COPIED!";
                    dropLabel.BackColor = System.Drawing.Color.PaleGreen;
                    var timer = new System.Windows.Forms.Timer { Interval = 1500 };
                    timer.Tick += (ts, te) => { dropLabel.Text = "DRAG EMAIL HERE"; dropLabel.BackColor = System.Drawing.Color.Transparent; timer.Stop(); };
                    timer.Start();
                }
            }
            catch (Exception ex)
            {
                previewBox.Text = "Error: " + ex.Message;
            }
        };

        layout.Controls.Add(dropLabel, 0, 0);
        layout.Controls.Add(previewBox, 0, 1);
        mainForm.Controls.Add(layout);
        Application.Run(mainForm);
    }
}