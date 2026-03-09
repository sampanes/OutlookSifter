# 📧 Outlook Sifter

**Outlook Sifter** is a lightweight Windows utility designed to solve the "New Outlook" drag-and-drop headache. It allows you to drag an email directly from your Outlook inbox into the app to instantly get a clean, plain-text version of the message—including key headers—copied to your clipboard.

## 🚀 The Problem
The "New" Microsoft Outlook (Chromium-based) no longer allows you to drag-and-drop email text directly into web browsers or many third-party apps without getting a "mumbly" mess of HTML code, Base64 image data (like company logos), and metadata.

## ✨ Features
* **Direct Drag-and-Drop**: Works specifically with the "FileDrop" format used by New Outlook.
* **Header Extraction**: Automatically pulls and formats:
    * Subject
    * From (Sender)
    * Date
    * To (Recipients)
* **Smart Cleaning**: 
    * Strips all HTML tags and CSS styles.
    * Decodes HTML entities (like `&nbsp;`) into readable text.
    * **Garbage Disposal**: Automatically detects and removes massive Base64 strings (embedded images/attachments) to keep your text clean.
* **Instant Clipboard**: As soon as you drop the email, the cleaned text is placed on your Windows clipboard.

## 🛠️ How to Use
1.  **Download**: Grab the latest `OutlookSifter.exe` from the [Releases](#) tab.
2.  **Run**: Open the app (it stays "Always on Top" for easy access).
3.  **Sift**: Drag any email from your Outlook list onto the **"DRAG EMAIL HERE"** zone.
4.  **Paste**: Your clean text is now ready to be pasted anywhere (Notepad, Teams, Jira, etc.).

## 💻 Technical Details
* **Framework**: .NET 10.0 (Windows Forms).
* **Language**: C# 14.
* **Architecture**: 100% Client-side. No data ever leaves your machine.

## 🔨 Building from Source
If you have the .NET 10 SDK installed, you can build your own standalone executable:

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
