# Outlook Contact Saver Add-in

## What Was Built
The Outlook Contact Saver is an add-in that automatically extracts contact details—such as name, email, phone number, and job title—from email signatures. It allows users to review the extracted information in a side panel before saving it directly to their Outlook contacts and a centralized contact database stored as an Excel file in OneDrive.

## Setup
1. **Host the Project**: Upload the files to a GitHub repository and enable **GitHub Pages** in the repository settings.
2. **Register Azure App**: 
   - Create a new "App registration" in the Azure Portal.
   - Set the platform to **Single-page application (SPA)** and enter your GitHub Pages URL followed by `/index.html` as the Redirect URI.
   - Add the following Microsoft Graph API permissions: `Contacts.ReadWrite` and `Files.ReadWrite`.
3. **Configure Settings**: Open `config.js` and replace the placeholders with your **Azure Client ID** and **GitHub Pages URL**. Push these changes to GitHub.
4. **Create Icons**: Open `generate-icons.html` in a web browser, download the four generated PNG icons, and upload them to your GitHub repository root.
5. **Install Add-in**: In Outlook, go to **Get Add-ins** $\rightarrow$ **My Add-ins** $\rightarrow$ **Add from URL** (or upload the `manifest.xml` file) to install the tool.

## How to Use
1. Open any email in Outlook.
2. Click the **Save Contact** button in the reading pane toolbar.
3. Sign in with your Microsoft account when prompted.
4. Review the extracted details in the side panel and make any necessary corrections.
5. Click **Save Contact** to create a new entry, or **Update Existing** if the contact is already in your address book.

## Files
- `manifest.xml`: Configuration file used to install the add-in in Outlook.
- `config.js`: Stores the Azure Client ID and hosting URLs.
- `auth.js`: Handles user login and authentication tokens.
- `extractor.js`: Contains the logic to find contact details in email text.
- `graph.js`: Manages communication with the Microsoft Graph API.
- `taskpane.js`: Controls the logic and behavior of the user interface.
- `taskpane.html`: The structure of the side panel.
- `taskpane.css`: The styling for the side panel.
- `index.html`: Handles the authentication redirect process.
- `generate-icons.html`: A tool to create the required add-in icons.
- `test.js`: A script to verify the extraction logic.

## Notes
- **OneDrive Integration**: The tool automatically creates a file named `OutlookContacts.xlsx` in your OneDrive root folder upon the first save.
- **Extraction Accuracy**: The tool uses pattern matching to find signatures; accuracy may vary depending on the sender's signature format.
- **Permissions**: Users must grant consent for the app to read/write contacts and files during the first login.