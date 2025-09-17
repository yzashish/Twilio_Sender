# User Guide: Twilio WhatsApp Messenger
*Version 2.6 (Final)*

## 1. Introduction

Welcome! This tool allows you to send bulk, personalized WhatsApp messages directly from Excel using Twilio's official Content API for approved templates.

This guide will walk you through setting up and using the tool, including its powerful hybrid variable system.

## 2. Initial Setup: Enabling Macros

This is a one-time process to enable the macro that powers the tool.

### Step 2.1: Enable the Developer Tab
If you don't have a "Developer" tab in your Excel ribbon, follow these steps:
1.  Go to `File` > `Options` > `Customize Ribbon`.
2.  In the right-hand list, check the box for **Developer**. Click **OK**.

### Step 2.2: Import the VBA Macro
1.  Open the `TwilioSender.xlsx` file.
2.  Press **Alt + F11** to open the VBA Editor.
3.  Go to `File` > `Import File...` and select the **`twilio_macro_v2.6.vbs`** file.
4.  Close the VBA Editor.

### Step 2.3: Add the "Send Messages" Button
1.  On the **Dashboard** sheet, go to the `Developer` tab and ensure **Design Mode** is turned on.
2.  Click the `Insert` dropdown. In the second section, labeled **ActiveX Controls**, select the **Command Button** icon.
3.  Draw the new button on the sheet.
3.  Right-click the button, select **Properties**, and change its **(Name)** to `btnSendMessages` and its **Caption** to `Send Messages`.
4.  Double-click the button and, in the code window that appears, type `SendTwilioMessages` between the `Private Sub` and `End Sub` lines.

### Step 2.4: Save as Macro-Enabled Workbook
1.  Go to `File` > `Save As`.
2.  Change the "Save as type" to **Excel Macro-Enabled Workbook (*.xlsm)**.
3.  Save the file. You will use this new `.xlsm` file from now on.

**Setup is now complete!**

## 3. How It Works: The Hybrid Variable System

This tool uses a powerful two-tiered variable system:
- **Base Variables:** These are defined once per message template in the `Templates` sheet. They are useful for content that is the same for every contact receiving that message (e.g., an offer detail, a campaign name).
- **Contact Variables:** These are defined for each person in the `Contacts` sheet. They are used for personalization (e.g., names, appointment times, specific product details).

The magic is that you can **use Contact Variables inside Base Variables**.

## 4. Step-by-Step Usage

### Step 4.1: Configure Settings
- Go to the **Settings** sheet and enter your Twilio Account SID, Auth Token, and your Twilio "From" WhatsApp number.

### Step 4.2: Define Your Message Templates
1.  Go to the **Templates** sheet.
2.  For each pre-approved Twilio template you want to use:
    *   **Message UID:** Give it a short, memorable name (e.g., `PROMO_01`).
    *   **ContentSid:** Paste the official Template SID from your Twilio account (e.g., `HX...`).
    *   **Description / Template Content:** Use these columns for your own reference to remember what each template is for. **Editing these columns will not change the message that is sent.**
    *   **Sample Final Content:** This column provides a helpful preview of what your message will look like after the Base Variables are filled in.
    *   **Base Variable {{1}} - {{15}}:** Fill in the values for the placeholders in your Twilio template.
        *   **Example:** If Base Variable `{{2}}` is for a greeting, you could enter `Hello {{name}}! Check out this offer.` Here, `{{name}}` is a placeholder that will be filled in from the `Contacts` sheet.

### Step 4.3: Add Your Contacts
1.  Go to the **Contacts** sheet.
2.  For each recipient, fill in their `Name` and `WhatsApp Number`.
3.  Set `Send?` to "Yes" for everyone you want to message.
4.  **Contact Variable {{1}} - {{15}}:** Fill in any data that is unique to this contact. This data can be used in the templates, as seen in the example above.

### Step 4.4: Send Your Campaign
1.  Go to the **Dashboard**.
2.  Enter the **Message UID** you want to send in cell `C5`.
3.  Set your desired **Messages Per Second (MPS)** rate in cell `C6` (e.g., `1`).
4.  Click the **Send Messages** button.

### Step 4.5: Review the Log
- After sending, go to the **Log** sheet to see a detailed record of every message sent, including the timestamp and success/failure status.
