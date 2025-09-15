# User Guide: Twilio WhatsApp Messenger for Excel

## 1. Introduction

Welcome! This tool allows you to send bulk WhatsApp messages directly from an Excel spreadsheet using your Twilio account. You can manage your contacts and message templates all within this workbook.

This guide will walk you through a simple, one-time setup process to enable the messaging macro.

## 2. Initial Setup: Enabling Macros

Because this workbook uses a VBA macro to send messages, you need to import the provided macro script and save the file in a macro-enabled format (`.xlsm`). This is a one-time process.

### Step 2.1: Enable the Developer Tab in Excel

If you don't have a "Developer" tab in your Excel ribbon, follow these steps:
1.  Go to `File` > `Options` > `Customize Ribbon`.
2.  In the right-hand list under "Main Tabs", check the box for **Developer**.
3.  Click **OK**.

### Step 2.2: Import the VBA Macro Script

1.  Open the `TwilioSender.xlsx` file.
2.  Press **Alt + F11** to open the VBA Editor. (Or go to the `Developer` tab and click `Visual Basic`).
3.  In the VBA Editor, go to `File` > `Import File...`.
4.  Navigate to and select the `twilio_macro.vbs` file that was provided alongside this workbook. Click **Open**.
5.  You should now see `Module1` under the "Modules" folder in the "Project - VBAProject" window on the left. You can double-click it to see the imported code.
6.  Close the VBA Editor by clicking the 'X' or pressing **Alt + F11** again.

### Step 2.3: Add the "Send Messages" Button

1.  Go to the **Dashboard** sheet.
2.  Go to the `Developer` tab in the Excel ribbon.
3.  Click `Insert`, and under "ActiveX Controls", select the **Command Button**.
4.  Your cursor will turn into a crosshair. Draw a button in the "Button Placeholder" area on the Dashboard.
5.  Right-click the new button and select **Properties**.
6.  In the Properties window:
    *   Change the **(Name)** to `btnSendMessages`.
    *   Change the **Caption** to `Send Messages`.
7.  Close the Properties window.
8.  Double-click the button to open the VBA editor again. It will show a new private sub for the button click.
9.  Inside that sub, type `SendTwilioMessages` on the line between `Private Sub...` and `End Sub`. It should look like this:
    ```vba
    Private Sub btnSendMessages_Click()
        SendTwilioMessages
    End Sub
    ```
10. Close the VBA editor. Make sure you are out of "Design Mode" by clicking the `Design Mode` icon on the Developer tab (it should not be highlighted).

### Step 2.4: Save as a Macro-Enabled Workbook

1.  Go to `File` > `Save As`.
2.  Choose a location for the file.
3.  In the "Save as type" dropdown, select **Excel Macro-Enabled Workbook (*.xlsm)**.
4.  Click **Save**. You can now close the original `.xlsx` file and use your new `.xlsm` file from now on.

**Your setup is now complete!**

## 3. How to Use the Messenger

### Step 3.1: Configure Your Settings
1.  Go to the **Settings** sheet.
2.  Enter your Twilio Account SID, Auth Token, and your Twilio "From" WhatsApp number in the corresponding cells.
    *Note: The sheet is password-protected ('twilio') to prevent accidental changes, but the value cells are editable.*

### Step 3.2: Add Your Contacts
1.  Go to the **Contacts** sheet.
2.  Enter the `Name` and `WhatsApp Number` for each recipient.
    *   The number should include the country code but no `+` or `00` (e.g., `919876543210`).
3.  In the `Send?` column, select "Yes" from the dropdown for each contact you want to message.
4.  If your template uses variables like `{{1}}`, `{{2}}`, etc., fill in the corresponding `Variable` columns for each contact.

### Step 3.3: Manage Your Templates
1.  Go to the **Templates** sheet.
2.  You can edit the sample templates or add new ones.
3.  Make sure each template has a unique value in the **Message UID** column. This UID is how you will select the message to send.

### Step 3.4: Send the Messages
1.  Go to the **Dashboard** sheet.
2.  In the yellow box (cell `C5`), enter the **Message UID** of the template you want to send.
3.  Click the **Send Messages** button.
4.  The button will be disabled and read "Sending...". Please wait until the process is complete.
5.  A confirmation message will appear once all messages have been sent. Check the `Status` column in the `Contacts` sheet for the result of each message.
