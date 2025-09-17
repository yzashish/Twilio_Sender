# Developer Guide: Twilio WhatsApp Messenger
*Version 2.0*

## 1. Architecture Overview

This tool facilitates sending bulk, templated WhatsApp messages from Excel using Twilio's Content API. The architecture is designed for robustness and compliance with WhatsApp's policies.

1.  **`create_excel_structure.py`**: A Python script using `openpyxl` that generates the structured `TwilioSender.xlsx` workbook. This script defines the data model for the entire tool.
2.  **`twilio_macro.vbs`**: A VBA script containing all the client-side logic. It reads data from the sheets, constructs the API payload, makes the API calls, and logs the results. It is manually imported by the user.
3.  **Hybrid-Variable Data Model**: The core of this tool is its two-tiered variable system. This allows for a flexible combination of static (per-template) and dynamic (per-contact) data in the final message.

## 2. File & Sheet Descriptions

### `create_excel_structure.py`
- **Purpose**: To programmatically generate the Excel file with the correct sheets and headers.
- **Key Functionality**:
    - Creates five worksheets: `Dashboard`, `Contacts`, `Templates`, `Log`, and `Settings`.
    - **`Contacts` Sheet**: Contains columns for `Name`, `WhatsApp Number`, `Send?`, `Status`, and 15 `Contact Variable {{i}}` columns for personalization.
    - **`Templates` Sheet**: Contains columns for `Message UID`, `ContentSid`, and 15 `Base Variable {{i}}` columns. This sheet maps a user-friendly UID to a Twilio `ContentSid` and defines the static part of the message variables. It also includes several columns for reference only (`Description`, `Template Content`, `Sample Final Content`) which are populated by the Python script but not used by the VBA macro.

### `twilio_macro.vbs`
This script contains the main application logic and several helper functions.

- **`SendTwilioMessages()` Subroutine**:
    - **Workflow**:
        1.  **Configuration:** Reads credentials from `Settings` and user inputs (`Message UID`, `MPS`) from the `Dashboard`. Performs validation on these inputs.
        2.  **Template Fetching:** Looks up the `Message UID` on the `Templates` sheet to get the corresponding `ContentSid` and loads the 15 `Base Variable` values into a VBA array.
        3.  **Contact Loop:** Iterates through each contact marked "Yes" on the `Contacts` sheet.
        4.  **Two-Tiered Replacement:** For each contact, it creates a new array of `resolvedVariables`. It loops through the `baseVariables` array and, for each item, performs a second replacement pass, substituting any `{{name}}` or `{{contact_var_i}}` placeholders with data from the current contact's row.
        5.  **JSON Construction:** It calls the `BuildContentVariablesJson` helper function to serialize the `resolvedVariables` array into the JSON string required by the Twilio API.
        6.  **API Call:** It constructs the final POST request body, which now includes `ContentSid` and the URL-encoded `ContentVariables` JSON, and sends the request.
        7.  **Rate-Limiting:** Before each API call within the loop, it checks a counter against the MPS setting and pauses for 1 second using `Application.Wait` if the rate is exceeded.
        8.  **Logging:** After each API call, it writes a new row to the `Log` sheet with the timestamp, contact info, status, and the full response from the Twilio API.
    - **Error Handling**: Uses `On Error GoTo` to catch runtime errors. The handler now uses the `Erl` function to report the specific line number where the error occurred, making debugging easier.

- **Helper Functions**:
    - **`BuildContentVariablesJson(vars() As String) As String`**: A custom function to build the `{"1":"val1", "2":"val2"}` JSON string from an array of resolved variables. It only includes keys for non-empty variables.
    - **`JsonEncode(str As String) As String`**: A helper to escape characters (like quotes and backslashes) to create a valid JSON string value.
    - **`UrlEncode(str As String) As String`**: Encodes strings for safe inclusion in the `x-www-form-urlencoded` request body.
    - **`Base64Encode(str As String) As String`**: Encodes credentials for HTTP Basic Authentication.
