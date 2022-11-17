# Outlook VBA Script - Open ticket in PerfexCRM

This VBA scripts can be used to create a macro at Microsoft Outlook. It takes the selected e-mail data, fills a form to manual adjustments, then allows open a ticket at your PerfexCRM instance.

## Requirements

- [REST API module for Perfex CRM](https://codecanyon.net/item/rest-api-for-perfex-crm/25278359)
- [VBA Json](https://github.com/VBA-tools/VBA-JSON)

## Installation

Import all files to VBA:
- JsonConverter.bas
- PerfexCRM.bas
- frmPerfexCRM.frm
- frmPerfexCRM.frx

Add Dictionary reference/class
- For Windows-only, include a reference to "Microsoft Scripting Runtime"
- For Windows and Mac, include VBA-Dictionary

For more information, take a look: [VBA Json](https://github.com/VBA-tools/VBA-JSON)


Open PerfexCRM.bas and replace the Const with your data:

```vba
Public Const perfexcrm_url As String = "https://myperfexcrm.com"
Public Const authtoken As String = "<...>"
Public Const default_priority As Integer = 2
Public Const default_departmentID As Integer = 5
```
Customize your outlook ribbon. Add a button to trigger the macro PerfexCRM_OpenTicket

## Usage

1. Select the e-mail you wanna transform into a ticket.
2. Hit the button to run the macro.
3. Review all the data at the form and do adjustments if needed.
4. Click "Open ticket" to post it to you PerfexCRM instance.

## Contributing

Pull requests are welcome. Please make sure to update tests as appropriate.

## License

Absolutely free and shared.
