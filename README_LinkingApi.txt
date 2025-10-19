CrmRegardingAddin — Linking API (Prepare vs Commit)

This patch adds a new file: LinkingApi.cs (VS2015, C#6), and updates CrmRegardingAddin.csproj to include it.

Goals:
- Clean separation between:
  1) Prepare: write only Outlook/Exchange named properties (UDF/DASL) — NO CRM I/O
  2) Commit: write only to CRM — NO Outlook/Exchange writes

Public methods:

// MAIL
LinkingApi.PrepareMailLinkInOutlookStore(Outlook.MailItem mi, Guid regardingId, string regardingLogicalName, string regardingReadableName)
LinkingApi.CommitMailLinkToCrm(IOrganizationService org, Outlook.MailItem mi)

// APPOINTMENT
LinkingApi.PrepareAppointmentLinkInOutlookStore(Outlook.AppointmentItem appt, Guid regardingId, string regardingLogicalName, string regardingReadableName)
LinkingApi.CommitAppointmentLinkToCrm(IOrganizationService org, Outlook.AppointmentItem appt)

What the prepare step writes (ID DASL + string-named variants):
- 0x80C8 crmLinkState (PT_DOUBLE) = 1.0
- 0x80C9 crmRegardingId (PT_UNICODE) = {GUID} (uppercase with braces)
- 0x80CA crmRegardingObjectType (PT_UNICODE) = logical name (e.g. "contact", "account", "new_instance")
- 0x80CB Regarding (PT_UNICODE) = human-readable label (best effort)

It also writes the string-named names for compatibility, including variants historically used by Microsoft's add-in:
- "crmLinkState"
- "crmRegardingId"
- "crmRegardingObjectType"
- "crmregardingobjectid" / "crmRegardingObjectId"
- "crmregardingobjecttypecode" / "crmRegardingObjectTypeCode"
- "Regarding"

The commit step:
- Reads the prepared props from the item
- For emails: finds or creates the CRM email by InternetMessageId; sets Regarding + parties; sets directioncode heuristically
- For appointments: finds or creates the CRM appointment by GlobalAppointmentID; sets Regarding + attendees
- Does NOT change any Outlook/Exchange property (no SetProperty, no Save)

Integration hint (example):
- On compose/new item: call Prepare* immediately after the user picks a Regarding
- On send (MailItem.Send event) → call CommitMailLinkToCrm
- On save of a newly created appointment (AppointmentItem.Write or AfterWrite) → call CommitAppointmentLinkToCrm

Notes:
- No existing methods were modified, except CrmRegardingAddin.csproj (added LinkingApi.cs).
- All code targets .NET 4.7.2 and uses only APIs already referenced in this project.

