
README_CrmMapiInterop
=====================

Add CrmMapiInterop.cs to your project.

Usage in CrmActions (examples)
------------------------------
MAIL after you set the CRM regarding:
    CrmMapiInterop.ApplyMsCompatForMail(
        mail,
        "",                 // regardingLogicalName or numeric code as string if known
        regardingId,        // Guid
        regardingDisplay,   // e.g. "ACCOUNT: Contoso"
        currentUserSmtp,    // your Outlook/CRM user's SMTP
        currentUserCrmId,   // Guid? of systemuser
        fromMember,         // CrmMapiInterop.CrmPartyMember (sender information)
        recipientMembers,   // IEnumerable<CrmMapiInterop.CrmPartyMember>
        isIncoming          // true for received, false for sent
    );

On unlink (keep CRM record):
    CrmMapiInterop.RemoveMsCompatFromMail(mail);

APPOINTMENT when linking:
    CrmMapiInterop.ApplyMsCompatForAppointment(
        appt,
        regardingId,
        regardingDisplay,
        organizerSmtp,
        currentUserCrmId
    );

On unlink:
    CrmMapiInterop.RemoveMsCompatFromAppointment(appt);
