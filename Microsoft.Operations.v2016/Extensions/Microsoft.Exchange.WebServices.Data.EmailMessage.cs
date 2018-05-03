using Microsoft.Exchange.WebServices.Data;

/// <summary>
/// Anything related to efficiently processing Exchange Web Services (EWS) emails (all common routines)
/// </summary>
public static class ExtensionMethodsEmail
{
    /// <summary>
    /// This function is intended to be common to all routines which regularly look at Service
    /// Account Mailbox contents. It provides a COARSE assessment of whether the email message is
    /// 'JUNK' or not, based on predefined rules, most of the time it is deemed junk because the item
    /// is inherently *** NON-ACTIONABLE by the service account ***. A text summary will be provided
    /// on the assessment, but no other detail.
    ///
    /// The following general assumptions are made (this may help you determine if you wish to use it):
    /// 1. Anything from a non-microsoft address is automatically junk, for security reasons.
    /// 2. Any 'auto-reply' is junk - i.e. completely non actionable by a service account.
    /// 3. Any broad distribution mail or notification item is junk.
    /// </summary>
    /// <param name="email">The EmailMessage object that needs to be analysed.</param>
    /// <param name="explanationOfDecisionForLoggingPurposes">
    /// Text which should be used for logging information about the determination. NOTE: the address
    /// is NOT included as part of this output.
    /// </param>
    /// <param name="confidenceLevel">
    /// Value between 0 and 100, relates to the clarity of the initial decision - could be used for
    /// making decisions about whether to HardDelete or not, depending on your audit requirements.
    /// </param>
    public static bool IsJunkMail(this EmailMessage email, out string explanationOfDecisionForLoggingPurposes, out int confidenceLevel)
    {
        bool isJunk;
        string from = Optimize.SafeString(email.From.Address).ToLower(); // should always have a value
        string subject = Optimize.SafeString(email.Subject);

        //if (!from.Contains("@microsoft.com"))
        //{
        //    explanationOfDecisionForLoggingPurposes = "For security reasons, this service cannot accept emails from EXTERNAL users, regardless of content.";
        //    isJunk = true;
        //    confidenceLevel = 75; // because occasionally legitimate items do come from external accounts (for whatever reason) and they may require action.
        //}

        if (from == "gmodops@microsoft.com")
        {
            explanationOfDecisionForLoggingPurposes = "Broadcast email from Security group membership.";
            isJunk = true;
            confidenceLevel = 90; // because there is still a chance this might be impacting.
        }
        else if (from == "invitations@linkedin.com")
        {
            explanationOfDecisionForLoggingPurposes = "LinkedIn invitations are known be be a vector for spam/phishing - often fake and repeated attempts.";
            isJunk = true;
            confidenceLevel = 100; // because the service account has no business working with LinkedIn profiles.
        }
        else if (from == "tfsnot@microsoft.com" || from == "bgitnot@microsoft.com")
        {
            explanationOfDecisionForLoggingPurposes = "This is a generic VSTF Notification email about maintenance on server - typically sent to ALL users registered in a BGIT-hosted TFS Server instance.)";
            isJunk = true;
            confidenceLevel = 100; // because the same emails are already going to sysadmins, administrators
        }
        else if (from == "mpsdsr@microsoft.com")
        {
            explanationOfDecisionForLoggingPurposes = "No idea why emails from this support group, are coming here.";
            isJunk = true;
            confidenceLevel = 100; // because yammer just broadcasts crap indiscriminately
        }
        else if (from == "notifications@yammerqa.com")
        {
            explanationOfDecisionForLoggingPurposes = "Yammer notifications - argh! Now a very common junk item, first wave started in 2014/2015.";
            isJunk = true;
            confidenceLevel = 100; // because yammer just broadcasts crap indiscriminately
        }
        else if (from.Contains("v-srinb"))
        {
            explanationOfDecisionForLoggingPurposes = "Well-meaning vendors who nonetheless spam entire threads/service mailboxes with random information because they don't diligency prune TO: and CC:";
            isJunk = true;
            confidenceLevel = 100;
        }
        else if (from == "itex@microsoft.com")
        {
            explanationOfDecisionForLoggingPurposes = "Default email from IT systems - usually relating to network or systems status.";
            isJunk = true;
            confidenceLevel = 100; // because generally these don't contain anything that's useful - these are the 'we are obliged to notify you' type-emails.
        }
        else if (from == "azureTeam@e-mail.microsoft.com")
        {
            explanationOfDecisionForLoggingPurposes = "Default news email from Azure systems because service account has a registration, not required to view or read.";
            isJunk = true;
            confidenceLevel = 100; // because service accounts don't need to read news
        }
        else if (subject.StartsWith("Automatic reply") || subject.StartsWith("Răspuns automat") || subject.StartsWith("Respuesta automática") || subject.StartsWith("Automatische Antwort") || subject.StartsWith("Réponse automatique") || subject.StartsWith("Auto Response") || subject.StartsWith("Automatikus válasz") || subject.StartsWith("Automatisch antwoord") || subject.StartsWith("Autosvar"))
        {
            // It's difficult to get a rule which covers auto-reply in all languages. Add patterns as
            // required, or if it gets too unruly, create a better index/pattern matching structure.
            explanationOfDecisionForLoggingPurposes = "This is an automatic reply from another email previously sent, it can be ignored safely because it truly requires no action.";
            isJunk = true;
            confidenceLevel = 100;
        }
        else if (subject.Contains("Awareness: Commercial Data Platform"))
        {
            explanationOfDecisionForLoggingPurposes = "This is a notice regarding the availability of DataMart databases which the service account may be a member of (average: 2 emails/week).";
            isJunk = true;
            confidenceLevel = 100; // because it really is junk
        }
        else if (subject.StartsWith("Clutter "))
        {
            explanationOfDecisionForLoggingPurposes = "The self-referencing irony of a clutter-removal service adding more clutter is palpable.";
            isJunk = true;
            confidenceLevel = 80; // because it is still possible someone might send an important email starting with 'Clutter'
        }
        else if (subject.StartsWith("Daily schedule on") && from.Contains("no-reply@microsoft.com"))
        {
            explanationOfDecisionForLoggingPurposes = "This is some kind of weird/strange outlook calendar thing, which has no place in a service account mailbox!";
            isJunk = true;
            confidenceLevel = 100;
        }
        else if (subject == "You're Invited to participate in the Microsoft Cloud Solution Provider program!")
        {
            explanationOfDecisionForLoggingPurposes = "SPAM, from our own account, somehow (and is not a forward or reply to)";
            isJunk = true;
            confidenceLevel = 100;
        }
        else if (subject.Contains("CRMSMB data stale?"))
        {
            explanationOfDecisionForLoggingPurposes = "This is a stupid email which won't be removed properly. Noise from the CRM team.";
            isJunk = true;
            confidenceLevel = 100;
        }
        else
        {
            // If none of the above conditions were 'tripped' then it's not junk.
            explanationOfDecisionForLoggingPurposes = "NOT junk mail, according to all our pre-defined conditions.";
            isJunk = false;
            confidenceLevel = 100; // because it was quite a comprehensive test.
        }

        return isJunk;
    }
}