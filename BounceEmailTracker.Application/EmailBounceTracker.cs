using System;
using System.Collections.Generic;
using System.Text;
using BounceEmailTracker.Web.Pop3;
using System.Diagnostics;

namespace BounceEmailTracker.Application
{
    public class EmailBounceTracker
    {
        private static string GetEmailFromValue(string szValue)
        {
            int nIndex = szValue.IndexOf(';');
            if (0 <= nIndex)
                szValue = szValue.Substring(nIndex + 1).Trim();

            nIndex = szValue.IndexOf('<');
            if (0 <= nIndex)
                szValue = szValue.Substring(nIndex + 1).Trim();

            nIndex = szValue.IndexOf('>');
            if (0 < nIndex)
                szValue = szValue.Substring(0, nIndex).Trim();

            return szValue;
        }
        private static string CleanSubject(string szSubject)
        {
            string[] strs = "[SPAM]".Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string s in strs)
                szSubject = szSubject.Replace(s, "").Trim();

            return szSubject;
        }
        private static string GetValueFromKey(string szData, string szKey)
        {
            string szValue = "";
            int nIndex = szData.ToLower().IndexOf(szKey.ToLower());
            if (0 <= nIndex)
            {
                // let's see if this is the beginning of a word
                bool bFoundWholeWord = false;
                while (0 < nIndex)
                {
                    // look at the character before the beginning of the substring
                    char c = szData[nIndex - 1];
                    switch (c)
                    {
                        case '\n':
                            bFoundWholeWord = true;
                            break;
                        default:
                            nIndex = szData.ToLower().IndexOf(szKey.ToLower(), nIndex + 1);
                            break;
                    }

                    if (bFoundWholeWord || 0 > nIndex)
                        break;
                }

                // if not found as a whole word, then return empty string
                if (0 > nIndex)
                    return "";

                szValue = szData.Substring(nIndex + szKey.Length).Trim();

                nIndex = szValue.IndexOf('\n');
                if (-1 == nIndex)
                    nIndex = szValue.IndexOf('\r');

                if (0 < nIndex)
                    szValue = szValue.Substring(0, nIndex).Trim();
            }

            return szValue;
        }
        public static string FindAndRecordEmailBounces()
        {
            Pop3Client email = null;
            string message = string.Empty;
            try
            {

                email = new Pop3Client("noreply@thebirthdayregister.com", "deadFax186", "mail.thebirthdayregister.com");
                email.OpenInbox();

                // only process x amount of messages at a time because deleting is not really done until the session is closed
                const long nMaxProcess = 2;
                long nTotalCount = Math.Min(nMaxProcess, email.MessageCount);
                long nMaxEmail = email.MessageCount;
                long nMinEmail = Math.Max(0, email.MessageCount - nMaxProcess);
                for (long i = nMaxEmail - 1; i >= nMinEmail; i--)
                {
                    email.NextEmail(i);
                    if (0 != string.Compare(email.To, "noreply@thebirthdayregister.com", true))
                        continue;
                    // check the subject, delete items as appropriate
                    if (email.Subject.ToLower().Contains("delay"))
                    {
                        email.DeleteEmail();
                    }
                    string szRecipient = "", szAction = "", szStatus = "", szTo = "", szFrom = "", szSubject = "", szDate = "", szReply_to = "";
                    if (email.Subject.ToLower().Contains("failure notice") || email.Subject.ToLower().Contains("Delivery Status Notification (Failure)") || email.Subject.ToLower().Contains("Mail Delivery Failure") || email.Subject.ToLower().Contains("Mail System Error - Returned Mail"))
                    {
                        if (!email.IsMultipart)
                        {
                            szTo = GetEmailFromValue(GetValueFromKey(email.Body, "To:"));
                            szFrom = GetEmailFromValue(GetValueFromKey(email.Body, "From:"));
                            szSubject = CleanSubject(GetValueFromKey(email.Body, "Subject:"));
                            szDate = GetValueFromKey(email.Body, "Date:");
                            szRecipient = GetEmailFromValue(GetValueFromKey(email.Body, "Final-Recipient:"));
                            szAction = GetValueFromKey(email.Body, "Action:");
                            szStatus = GetValueFromKey(email.Body, "Status:");
                            szReply_to = GetEmailFromValue(GetValueFromKey(email.Body, "Reply-To:"));
                        }
                        else
                        {
                            System.Collections.IEnumerator enumerator = email.MultipartEnumerator;
                            while (enumerator.MoveNext())
                            {
                                Pop3Component multipart = (Pop3Component)enumerator.Current;

                                if (string.IsNullOrEmpty(szTo))
                                    szTo = GetEmailFromValue(GetValueFromKey(multipart.Data, "To:"));
                                if (string.IsNullOrEmpty(szTo))
                                    szTo = GetEmailFromValue(GetValueFromKey(multipart.Data, "Original-Recipient:"));
                                if (string.IsNullOrEmpty(szFrom))
                                    szFrom = GetEmailFromValue(GetValueFromKey(multipart.Data, "From:"));
                                if (string.IsNullOrEmpty(szSubject))
                                    szSubject = CleanSubject(GetValueFromKey(multipart.Data, "Subject:"));
                                if (string.IsNullOrEmpty(szDate))
                                    szDate = GetValueFromKey(multipart.Data, "Date:");
                                if (string.IsNullOrEmpty(szDate))
                                    szDate = GetValueFromKey(multipart.Data, "Arrival-Date:");
                                if (string.IsNullOrEmpty(szRecipient))
                                    szRecipient = GetEmailFromValue(GetValueFromKey(multipart.Data, "Final-Recipient:"));
                                if (string.IsNullOrEmpty(szAction))
                                    szAction = GetValueFromKey(multipart.Data, "Action:");
                                if (string.IsNullOrEmpty(szStatus))
                                    szStatus = GetValueFromKey(multipart.Data, "Status:");
                                if (string.IsNullOrEmpty(szReply_to))
                                    szReply_to = GetEmailFromValue(GetValueFromKey(multipart.Data, "Reply-To:"));
                            }
                        }
                        message = message + szTo + " | " + szFrom + " | " + szReply_to + " | " + szSubject + " | " + szDate + " <br/>";
                        email.DeleteEmail();
                    }
                }
                message = message + string.Format("Total message count: {0}, Process message count: {1}", email.MessageCount.ToString(), nTotalCount.ToString());
                email.CloseConnection();
            }
            catch (Exception ex)
            {
                if (null != email)
                {
                    try
                    {
                        email.CloseConnection();
                    }
                    catch
                    {
                    }
                }
                message = "Exception thrown in FindAndRecordEmailBounces.  Message: " + ex.Message;
            }
            return message;
        }
    }
}
