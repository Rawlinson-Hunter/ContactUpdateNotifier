using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using RhLibs.Sql;

namespace ContactUpdateNotifier
{
    internal class Program
    {
        private class Ret
        {
            public int Count { get; set; }
            public string html { get; set; }
        }

        private static void Main(string[] args)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
            {
                Credentials = new NetworkCredential("EWSContactsServAcc", "Nerada!", "RH"),
                Url = new Uri("https://mail.rawlinson-hunter.com/ews/exchange.asmx"),
                TraceEnabled = false,
                TraceFlags = TraceFlags.All
            };

            var partners = new List<Partner>();

            foreach (DataRow dr in Datalayer.Partners().Rows)
            {
                Console.WriteLine("Adding Partner {0}", Convert.ToString(dr["username"]));

                partners.Add(new Partner
                {
                    Inits = Convert.ToString(dr["username"])
                });
            }

            foreach (DataRow dr in Datalayer.ListOfChangesForSecAndPartner().Rows)
            {
                foreach (Partner p in partners)
                {
                    try
                    {
                        bool inc = Convert.ToBoolean(dr[p.Inits]);

                        if (!inc) continue;

                        DataTable dt = Datalayer.GetAddresses(p.Inits);

                        p.PartnerAddress = Convert.ToString(dt.Rows[0]["partnersmtp"]);
                        p.SecretaryAddress = Convert.ToString(dt.Rows[0]["secsmtp"]);

                        Console.WriteLine("Adding {0} to {1}", Convert.ToString(dr["contname"]), p.Inits);

                        p.AddEntry(new Entry
                        {
                            ActionDate = Convert.ToDateTime(dr["actiondate"]),
                            Narrative = Convert.ToString(dr["narrative"]),
                            Name = Convert.ToString(dr["contname"]),
                            ID = Convert.ToInt32(dr["idx"]),
                            Contindex = Convert.ToInt32(dr["contindex"]),
                            Reason = Convert.ToString(dr["contnotes"]),
                            Company = Convert.ToString(dr["companyname"]),
                            Staff = Convert.ToString(dr["fullname"])
                        });
                    }
                    catch
                    {
                    }
                }
            }

            List<Entry> secchanges = new List<Entry>();

            foreach (DataRow dr in Datalayer.ListOfChangesForDBA().Rows)
            {

                try
                {
                    secchanges.Add(new Entry
                    {
                        ActionDate = Convert.ToDateTime(dr["actiondate"]),
                        Narrative = Convert.ToString(dr["narrative"]),
                        Name = Convert.ToString(dr["contname"]),
                        ID = Convert.ToInt32(dr["idx"]),
                        Contindex = Convert.ToInt32(dr["contindex"]),
                        Reason = Convert.ToString(dr["contnotes"]),
                        Company = Convert.ToString(dr["companyname"]),
                        Staff = Convert.ToString(dr["fullname"]),
                        Owner = Convert.ToString(dr["staffuser"])
                    });
                }
                catch
                {
                }
            }

            SendEmailForSecAndPartner(partners, service);

            if (secchanges.Count > 0)
                SendEmailForDBA(secchanges, service);

            JRSDebug(partners, service);
        }

        private static void InsertNotify(Partner partner)
        {
            foreach (Entry e in partner.Entries)
            {
                Datalayer.Insert(e.ID);
            }
        }

        private static Ret GetBodyForSecAndPartner(Partner partner, bool showinits = false)
        {
            var sb = new StringBuilder();
            const string url = "http://intranet/sites/intranet/search/Pages/contactresults.aspx#Default=";

            if (showinits) sb.Append(partner.Inits);

            sb.Append("<table>");
            sb.Append("<tr class=\"header\">");
            sb.Append("<td>");
            sb.Append("<span>Action Date</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Who</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Contact Name</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Company Name</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Action</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Reason</span>");
            sb.Append("</td>");
            sb.Append("</tr>");

            int a = 0;

            foreach (Entry entry in partner.Entries)
            {
                if (entry.Ignore) continue;

                sb.AppendFormat("<tr class=\"{0}\">", a % 2 == 0 ? "alt" : "");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.ActionDate.ToString("dd/MM/yyyy HH:mm"));
                sb.Append("</span>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.Staff);
                sb.Append("</span>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.AppendFormat("<a href=\"{0}{1}\">", url, WebUtility.UrlEncode(string.Concat("{\"k\":\"", entry.Name, "\"}")));
                sb.Append(entry.Name);
                sb.Append("</a>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.AppendFormat("<a href=\"{0}{1}\">", url, WebUtility.UrlEncode(string.Concat("{\"k\":\"", entry.Company, "\"}")));
                sb.Append(entry.Company);
                sb.Append("</a>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.FixedNarrative);
                sb.Append("</span>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.FixedReason);
                sb.Append("</span>");

                sb.Append("</td>");
                sb.Append("</tr>");

                a++;
            }

            sb.Append("</table>");

            return new Ret { Count = a, html = sb.ToString() };
        }

        private static string GetBodyForDBA(List<Entry> entries)
        {
            var sb = new StringBuilder();
            const string url = "http://intranet/sites/intranet/search/Pages/contactresults.aspx#Default=";

            sb.Append("<table>");
            sb.Append("<tr class=\"header\">");
            sb.Append("<td>");
            sb.Append("<span>Action Date</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Who</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Contact Name</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Company Name</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Action</span>");
            sb.Append("</td>");
            sb.Append("<td>");
            sb.Append("<span>Owner</span>");
            sb.Append("</td>");
            sb.Append("</tr>");

            int a = 0;

            foreach (Entry entry in entries)
            {
                if (entry.Ignore) continue;

                sb.AppendFormat("<tr class=\"{0}\">", a % 2 == 0 ? "alt" : "");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.ActionDate.ToString("dd/MM/yyyy HH:mm"));
                sb.Append("</span>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.Staff);
                sb.Append("</span>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.AppendFormat("<a href=\"{0}{1}\">", url, WebUtility.UrlEncode(string.Concat("{\"k\":\"", entry.Name, "\"}")));
                sb.Append(entry.Name);
                sb.Append("</a>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.AppendFormat("<a href=\"{0}{1}\">", url, WebUtility.UrlEncode(string.Concat("{\"k\":\"", entry.Company, "\"}")));
                sb.Append(entry.Company);
                sb.Append("</a>");

                sb.Append("</td>");
                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.FixedNarrative);
                sb.Append("</span>");

                sb.Append("</td>");

                sb.Append("<td>");

                sb.Append("<span>");
                sb.Append(entry.Owner);
                sb.Append("</span>");

                sb.Append("</td>");

                sb.Append("</tr>");

                a++;
            }

            sb.Append("</table>");

            return sb.ToString();
        }

        private static void SendEmailForSecAndPartner(List<Partner> partners, ExchangeService service)
        {
            service.HttpHeaders.Remove("X-AnchorMailbox");
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Properties.Settings.Default.dbaemail);
            service.HttpHeaders.Add("X-AnchorMailbox", Properties.Settings.Default.dbaemail);

            foreach (Partner partner in partners)
            {
                if (partner.Entries.Count == 0) continue;

                var email = new EmailMessage(service);

                //email.ToRecipients.Add("bevan.johnson@rawlinson-hunter.com");

                email.CcRecipients.Add(partner.PartnerAddress);
                email.ToRecipients.Add(partner.SecretaryAddress);

                email.Subject = "Contact change notification";

                StringBuilder body = new StringBuilder();
                body.Append(Properties.Resources.Style);
                body.Append(GetBodyForSecAndPartner(partner).html);

                email.Body = new MessageBody(BodyType.HTML, body.ToString());

                email.Send();

                InsertNotify(partner);
            }
        }

        private static void JRSDebug(List<Partner> partners, ExchangeService service)
        {
            StringBuilder sb = new StringBuilder();

            service.HttpHeaders.Remove("X-AnchorMailbox");
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Properties.Settings.Default.dbaemail);
            service.HttpHeaders.Add("X-AnchorMailbox", Properties.Settings.Default.dbaemail);

            sb.Append(Properties.Resources.Style);

            bool cont = false;

            foreach (Partner partner in partners)
            {
                if (partner.Entries.Count == 0) continue;
                Ret r = GetBodyForSecAndPartner(partner, true);
                if (r.Count > 0)
                {
                    sb.AppendLine(r.html);
                    InsertNotify(partner);
                    cont = true;
                }
            }

            if (!cont) return;

            var email = new EmailMessage(service);

            email.ToRecipients.Add("james.symonds@rawlinson-hunter.com");
            //email.ToRecipients.Add("bevan.johnson@rawlinson-hunter.com");

            email.Subject = "Contact change notification";
            email.Body = new MessageBody(BodyType.HTML, sb.ToString());
            email.Send();
        }

        private static void SendEmailForDBA(List<Entry> secchanges, ExchangeService service)
        {
            service.HttpHeaders.Remove("X-AnchorMailbox");
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "helpdesk@rawlinson-hunter.com");
            service.HttpHeaders.Add("X-AnchorMailbox", "helpdesk@rawlinson-hunter.com");

            var email = new EmailMessage(service);

            email.ToRecipients.Add(Properties.Settings.Default.dbaemail);
            //email.ToRecipients.Add("bevan.johnson@rawlinson-hunter.com");

            email.Subject = "Contact change notification";

            StringBuilder body = new StringBuilder();
            body.Append(Properties.Resources.Style);
            body.Append(GetBodyForDBA(secchanges));

            email.Body = new MessageBody(BodyType.HTML, body.ToString());

            email.Send();

            foreach (var entry in secchanges)
            {
                Datalayer.Insert(entry.ID);
            }
        }
    }

    public static class Datalayer
    {
        private static readonly string _conn = SqlHelper.GetConnection("CustomPE");

        public static DataTable GetAddresses(string partner)
        {
            var p = new SqlParameter("@partner", SqlDbType.VarChar) { Value = partner };
            return SqlHelper.ExecuteDataset(_conn, CommandType.StoredProcedure, "[contacteditnotification].getaddresses", p).Tables[0];
        }

        public static void Insert(int id)
        {
            var p = new SqlParameter("@id", SqlDbType.Int) { Value = id };
            SqlHelper.ExecuteNonQuery(_conn, CommandType.StoredProcedure, "[contacteditnotification].insertnotified", p);
        }

        public static DataTable ListOfChangesForSecAndPartner()
        {
            return SqlHelper.ExecuteDataset(_conn, CommandType.StoredProcedure, "[contacteditnotification].ListOfChangesForSecAndPartner").Tables[0];
        }

        public static DataTable ListOfChangesForDBA()
        {
            return SqlHelper.ExecuteDataset(_conn, CommandType.StoredProcedure, "[contacteditnotification].ListOfChangesForDBA").Tables[0];
        }

        public static DataTable Partners()
        {
            return SqlHelper.ExecuteDataset(_conn, CommandType.StoredProcedure, "[contacteditnotification].partners").Tables[0];
        }
    }

    public class Partner
    {
        public Partner()
        {
            Entries = new List<Entry>();
        }

        public string Inits { get; set; }
        public string PartnerAddress { get; set; }
        public string SecretaryAddress { get; set; }
        public List<Entry> Entries { get; set; }

        public void AddEntry(Entry e)
        {
            foreach (Entry x in Entries)
            {
                if (x.Exists(e)) return;
            }

            Entries.Add(e);
        }
    }

    public class Entry
    {
        public int ID { get; set; }
        public DateTime ActionDate { get; set; }
        public string Narrative { get; set; }
        public string Name { get; set; }
        public int Contindex { get; set; }
        public string Reason { get; set; }
        public string Company { get; set; }
        public string Staff { get; set; }
        public string Owner { get; set; }

        public string FixedReason
        {
            get
            {

                int start = Reason.IndexOf("#");

                if (start == -1) return string.Empty;

                int end = Reason.LastIndexOf("#");

                return end - start == 0 ? Reason : Reason.Substring(start + 1, end - start - 1);
            }
        }

        public bool Ignore
        {
            get
            {
                int start = Reason.IndexOf("#");
                int end = Reason.LastIndexOf("#");
                return end - start == 1;
            }
        }

        public string FixedNarrative
        {
            get { return Narrative.Replace("  ", " *Blank* "); }
        }

        public bool Exists(Entry e)
        {
            return Convert.ToInt16(Math.Abs(e.ActionDate.Subtract(ActionDate).TotalHours)) < 2 && e.Narrative == Narrative && e.Contindex == Contindex;
        }
    }
}