using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendOutlookAppointment.Data
{
    public class AppointmentService
    {
        public void AddAppointment()
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.AppointmentItem agendaMeeting = (Outlook.AppointmentItem)
         (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.
         olAppointmentItem);

                if (agendaMeeting != null)
                {
                    agendaMeeting.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                    agendaMeeting.Location = "New York";
                    agendaMeeting.Subject = "Halloween";
                    agendaMeeting.Body = "SPOOKY HALLOWEEN";
                    agendaMeeting.Start = new DateTime(2024, 10, 31, 22, 0, 0);
                    agendaMeeting.Duration = 60;
                    Outlook.Recipient recipient =
                        agendaMeeting.Recipients.Add("Jane Doe");
                    recipient.Type =
                        (int)Outlook.OlMeetingRecipientType.olRequired;
                    ((Outlook._AppointmentItem)agendaMeeting).Send();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("The following error occurred: " + ex.Message);
            }
        }
    }
}
