using Microsoft.Office.Interop.Outlook;
using System.Timers;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ContactListFixer
{
    /*
     
    Dim obApp As Application
    Dim olContacts As Outlook.Items
    Dim obj As Object
    Dim oContact As Outlook.ContactItem
    Dim strName As String
    
    Set olContacts = Session.GetDefaultFolder(olFolderContacts).Items
    
    For Each obj In olContacts
        If TypeName(obj) = "ContactItem" Then
            Set oContact = obj
            With oContact
                strName = .LastNameAndFirstName & " (" & .Email1Address & ")"
                .Email1DisplayName = strName
                
                If Not .Email2Address = "" Then
                    strName = .LastNameAndFirstName & " (" & .Email2Address & ")"
                    .Email2DisplayName = strName
                End If
                
                If Not .Email3Address = "" Then
                    strName = .LastNameAndFirstName & " (" & .Email3Address & ")"
                    .Email3DisplayName = strName
                End If
                
                .Save
            End With
        End If
    Next
     
     */

    public partial class ContactListFixer
    {
        private System.Timers.Timer timer;

        private void InternalStartup()
        {
            Application.Startup += Application_Startup;
        }

        private void Application_Startup()
        {
            timer = new System.Timers.Timer(10000);
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Stop();
            var contactsItems = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Items;
            foreach (var obj in contactsItems)
            {
                if (obj is Outlook.ContactItem contact)
                    ChangeContactDisplayName(contact);
            }
            timer.Start();
        }

        private void ChangeContactDisplayName(ContactItem contact)
        {
            var strName = contact.LastNameAndFirstName + " (" + contact.Email1Address + ")";
            contact.Email1DisplayName = strName;

            if (contact.Email2Address != string.Empty)
            {
                strName = contact.LastNameAndFirstName + " (" + contact.Email2Address + ")";
                contact.Email2DisplayName = strName;
            }

            if (contact.Email3Address != string.Empty)
            {
                strName = contact.LastNameAndFirstName + " (" + contact.Email3Address + ")";
                contact.Email3DisplayName = strName;
            }

            contact.Save();
        }
    }
}
