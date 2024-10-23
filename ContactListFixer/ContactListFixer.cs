using Microsoft.Office.Interop.Outlook;
using System.Timers;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ContactListFixer
{
    /*
     * ContactListFixer
     * 
     * Небольшой плагин для Outlook, по таймеру меняющий "DisplayName" контактов.
     * Проблема в том, что Outlook как-то загадочно хранит контакты и искать их по фамилии
     * (а создавая письмо рефлекторно начинается писать фамилию) просто невозможно, потому что
     * поиск идёт по полю "Хранить как" для первого адреса электронной почты.
     * Этот аддон кадые 60 секунд перетряхивает контакты и переделывает нужное поле, ремонтируя
     * тем самым нормальный поиск.
     * 
     * Автор: Стефаняк Антон Юрьевич / Ant0nRocket / anton.stephanyak@ya.ru
     * 
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
            timer = new System.Timers.Timer(60000);
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Stop(); // остановим таймер на время работы с контактами

            var contactsItems = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Items;
            foreach (var obj in contactsItems)
            {
                if (obj is Outlook.ContactItem contact)
                    ChangeContactDisplayName(contact);
            }

            timer.Start(); // всё сделано, запускаем таймер
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
