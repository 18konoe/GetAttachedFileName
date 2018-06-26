using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace GetAttachedFileName
{
    public partial class Ribbon1
    {
        private List<RibbonButton> buttons = new List<RibbonButton>();
        private Microsoft.Office.Interop.Outlook.MailItem mailItem = null;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            if (mailItem == null)
            {
                var inspector = base.Context as Microsoft.Office.Interop.Outlook.Inspector;
                mailItem = inspector.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;
                //ファイルが添付された時のイベントを登録
                (mailItem as ItemEvents_10_Event).BeforeAttachmentAdd += OnBeforeAttachmentAdd;
                (mailItem as ItemEvents_10_Event).AttachmentRemove += OnAttachmentRemove;

                foreach (Attachment mailItemAttachment in mailItem.Attachments)
                {
                    AddMenuItem(mailItemAttachment.DisplayName);
                }

                if (!HasMenuItem())
                {
                    insertNameMenu.Enabled = false;
                }
            }
        }

        private void OnAttachmentRemove(Attachment attachment)
        {
            RemoveMenuItem(attachment.DisplayName);
        }

        private void OnBeforeAttachmentAdd(Attachment attachment, ref bool cancel)
        {
            if (!cancel)
            {
                AddMenuItem(attachment.DisplayName);
            }
        }

        public void AddMenuItem(string label)
        {
            RibbonButton button = Factory.CreateRibbonButton();
            button.Label = label;
            button.Click += (sender, args) =>
            {
                Microsoft.Office.Interop.Word.Document document = mailItem.GetInspector.WordEditor;
                Microsoft.Office.Interop.Word.Application application = document.Application;
                var selection = application.Selection;
                if (application.Options.Overtype)
                {
                    application.Options.Overtype = false;
                }

                if (selection.Active)
                {
                    selection.TypeText(button.Label);
                }
            };
            insertNameMenu.Items.Add(button);

            if (!insertNameMenu.Enabled && HasMenuItem())
            {
                insertNameMenu.Enabled = true;
            }
        }

        public void RemoveMenuItem(string label)
        {
            var button = insertNameMenu.Items.FirstOrDefault(control => (control as RibbonButton).Label == label);
            if (button != null)
            {
                insertNameMenu.Items.Remove(button);
                button.Dispose();
                button = null;
            }

            if (insertNameMenu.Enabled && !HasMenuItem())
            {
                insertNameMenu.Enabled = false;
            }
        }

        public bool HasMenuItem()
        {
            return insertNameMenu.Items.Count != 0;
        }
    }
}
