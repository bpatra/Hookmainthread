using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Outlook;

using Exception = System.Exception;

namespace BonenLawyer
{
    public class HookedFolderItems
    {
        private HookedItemsCallbacks _hookedItemsCallbacks;
        public HookedFolderItems( HookedItemsCallbacks hookedItemsCallbacks)
        {
            if (hookedItemsCallbacks == null) { throw new ArgumentNullException("hookedItemsCallbacks"); }
            _hookedItemsCallbacks = hookedItemsCallbacks;
        }


         public Items InboxItems { get; private set; }
        
        private void CleanRegisteredCallBacks()
        {
            if (InboxItems != null)
            {
                InboxItems.ItemAdd -= AddNewInboxItems;
            }        }

        public void Initialize()
        {
            CleanRegisteredCallBacks();

            InboxItems = Globals.ThisAddIn.Application.Session.DefaultStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                    .Items;
            InboxItems.ItemAdd += AddNewInboxItems;
        }

        private void AddNewInboxItems(object item)
        {
            try
            {
                MailItem mailItem = item as MailItem;
                if (mailItem != null)
                {
                    _hookedItemsCallbacks.TryProcessInboxMailAsync(mailItem);
                }
            }
            catch (Exception ex)
            {
                Log.Exception(ex);
                Log.Info("Failure in AddNewInboxItems Hook");
            }
        }

    }
}
