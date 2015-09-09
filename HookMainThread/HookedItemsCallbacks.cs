using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace BonenLawyer
{

    public class HookedItemsCallbacks
    {

        public async void TryProcessInboxMailAsync(MailItem mailItem)
        {
            string subject = mailItem.Subject;
            var results = await GetMattersAsync(subject);
            foreach (var result in results)
            {
                Log.Info("found result {0}", result);
            }
        }

        private async Task<List<string>> GetMattersAsync(string subject)
        {
            return await Task.Run( () =>
            {
                Thread.Sleep(TimeSpan.FromSeconds(20));
                int length  = subject == null ? 0 : subject.Length;
                return new[]{"toto" + length,"tata" + length, "tutu" +length}.ToList();
            });
        }
    }
}
