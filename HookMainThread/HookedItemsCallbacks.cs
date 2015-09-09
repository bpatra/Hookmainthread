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
            Log.Info("Start hook processing in {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
            var results = await GetMattersAsync(subject);
            Log.Info("Finishing in {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
            foreach (var result in results)
            {
                Log.Info("found result {0}", result);
            }
        }

        private async Task<List<string>> GetMattersAsync(string subject)
        {
            return await Task.Run( () =>
            {
                Log.Info("CPU intensive job in {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
                Thread.Sleep(TimeSpan.FromSeconds(20));
                int length  = subject == null ? 0 : subject.Length;
                return new[]{"toto" + length,"tata" + length, "tutu" +length}.ToList();
            });
        }
    }
}
