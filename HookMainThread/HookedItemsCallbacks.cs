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

        public void TryProcessInboxMailAsync(MailItem mailItem)
        {
            string subject = mailItem.Subject;
            Log.Info("Start hook processing in {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
            var results = GetMattersAsync(subject);
            results.ContinueWith((tsk) =>
            {
                Log.Info("Finishing in {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
                foreach (var result in tsk.Result)
                {
                    Log.Info("found result {0}", result);
                }
            },Globals.ThisAddIn.TaskScheduler);

        }

        private Task<List<string>> GetMattersAsync(string subject)
        {
            return Task.Factory.StartNew( () =>
            {
                Log.Info("CPU intensive job in {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
                Thread.Sleep(TimeSpan.FromSeconds(20));
                int length  = subject == null ? 0 : subject.Length;
                return new[]{"toto" + length,"tata" + length, "tutu" +length}.ToList();
            });
        }
    }
}
