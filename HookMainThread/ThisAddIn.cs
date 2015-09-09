using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace BonenLawyer
{
    public partial class ThisAddIn
    {
        //TODO: review the way we handle Singletons: HookedFolderItems, TopLevelAcceptableMailFolderCache ,WebStatus etc.
        public static object MissingType = null; //use this instead of direct Type.Missing to allow unit testing.
        public HookedFolderItems HookedFolderItems; //maintain a global ref to the object to avoid GC and then loose the event hooking


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
                InitializeAddin();
          
        }

        private void InitializeAddin()
        {
            MissingType = System.Type.Missing;
            Log.InitLog(@"D:/hookeditem.log");
            Log.Info("*******START Addin*********");
            Log.Info("Main thread VSTA is {0} {1}", Thread.CurrentThread.Name, Thread.CurrentThread.ManagedThreadId);
            HookedItemsCallbacks callbacks = new HookedItemsCallbacks();
            HookedFolderItems = new HookedFolderItems(callbacks);
            HookedFolderItems.Initialize();
        }
    
        private void ThisAddIn_Quit()
        {
            Log.Info("*******Quit Addin*********");
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }

        #endregion
    }
}
