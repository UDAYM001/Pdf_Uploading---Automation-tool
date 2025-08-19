using System;
using System.IO;
using System.Windows.Forms;

namespace PdfAutomationApp
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                ApplicationConfiguration.Initialize();
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                // Log to a file in the same folder as the EXE
                File.WriteAllText("error_log.txt", ex.ToString());

                // Optional: Show error message
                MessageBox.Show("Something went wrong: " + ex.Message, "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
