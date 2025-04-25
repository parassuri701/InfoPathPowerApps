namespace DocumentGenerator
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.ApplicationExit += Application_ApplicationExit; ;
            Application.Run(new FrmReadFile());
        }

        private static void Application_ApplicationExit(object? sender, EventArgs e)
        {
            string rootDirectory = AppContext.BaseDirectory;

            try
            {
                // Get all directories containing "staging" in their names (case insensitive)
                var directories = Directory.GetDirectories(rootDirectory, "*staging*", SearchOption.AllDirectories);

                foreach (var directory in directories)
                {
                    try
                    {
                        Directory.Delete(directory, true); // Delete the directory and all its contents
                        Console.WriteLine($"Deleted: {directory}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to delete {directory}: {ex.Message}");
                    }
                }

                Console.WriteLine("All matching directories have been processed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while searching directories: {ex.Message}");
            }
        }
    }
}