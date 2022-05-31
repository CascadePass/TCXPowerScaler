using System;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace TcxPowerScaler
{
    /// <summary>
    /// The entire TCX Power Scaler console application.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Gets or sets the scale factor, eg the amount to adjust the power values in the file.
        /// </summary>
        public static double ScaleFactor { get; set; }

        /// <summary>
        /// Gets or sets the folder with TCX files to scale.
        /// </summary>
        public static string WorkingFolder { get; set; }

        /// <summary>
        /// The main entry point for the TCX Power Scaler software.
        /// </summary>
        /// <param name="args">Optional command-line arguments.</param>
        static void Main(string[] args)
        {
            Program.ParseArguments(args);
            if (!Program.CheckValues())
            {
                return;
            }

            Program.ProcessFolder();

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        /// <summary>
        /// Processes all the TCX files in a folder.
        /// </summary>
        public static void ProcessFolder()
        {
            if(!string.IsNullOrWhiteSpace(Program.WorkingFolder) && Directory.Exists(Program.WorkingFolder))
            {
                foreach (string filename in Directory.GetFiles(Program.WorkingFolder, "*.tcx"))
                {
                    Program.UpdateFile(filename);
                }
            }
        }

        /// <summary>
        /// Updates the power values in a file and preserves the original data.
        /// </summary>
        /// <param name="filename">The name of the file to process.</param>
        public static void UpdateFile(string filename)
        {
            Console.WriteLine(filename);

            XmlDocument data = new XmlDocument();
            data.Load(filename);

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(data.NameTable);
            nsmgr.AddNamespace("ns3", "http://www.garmin.com/xmlschemas/ActivityExtension/v2");

            var toUpdate = data.SelectNodes(@"//ns3:Watts", nsmgr);

            foreach (XmlNode node in toUpdate)
            {
                if (double.TryParse(node.InnerText, out double rawOriginalPower))
                {
                    double newPower = Math.Round(rawOriginalPower * Program.ScaleFactor);
                    node.InnerText = newPower.ToString();

                    if(Debugger.IsAttached)
                    {
                        Console.WriteLine($"{rawOriginalPower} -> {newPower}");
                    }
                }
            }

            string newFilename = $"{filename}.original";

            if (File.Exists(newFilename))
            {
                newFilename = $"{newFilename}.{Guid.NewGuid().ToString().Replace("-", string.Empty)}";
            }

            File.Copy(filename, newFilename);
            data.Save(filename);
        }

        /// <summary>
        /// Parses command line arguments.
        /// </summary>
        /// <param name="args">A <see cref="string"/> array of command line arguments.</param>
        public static void ParseArguments(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                return;
            }

            foreach (string arg in args)
            {
                if(arg.StartsWith("scale:", StringComparison.OrdinalIgnoreCase))
                {
                    string scaleValueText = arg.Substring(5);

                    if (Double.TryParse(scaleValueText, out double d))
                    {
                        Program.ScaleFactor = d;
                    }
                }

                if (arg.StartsWith("folder:", StringComparison.OrdinalIgnoreCase))
                {
                    string folder = arg.Substring(6);
                    if (arg.StartsWith("\"") && arg.EndsWith("\""))
                    {
                        folder = folder.Substring(1, folder.Length - 2);
                    }

                    Program.WorkingFolder = folder;
                }
            }
        }

        /// <summary>
        /// Assigns values to any properties that were not passed as command line arguments.
        /// </summary>
        /// <returns>A <see cref="bool"/> value indicating whether the program can continue.</returns>
        public static bool CheckValues()
        {
            if (string.IsNullOrWhiteSpace(Program.WorkingFolder))
            {
                Program.WorkingFolder = Environment.CurrentDirectory;
            }

            while (Program.ScaleFactor == 0)
            {
                Console.WriteLine("Enter scaling factor");

                string userInput = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(userInput))
                {
                    return false;
                }

                if (Double.TryParse(userInput, out double d))
                {
                    string key = null;
                    while(key != "Y" && key != "N")
                    {
                        Console.WriteLine($"Power will be adjusted by {d * 100}% - is this correct?  (Y/N)");
                        key = Console.ReadLine().ToString().ToUpper();
                    };

                    Program.ScaleFactor = d;
                }
            }

            return true;
        }
    }

}
