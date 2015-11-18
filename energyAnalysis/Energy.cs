using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;

namespace energyAnalysis
{
    public class Energy
    {
        private static Logger logger = new Logger("EnergyAnalysis");

        private string epPath;

        public Energy(string epPath)
        {
            this.epPath = epPath;
        }

        private void tryToKill(Process process)
        {
            try
            {
                logger.log("Closing process '" + process.ProcessName + "' with id " + process.Id);
                process.Kill();
                process.WaitForExit();
            }
            catch (Exception ex)
            {
                logger.log(ex.Message, Logger.LogType.ERROR);
            }
        }
        public bool Simulate(string inFile,string weatherFile, string outFile)
        {
            Process[] processes;
            try
            {
                bool isDirectory = Directory.Exists(outFile);
                bool result = false;
                Process process;
                string tmpFile;
                string appPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\EnergyAnalysis";
                string tmpFolder;
                if (!Directory.Exists(appPath))
                {
                    Directory.CreateDirectory(appPath);
                }


                if (isDirectory)
                {
                    outFile = outFile + @"\" + Path.ChangeExtension(Path.GetFileName(inFile), "zip");
                }
                string outFileName = Path.GetFileName(outFile);
                tmpFolder = appPath +@"\"+outFileName;
                if(!Directory.Exists(tmpFolder))
                {
                    Directory.CreateDirectory(tmpFolder);
                }
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.WorkingDirectory = this.epPath;
                startInfo.FileName = this.epPath + @"\energyplus.exe";
                startInfo.Arguments = "-w " + weatherFile + " -d " + tmpFolder + " " + inFile;
                process = Process.Start(startInfo);
                logger.log("Process '" + process.ProcessName + "' with id " + process.Id + " started");
                process.WaitForExit();
                logger.log("Closing process '" + process.ProcessName + "' with id " + process.Id);
                tmpFile = tmpFolder + @"\eplustbl.htm";
                if (File.Exists(tmpFile))
                {
                    logger.log("Generating zip file...");
                    ZipFile.CreateFromDirectory(tmpFolder, outFile);
                    result = true;
                }
                Directory.Delete(tmpFolder, true);
                return result;
            }
            catch(Exception e)
            {
                logger.log(e.Message, Logger.LogType.ERROR);
                processes = Process.GetProcessesByName("energyplus");
                for (var i = 0; i < processes.Length;i++ )
                {
                    tryToKill(processes[i]);
                }
                return false;
            }
        }

    }
}
