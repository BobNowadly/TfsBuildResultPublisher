using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace TfsBuildResultPublisher
{
    public interface ITestRunPublisher
    {
        bool PublishTestRun(Configuration configuration);
    }

    public class TestRunPublisher : ITestRunPublisher
    {
        public bool PublishTestRun(Configuration configuration)
        {
            if (configuration.TestSuiteId == null && !configuration.TryAllSuites)
                throw new ArgumentException("/testSuiteId must be specified when publishing test results");
            if (configuration.TestConfigId == null)
                throw new ArgumentException("/testConfigId must be specified when publishing test results");

            string trxPath = Path.Combine(
                Path.GetDirectoryName(configuration.TestResults),
                Path.GetFileNameWithoutExtension(configuration.TestResults) + "_TestRunPublish.trx");

            Console.WriteLine("Taken copy of results file to update for publish ({0})", trxPath);
            File.Copy(configuration.TestResults, trxPath);

            if (configuration.FixTestIds)
                TrxFileCorrector.FixTestIdsInTrx(trxPath);

            ChangeTestIdAndRewriteToDisk(trxPath);

            var paths = new[]
            {
                @"C:\Program Files (x86)\Microsoft Visual Studio 11.0\Common7\IDE\TCM.exe",
                @"C:\Program Files\Microsoft Visual Studio 11.0\Common7\IDE\TCM.exe"
            };
            var tcm = File.Exists(paths[0]) ? paths[0] : paths[1];
            var suiteIds = new List<int>();
            if (configuration.TryAllSuites)
            {
                suiteIds = FetchAllSuiteIdsForProject(configuration, tcm);
            }
            else
            {
                suiteIds.Add(configuration.TestSuiteId.Value);
            }

            var exitCode = 0;
            foreach (var suiteId in suiteIds)
            {
                Console.WriteLine("Processing Suite ({0})", suiteId);

                // we need to update the Test Id for each suite to Trick MTM to thinking it is a new Trx file
                ChangeTestIdAndRewriteToDisk(trxPath);

                const string argsFormat = "run /publish /suiteid:{0} /configid:{1} " +
                                          "/resultsfile:\"{2}\" " +
                                          "/collection:\"{3}\" /teamproject:\"{4}\" " +
                                          "/build:\"{5}\" /builddefinition:\"{6}\" /resultowner:\"{7}\"";

                var args = string.Format(argsFormat, suiteId, configuration.TestConfigId, trxPath,
                    configuration.Collection, configuration.Project, configuration.BuildNumber,
                    configuration.BuildDefinition, configuration.TestRunResultOwner ?? Environment.UserName);

                //Optionally override title
                if (!string.IsNullOrEmpty(configuration.TestRunTitle))
                    args += " /title:\"" + configuration.TestRunTitle + "\"";

                Console.WriteLine("Launching tcm.exe {0}", args);

                string stdOut;
                string stdErr;
                var processStartInfo = new ProcessStartInfo(tcm, args)
                {
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true
                };
                var process = Process.Start(processStartInfo);                

                process.InputAndOutputToEnd(string.Empty, out stdOut, out stdErr);

                Console.Write(stdOut);
                Console.Write(stdErr);

                // Since we are giving this a try we will not return an error code if 
                // tests are not found
                if (!configuration.TryAllSuites)
                {
                    exitCode += process.ExitCode;
                }
            }

            return exitCode == 0;
        }

        private List<int> FetchAllSuiteIdsForProject(Configuration configuration, string tcm)
        {
            string stdOut;
            string stdErr;

            const string argsFormat = "suites /list /collection:\"{0}\" /teamproject:\"{1}\" ";

            var args = string.Format(argsFormat, configuration.Collection, configuration.Project);

            var processStartInfo = new ProcessStartInfo(tcm, args)
            {
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true
            };
            var process = Process.Start(processStartInfo);

            process.InputAndOutputToEnd(string.Empty, out stdOut, out stdErr);
            var output = stdOut.Split('\n');

            var suiteIdsString = output.Select(s => Regex.Replace(s, "^(\\d{1,99}).*", "$1"));
            var suiteIdsStringInts = suiteIdsString.Where(i => Regex.IsMatch(i,"^(\\d{1,99})$"));
            var suiteIds = suiteIdsStringInts.Select(r => Convert.ToInt32(r));

            return suiteIds.ToList();
        }

        private void ChangeTestIdAndRewriteToDisk(string trxPath)
        {
            var fixedFile = TrxFileCorrector.ChangeTestRunId(trxPath);

            File.WriteAllText(trxPath, fixedFile);
        }
    }
}