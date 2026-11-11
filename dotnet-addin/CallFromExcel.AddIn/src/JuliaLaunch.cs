using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;

namespace CallFromExcel.AddIn
{
    // C# port of JuliaLaunch (modMain.bas). The flag-file/window-title readiness protocol,
    // the generated Julia startup script, and the stale-port/stale-window checks are ported
    // as closely as .NET's own APIs allow -- see inline notes for the handful of places
    // .NET's Process/ProcessStartInfo don't have an exact equivalent of WshShell.Run's
    // behavior. WSL support (bash statements, wsl.exe, Windows<->WSL path translation) is
    // ported too, below.
    public static class JuliaLaunch
    {
        private const string PackageName = "CallFromExcel";
        private const string WSLExecutable = @"C:\Windows\System32\wsl.exe";
        private const string ConhostExecutable = @"C:\Windows\System32\conhost.exe";

        // Julia's short options that take a value, either attached (-Lfile) or as the next
        // argument (-L file), per julia --help. Any other letter in a short-option group is a flag
        // (or, for -O and -g, an optional attached value, which a flag scan handles harmlessly).
        private const string ShortOptionsWithValue = "eEHLJCtpm";

        // Checks the user's CommandLineOptions and adds -i and --threads=auto,1 unless the user
        // gave their own. Looks at the actual options - each argument, quotes respected - not at
        // substrings of the whole string, so that e.g. --project=C:/Users/Jean-Luc is not taken
        // for -L, nor --project=C:/multi-index for -i. Public so the tests can call it directly.
        public static string NormaliseCommandLineOptions(string commandLineOptions)
        {
            bool hasLoad = false, hasInteractive = false, hasThreads = false;
            List<string> args = SplitArguments(commandLineOptions ?? "");
            for (int k = 0; k < args.Count; k++)
            {
                string arg = args[k];
                if (arg.StartsWith("--", StringComparison.Ordinal))
                {
                    string name = arg.Split('=')[0];
                    if (name == "--load") hasLoad = true;
                    else if (name == "--threads") hasThreads = true;
                }
                else if (arg.Length > 1 && arg[0] == '-')
                {
                    for (int c = 1; c < arg.Length; c++)
                    {
                        char option = arg[c];
                        if (option == 'L') hasLoad = true;
                        else if (option == 'i') hasInteractive = true;
                        else if (option == 't') hasThreads = true;
                        if (ShortOptionsWithValue.IndexOf(option) >= 0)
                        {
                            if (c == arg.Length - 1) k++;   // the value is the next argument
                            break;                          // the rest of this one is the value
                        }
                    }
                }
            }

            if (hasLoad)
            {
                throw new ArgumentException(
                    "CommandLineOptions cannot include the -L or --load options. Instead use JuliaLaunch " +
                    "without that option and then use JuliaCall(\"include\",\"path_to_file\")");
            }
            string result = commandLineOptions.Trim();
            if (!hasInteractive)
            {
                // Lets OhMyREPL work correctly in the REPL, which requires -i (interactive).
                // https://github.com/KristofferC/OhMyREPL.jl/issues/271
                result = (result + " -i").Trim();
            }
            if (!hasThreads)
            {
                // Without a dedicated interactive thread, HTTP.jl's server tasks share
                // Julia's single default thread with whatever expression is currently
                // being evaluated - a long-running or blocking evaluation (even exit())
                // can then leave the server unable to respond, or even to shut down
                // cleanly. See the "Server" docstring in HTTP.jl's http_server.jl.
                result = (result + " --threads=auto,1").Trim();
            }
            return result;
        }

        // The --project option for a project folder, ready to put in CommandLineOptions; "" for
        // no folder. On Windows, forward slashes and double quotes, so spaces work and a trailing
        // backslash can't escape the closing quote. Under WSL, the folder as a Linux path in
        // bash single quotes (WslArguments' outer double quotes rule out double quotes here): a
        // Windows drive path becomes /mnt/<drive>/..., a \\wsl$\<distro>\ or
        // \\wsl.localhost\<distro>\ path becomes the path inside the distribution, and anything
        // else, e.g. /home/me/proj or @myenv, is used as it is. Public so the tests can call it.
        public static string ProjectOption(string folder, bool useLinux)
        {
            folder = (folder ?? "").Trim();
            if (folder == "") return "";
            if (!useLinux) return "\"--project=" + folder.Replace('\\', '/') + "\"";

            string linuxPath = folder;
            foreach (string prefix in new[] { @"\\wsl$\", @"\\wsl.localhost\" })
            {
                if (folder.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    int afterDistro = folder.IndexOf('\\', prefix.Length);
                    linuxPath = afterDistro < 0 ? "/" : folder.Substring(afterDistro).Replace('\\', '/');
                }
            }
            if (folder.Length >= 3 && folder[1] == ':' && (folder[2] == '\\' || folder[2] == '/'))
            {
                linuxPath = WSLAddress(folder);
            }
            return "--project=" + BashQuote(linuxPath);
        }

        // Splits a command line into arguments at spaces outside double quotes; quotes are
        // removed, as Windows' own argument parsing does.
        private static List<string> SplitArguments(string commandLine)
        {
            var args = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false, any = false;
            foreach (char ch in commandLine)
            {
                if (ch == '"') { inQuotes = !inQuotes; any = true; }
                else if (char.IsWhiteSpace(ch) && !inQuotes)
                {
                    if (any) args.Add(current.ToString());
                    current.Clear();
                    any = false;
                }
                else { current.Append(ch); any = true; }
            }
            if (any) args.Add(current.ToString());
            return args;
        }

        public static string Launch(
            bool useLinux = false,
            bool minimiseWindow = false,
            string commandLineOptions = "",
            string packages = "",
            string bashStatements = "",
            int timeoutSeconds = 30)
        {
            string juliaExe = JuliaExeLocation.Find(useLinux);

            string userSuppliedCommandLineOptions = commandLineOptions ?? "";
            commandLineOptions = NormaliseCommandLineOptions(userSuppliedCommandLineOptions);

            int pid = Process.GetCurrentProcess().Id;
            string windowPartialTitle = "serving Excel PID " + pid; // must match Julia's CallFromExcel.settitle

            bool foundWindow = WindowUtils.TryGetHandleFromPartialCaption(windowPartialTitle, out IntPtr hwndJulia);
            int existingPort = JuliaPort.ReadJuliaPortFromFile();
            bool isListening = existingPort > 0 && JuliaHttpClient.IsListening(existingPort);

            // A window matching this Excel session can still be found, and still fail to
            // respond over HTTP, for a little while after JuliaEval("exit()") - confirmed
            // (via temporary diagnostics while fixing this) that Julia's own shutdown tears
            // things down in stages: its HTTP server stops responding and its window title
            // goes blank before the OS process itself has actually exited, so checking
            // whether the process has exited doesn't help - it genuinely hasn't yet. Without
            // retrying briefly here, that normal, transient state was indistinguishable from
            // a genuinely stuck/crashed session (e.g. sitting at a "julia>" prompt with a
            // dead HTTP server) and raised a scary "already running but not responding" error
            // for a session that was actually just finishing closing. Only engages in this
            // specific ambiguous state, so it costs nothing in the common case where Julia is
            // either already responding (isListening true) or genuinely not running at all
            // (foundWindow false).
            if (foundWindow && !isListening)
            {
                var retryStopwatch = Stopwatch.StartNew();
                while (foundWindow && !isListening && retryStopwatch.Elapsed.TotalSeconds < 5)
                {
                    Thread.Sleep(250);
                    foundWindow = WindowUtils.TryGetHandleFromPartialCaption(windowPartialTitle, out hwndJulia);
                    existingPort = JuliaPort.ReadJuliaPortFromFile();
                    isListening = existingPort > 0 && JuliaHttpClient.IsListening(existingPort);
                }
            }

            if (isListening && foundWindow && WindowUtils.WindowTitleFromHandle(hwndJulia).Contains(", port " + existingPort))
            {
                JuliaPort.SetJuliaPort(existingPort);
                return "Julia is already running in window \"" + WindowUtils.WindowTitleFromHandle(hwndJulia) + "\"";
            }
            if (isListening)
            {
                // Either no window was found, or the one found doesn't match the currently
                // listening port - most likely a session started by calling
                // CallFromExcel.serve_xl directly (e.g. from a Julia REPL in VS Code), which
                // has no dedicated console window to find. Read from disk rather than the
                // (possibly stale) cache, and refresh that cache, so this reliably
                // discovers a session attached this way instead of launching a redundant one.
                JuliaPort.SetJuliaPort(existingPort);
                return "Julia is already running, listening on port " + existingPort +
                       " with no dedicated console window - most likely a session attached via CallFromExcel.serve_xl";
            }
            if (foundWindow)
            {
                throw new InvalidOperationException(
                    "A Julia session titled \"" + WindowUtils.WindowTitleFromHandle(hwndJulia) + "\" is already running for " +
                    "this Excel session, but it is not responding to HTTP requests. Switch to that window: if it's sitting " +
                    "at a \"julia>\" prompt, its HTTP server has likely crashed - close the window and call JuliaLaunch " +
                    "again. If it's busy running code, either wait for it to finish, or press Ctrl+C to interrupt it, " +
                    "then call JuliaLaunch again.");
            }

            // Not exiting early - clear the cached port so GetJuliaPort can be used below
            // to test whether the connection was correctly established.
            JuliaPort.SetJuliaPort(0);

            string errorFile = Path.Combine(FileUtils.LocalTemp(), "LoadError_" + pid + ".txt");
            if (File.Exists(errorFile)) File.Delete(errorFile);

            string flagFile = Path.Combine(FileUtils.LocalTemp(), "Flag_" + pid + ".txt");
            FileUtils.SaveTextFile(flagFile, "");

            string loadFile = Path.Combine(FileUtils.LocalTemp(), "StartUp_" + pid + ".jl");

            string trimmedOptions = commandLineOptions.Trim();
            string flagFileX, commsFolderX, errorFileX, loadFileX, command;
            string processFileName, processArguments;

            if (useLinux)
            {
                if (!File.Exists(WSLExecutable))
                {
                    throw new InvalidOperationException(
                        "Cannot find the WSL executable at '" + WSLExecutable + "'. Check if the file exists and " +
                        "whether read and execute permissions are set for user '" + Environment.UserName + "'");
                }
                errorFileX = WSLAddress(errorFile);
                flagFileX = WSLAddress(flagFile);
                commsFolderX = WSLAddress(FileUtils.LocalTemp());
                loadFileX = WSLAddress(loadFile);

                // Only need a wrapper .sh script if there are BashStatements (e.g. to set
                // environment variables) to run before launching Julia itself - otherwise
                // the inner command below can invoke Julia directly.
                bool wslLaunchFileNecessary = !string.IsNullOrEmpty(bashStatements);
                string innerCommand;
                if (wslLaunchFileNecessary)
                {
                    string wslLaunchFile = Path.Combine(FileUtils.LocalTemp(), "launchjulia.sh");
                    string wslLaunchFileX = WSLAddress(wslLaunchFile);
                    string wslLaunchFileContents = "#!/bin/bash\n" + bashStatements.TrimEnd('\n') + "\n" +
                        WslJuliaCommand(juliaExe, trimmedOptions, loadFileX);
                    FileUtils.SaveTextFile(wslLaunchFile, wslLaunchFileContents);
                    // Run via bash rather than directly: files written from Windows onto /mnt/c lack execute permission.
                    innerCommand = "bash " + BashQuote(wslLaunchFileX) + " && exit";
                }
                else
                {
                    innerCommand = WslJuliaCommand(juliaExe, trimmedOptions, loadFileX);
                }
                // The full path, not "wsl": ShellExecute resolves a bare name partly against
                // the current directory, which follows Excel's last File > Open folder.
                processFileName = WSLExecutable;
                processArguments = WslArguments(innerCommand);
                command = "wsl " + processArguments;
            }
            else
            {
                flagFileX = flagFile.Replace("\\", "/");
                commsFolderX = FileUtils.LocalTemp().Replace("\\", "/");
                errorFileX = errorFile.Replace("\\", "/");
                loadFileX = loadFile.Replace("\\", "/");
                processFileName = juliaExe;
                processArguments = trimmedOptions + " --load \"" + loadFileX + "\"";
                command = "\"" + juliaExe + "\" " + processArguments;
            }

            // Run in the classic console window (conhost) rather than the user's default terminal.
            // Where that's Windows Terminal (the default on Windows 11), the taskbar would show
            // Terminal's icon, and Julia's window couldn't be found by its title. In conhost the
            // taskbar shows Julia's icon (for a Microsoft Store Julia), Julia sets its title-bar
            // icon itself (set_console_icon, in the start-up script), and MinimiseWindow is
            // respected. Still started via ShellExecute (see below), so networking for a Store
            // Julia is unaffected - checked with a loopback connection to Julia under conhost.
            processArguments = "\"" + processFileName + "\" " + processArguments;
            processFileName = ConhostExecutable;

            // MakeJuliaLiteral wraps its result in quotes (so it can stand alone as a
            // Julia string literal); strip them here since the escaped content is about
            // to be embedded inside another double-quoted println string below, which
            // supplies its own quotes.
            string literalCommand = JuliaLiteral.MakeJuliaLiteral(command);
            literalCommand = literalCommand.Substring(1, literalCommand.Length - 2);

            string loadFileContents = BuildStartupScript(packages, pid, commsFolderX, errorFileX, flagFileX, literalCommand);
            FileUtils.SaveTextFile(loadFile, loadFileContents);

            var stopwatch = Stopwatch.StartNew();

            // UseShellExecute MUST be true here, not false. This was found the hard way:
            // with UseShellExecute=false (a direct CreateProcess call, bypassing the
            // shell), a Julia installed via the Microsoft Store -- an MSIX/UWP package,
            // where "julia.exe" on PATH is an App Execution Alias stub that redirects into
            // the real packaged exe -- launches into a broken state on this machine: its
            // HTTP server reports itself as listening, but loopback TCP connections to it
            // fail even from Windows' own Test-NetConnection (not just from this code),
            // and its console window isn't discoverable via classic FindWindow/GetWindow
            // enumeration either. Both symptoms vanished when switched to
            // UseShellExecute=true, which goes through ShellExecuteEx -- the same
            // activation path Explorer/WshShell.Run use, which evidently matters for an
            // App Execution Alias to activate with correct loopback networking capability.
            // Confirmed independently both ways with Test-NetConnection while the launched
            // process was kept alive -- see this file's own git history for the full
            // investigation. This is plausibly specific to Microsoft Store Julia installs;
            // a standalone-installer Julia may not need this, but UseShellExecute=true
            // works for both, so there's no reason to special-case it.
            var psi = new ProcessStartInfo
            {
                FileName = processFileName,
                Arguments = processArguments,
                UseShellExecute = true,
                WindowStyle = minimiseWindow ? ProcessWindowStyle.Minimized : ProcessWindowStyle.Normal
            };
            Process juliaProcess = Process.Start(psi);

            // The VBA version detects "Julia's console window closed before start-up
            // finished" (e.g. from malformed CommandLineOptions) by counting top-level
            // windows whose caption contains "Julia" -- a necessary workaround since
            // WshShell.Run gives VBA no direct handle to the child process. .NET's
            // Process object gives us that handle directly via HasExited, which is both
            // simpler and (now that UseShellExecute=true gives us a real, correctly
            // activated process) reliable -- confirmed it correctly detects a deliberately
            // bad command-line option's quick exit, and correctly stays false throughout a
            // real successful launch.
            while (File.Exists(flagFile))
            {
                Thread.Sleep(50);
                if (juliaProcess.HasExited)
                {
                    // No exit code in the message: juliaProcess is conhost (see above), which
                    // exits with 0 whatever Julia's own exit code was.
                    string errDescription = "Julia's process exited before start-up finished.";
                    if (userSuppliedCommandLineOptions != "")
                    {
                        errDescription += " Check the CommandLineOptions are valid " +
                                           "(https://docs.julialang.org/en/v1/manual/command-line-interface/#command-line-interface)";
                    }
                    throw new InvalidOperationException(errDescription);
                }
                if (stopwatch.Elapsed.TotalSeconds > timeoutSeconds)
                {
                    return StillStartingMessage(timeoutSeconds);
                }
            }

            int port = JuliaPort.ReadJuliaPortFromFile();
            if (port > 0) JuliaPort.SetJuliaPort(port);
            FileUtils.CleanLocalTemp();

            if (File.Exists(errorFile))
            {
                throw new InvalidOperationException(
                    "Julia launched but encountered an error when executing '" + loadFile + "' the error was: " +
                    FileUtils.ReadTextFile(errorFile));
            }

            // Under WSL the port can take a while to become reachable from Windows after Julia
            // starts listening (WSL2 port forwarding), so wait for it within the same timeout.
            while (port > 0 && !JuliaHttpClient.IsListening(port))
            {
                if (stopwatch.Elapsed.TotalSeconds > timeoutSeconds)
                {
                    return StillStartingMessage(timeoutSeconds);
                }
                Thread.Sleep(100);
            }

            if (JuliaPort.GetJuliaPort() == 0)
            {
                throw new InvalidOperationException(
                    "Failed to establish connection between Julia and Excel. Is CallFromExcel installed correctly. " +
                    "See https://github.com/PGS62/CallFromExcel.jl#installation");
            }

            // Best-effort only: the window is found now that Julia runs in conhost (see above),
            // but wasn't when it ran in Windows Terminal, and might not be in some other
            // setup. That's cosmetic; it doesn't affect whether the launch itself succeeded.
            if (WindowUtils.TryGetHandleFromPartialCaption(windowPartialTitle, out IntPtr hwndFinal))
            {
                return "Julia launched in window \"" + WindowUtils.WindowTitleFromHandle(hwndFinal) + "\"";
            }
            return "Julia launched (listening on port " + JuliaPort.GetJuliaPort() + ")";
        }

        // Julia carries on starting after JuliaLaunch stops waiting for it. The message doesn't
        // suggest calling JuliaLaunch again: that would start a second Julia if the first hasn't
        // finished, and isn't needed - once Julia writes its port file, JuliaEval and JuliaCall
        // find it there.
        private static string StillStartingMessage(int timeoutSeconds)
        {
            return "Julia is still starting after " + timeoutSeconds + " seconds - it may be precompiling " +
                   "packages. Wait until its window shows \"CallFromExcel HTTP server listening\"; JuliaEval " +
                   "and JuliaCall will then connect to it.";
        }

        // "C:\foo\bar" -> "/mnt/c/foo/bar". Port of WSLAddress (modUtils.bas).
        // Quotes s as a single bash word: in single quotes, with each ' written as '\'' - so a
        // path like /mnt/c/Users/O'Brien/... (or one containing $) reaches Julia unchanged.
        public static string BashQuote(string s)
        {
            return "'" + s.Replace("'", "'\\''") + "'";
        }

        // The bash command that starts Julia under WSL. Public, as is WslArguments, so
        // test/runtests.jl can send a path with awkward characters through the real wsl.exe.
        public static string WslJuliaCommand(string juliaExe, string options, string loadFileX)
        {
            return juliaExe + " " + options + " --load " + BashQuote(loadFileX);
        }

        // wsl.exe's arguments to run innerCommand. A login shell (bash -lc) so juliaup's PATH
        // entry is picked up. --exec runs bash directly: without it, wsl.exe first passes the
        // command through the default shell, which expands any $ in the paths. innerCommand is
        // in Windows-level double quotes, so it mustn't contain any itself - paths can't.
        public static string WslArguments(string innerCommand)
        {
            return "--exec bash -lc \"" + innerCommand + "\"";
        }

        private static string WSLAddress(string windowsAddress)
        {
            if (windowsAddress.Length < 3 || windowsAddress[1] != ':' ||
                (windowsAddress[2] != '\\' && windowsAddress[2] != '/'))
            {
                throw new ArgumentException("windowsAddress must start with characters \"x:\\\" for some drive-letter x");
            }
            string driveLetter = char.ToLowerInvariant(windowsAddress[0]).ToString();
            string rest = windowsAddress.Substring(2).Replace("\\", "/");
            return "/mnt/" + driveLetter + rest;
        }

        // The file paths go through MakeJuliaLiteral: a '$' in the temp folder's path (e.g. from
        // the user name) would otherwise be Julia string interpolation. Public so
        // test/runtests.jl can check the generated script parses.
        public static string BuildStartupScript(string packages, int pid, string commsFolderX, string errorFileX, string flagFileX, string literalCommand)
        {
            // PGS 8 Dec 2021 (ported comment): it's important to make "using CallFromExcel"
            // the last "using" statement, believed to help avoid "world-age" problems.
            string allPackages = packages == ""
                ? "Dates," + PackageName
                : "Dates," + packages + "," + PackageName;

            var usingStatements = new StringBuilder();
            foreach (string pkg in allPackages.Split(','))
            {
                string trimmed = pkg.Trim();
                usingStatements.Append("    println(\"using " + trimmed + "\")\n");
                usingStatements.Append("    using " + trimmed + "\n");
            }

            return "try\n" +
                   usingStatements +
                   "    CallFromExcel.setxlpid(" + pid + ")\n" +
                   "    CallFromExcel.comms_folder(" + JuliaLiteral.MakeJuliaLiteral(commsFolderX) + ")\n" +
                   "    println(\"Julia $VERSION, using " + PackageName + " to serve Excel running as process ID " + pid + ".\")\n" +
                   "    println(\"Julia launched with command: " + literalCommand + " \")\n" +
                   "    CallFromExcel.set_console_icon()\n" +
                   "    CallFromExcel.start_server()\n" +
                   "    rm(" + JuliaLiteral.MakeJuliaLiteral(flagFileX) + ")\n" +
                   "catch e\n" +
                   // Julia's own message suggests Pkg.add(\"CallFromExcel\"), which fails: the package isn't registered.
                   "    theerror = if e isa ArgumentError && startswith(e.msg, \"Package " + PackageName + " not found\")\n" +
                   "        \"The " + PackageName + " package is not installed in Julia's active environment ($(Base.active_project())). " +
                   "Install it with: import Pkg; Pkg.add(url=\\\"https://github.com/PGS62/CallFromExcel.jl\\\") " +
                   "- or, if it is installed in a different environment, add --project=<that environment's folder> to JuliaLaunch's CommandLineOptions.\"\n" +
                   "    else\n" +
                   "        \"$e\"\n" +
                   "    end\n" +
                   "    @error theerror \n" +
                   "    errorfile = " + JuliaLiteral.MakeJuliaLiteral(errorFileX) + "\n" +
                   "    io = open(errorfile, \"w\")\n" +
                   "    write(io,theerror)\n" +
                   "    close(io)\n" +
                   "    rm(" + JuliaLiteral.MakeJuliaLiteral(flagFileX) + ")\n" +
                   "end";
        }
    }
}
