using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
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
                // One interactive thread, as HTTP.jl recommends for servers (see the "Server"
                // docstring in HTTP.jl's http_server.jl), alongside Julia's usual threads.
                // Added in August 2026, when exit() sent from Excel could fail to shut Julia
                // down with a single thread; checked in November 2026, that no longer
                // happens either way, so this is now just following HTTP.jl's advice.
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

        // Part of the title of the Julia window serving the Excel process pid - must match
        // Julia's CallFromExcel.settitle.
        internal static string JuliaWindowTitle(int pid)
        {
            return "serving Excel PID " + pid;
        }

        // Launches Julia (or finds it already running), then checks it's the same version of
        // CallFromExcel as this add-in, so a mismatch shows straight away in the launching cell
        // or dialog rather than at the first JuliaCall. No check while Julia is still starting
        // (no port yet): the first call after that checks instead.
        public static string Launch(
            bool useLinux = false,
            bool minimiseWindow = false,
            string commandLineOptions = "",
            string packages = "",
            string bashStatements = "",
            int timeoutSeconds = 30)
        {
            string message = LaunchCore(useLinux, minimiseWindow, commandLineOptions, packages, bashStatements, timeoutSeconds);
            int port = JuliaPort.GetJuliaPortOrZero();
            if (port > 0) JuliaPort.CheckVersion(port);
            return message;
        }

        private static string LaunchCore(
            bool useLinux,
            bool minimiseWindow,
            string commandLineOptions,
            string packages,
            string bashStatements,
            int timeoutSeconds)
        {
            string juliaExe = JuliaExeLocation.Find(useLinux);

            string userSuppliedCommandLineOptions = commandLineOptions ?? "";
            commandLineOptions = NormaliseCommandLineOptions(userSuppliedCommandLineOptions);

            int pid = Process.GetCurrentProcess().Id;
            string windowPartialTitle = JuliaWindowTitle(pid);

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
            // taskbar shows Julia's icon (for a Microsoft Store Julia), the add-in sets its title-bar
            // icon (JuliaIcon.TrySetOnConsole, below), and MinimiseWindow is
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

            string loadFileContents = BuildStartupScript(packages, pid, commsFolderX, errorFileX, flagFileX, literalCommand, FirstFreePort());
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
            // The window's Julia icon is set as soon as the window appears (within the first few
            // seconds), not left to Julia, which can only do it once its packages have loaded.
            bool iconSet = false;
            while (File.Exists(flagFile))
            {
                Thread.Sleep(50);
                if (!iconSet && stopwatch.Elapsed.TotalSeconds < 10)
                {
                    iconSet = JuliaIcon.TrySetOnConsole(juliaProcess.Id);
                }
                if (juliaProcess.HasExited)
                {
                    throw ProcessExitedError(userSuppliedCommandLineOptions.Trim(), useLinux, bashStatements);
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
                string[] error = FileUtils.ReadTextFile(errorFile).Split('\n');
                if (error[0].Trim() == VersionMismatchMarker)
                {
                    throw JuliaPort.MismatchError(
                        error.Length > 1 ? error[1].Trim() : "", error.Length > 2 ? error[2].Trim() : null, useLinux);
                }
                string text = string.Join("\n", error).Trim();
                System.Text.RegularExpressions.Match notInstalled = NotInstalled.Match(text);
                if (notInstalled.Success)
                {
                    throw new JuliaCommandException(text, "import Pkg; Pkg.add(\"" + notInstalled.Groups[1].Value + "\")");
                }
                if (text.StartsWith(ProjectMissingStart, StringComparison.Ordinal) ||
                    text.StartsWith(TemporaryNeedsJuliaStart, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(text);
                }
                throw new InvalidOperationException(
                    "Julia launched but encountered an error when executing '" + loadFile + "' the error was: " + text);
            }

            // Under WSL the port can take a while to become reachable from Windows after Julia
            // starts listening (WSL2 port forwarding), so wait for it within the same timeout.
            // Each ping may take up to the time remaining (at least 2 seconds): a first request
            // compiles Julia's request handling, and giving up on it sooner only made Julia print
            // an error when its late answer found the connection closed. (Julia now answers a
            // ping itself before saying it's ready - CallFromExcel.warm_up - so it's quick anyway.)
            while (port > 0 && !JuliaHttpClient.IsListening(port,
                       (int)Math.Max(2000, timeoutSeconds * 1000 - stopwatch.ElapsedMilliseconds)))
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
        public const string VersionMismatchMarker = "CallFromExcelVersionMismatch";

        // The start of the start-up script's message when a package - CallFromExcel, or one of
        // the user's Packages - isn't installed, which LaunchCore recognises, so the ribbon can
        // offer to copy the command that installs it.
        private static readonly System.Text.RegularExpressions.Regex NotInstalled =
            new System.Text.RegularExpressions.Regex(@"^The (\S+) package is not installed in ");

        // The start of the start-up script's message when the --project folder doesn't exist.
        public const string ProjectMissingStart = "The project folder";

        // The start of its message when --project=@temp is used with Julia 1.10.
        public const string TemporaryNeedsJuliaStart = "--project=@temp, a temporary environment, needs Julia 1.11";

        // Julia's documentation of its command line options.
        public const string CommandLineDocsUrl = "https://docs.julialang.org/en/v1/manual/command-line-interface/#command-line-switches-for-julia";

        // The error when Julia's process exits before start-up finishes. An invalid command line
        // option is the usual cause, so with options of the user's own the message shows them,
        // with a link to Julia's documentation of them - but not as the only possible cause, as
        // a ribbon launch with a Project folder always has one. Without options, it doesn't
        // suggest them. No exit code: the process watched is conhost (see LaunchCore), which
        // exits with 0 whatever Julia's own exit code was.
        private static Exception ProcessExitedError(string options, bool useLinux, string bashStatements)
        {
            const string exited = "Julia's process exited before start-up finished. ";
            string where = useLinux ? " under WSL" : "";
            string bash = useLinux && !string.IsNullOrEmpty(bashStatements) ? ", and that the Bash statements are valid" : "";
            string otherwise = "Julia may have failed to start, or its window may have been closed. " +
                "Check that Julia starts from a command prompt" + where + bash + ".";
            if (options != "")
            {
                return new HelpLinkException(exited + "Julia stops straight away if a command line option isn't valid, " +
                    "so check the options it was given: " + options + ". Julia's command line options are described at " +
                    CommandLineDocsUrl + ". Otherwise, " + otherwise,
                    CommandLineDocsUrl, options);
            }
            return new InvalidOperationException(exited + otherwise);
        }

        // The first port from `start` that nothing on Windows is listening on, where Julia then
        // starts its server (start_server tries upwards from the port it's given). Chosen here,
        // not by Julia, because Julia under WSL can't see Windows' ports: it took 2700 even with
        // a Windows Julia (e.g. one serving another Excel) already listening there, and Windows
        // then sent this Excel's requests to that one, which refused them. Ports forwarded from
        // WSL count as taken here too, so two WSL sessions don't clash either. Public so
        // test/runtests.jl can test it.
        public static int FirstFreePort(int start = 2700)
        {
            for (int port = start; port < start + 100; port++)
            {
                var listener = new TcpListener(IPAddress.Loopback, port) { ExclusiveAddressUse = true };
                try
                {
                    listener.Start();
                    return port;
                }
                catch (SocketException) { }
                finally
                {
                    listener.Stop();
                }
            }
            return start;
        }

        public static string BuildStartupScript(string packages, int pid, string commsFolderX, string errorFileX, string flagFileX, string literalCommand, int port)
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
                // A trailing or doubled comma ("DataFrames,") would otherwise give a bare "using",
                // which Julia can't parse - so the script never ran, and Excel waited out the
                // time-out with no error.
                if (trimmed == "") continue;
                usingStatements.Append("        println(\"using " + trimmed + "\")\n");
                usingStatements.Append("        using " + trimmed + "\n");
            }

            // The user's own Packages, as a Julia vector of names, for installing in a temporary
            // environment (below).
            var userPackages = new List<string>();
            foreach (string pkg in packages.Split(','))
            {
                if (pkg.Trim() != "") userPackages.Add(JuliaLiteral.MakeJuliaLiteral(pkg.Trim()));
            }
            string userPackagesVector = "String[" + string.Join(", ", userPackages) + "]";

            // First, that the project folder exists: Julia accepts --project naming a folder that
            // doesn't, as a new empty project, and still finds CallFromExcel in the default
            // environment - so a mistyped folder would otherwise go unnoticed. Reported through the
            // error file, starting ProjectMissingStart.
            //
            // Then the version check, before anything else in the package, which an older
            // CallFromExcel may not have. A mismatch is reported through the error file as
            // VersionMismatchMarker, Julia's version and its project folder ("" for the default
            // environment, as in comms.jl's project_folder), for LaunchCore to turn into the
            // same message as JuliaPort.CheckVersion's.
            //
            // The add-in sets the window's icon itself (JuliaIcon.TrySetOnConsole), as soon as the
            // window appears, rather than leaving it to CallFromExcel.set_console_icon here, which
            // only runs after the packages have loaded.
            string addInVersion = JuliaLiteral.MakeJuliaLiteral(JuliaPort.AddInVersion);
            string errorFile = JuliaLiteral.MakeJuliaLiteral(errorFileX);
            // --project=@temp, a temporary environment, needs Julia 1.11: Julia 1.10 reads it as
            // the permanent shared environment named "temp" instead, so it's refused there. With a
            // temporary environment, packages come from the default environment, so a mismatch
            // message names that ("" as the folder), not the temporary folder, gone by then.
            //
            // In a temporary environment, any of the user's Packages that can't be found (in it or
            // the default environment) are installed into it - harmless, as it's deleted when
            // Julia exits. PRESERVE_TIERED_INSTALLED makes Pkg try versions already installed
            // first (only falling back to others if they won't do), so a launch with the same
            // Packages as before gets the same versions, already precompiled, rather than
            // whatever is newest. (PRESERVE_ALL_INSTALLED would allow only installed versions,
            // so couldn't install a package for the first time.) Elsewhere, a missing package is an
            // error naming it and its environment, whose fix (import Pkg; Pkg.add(...)) the
            // ribbon's error dialog offers to copy; nothing is installed into a project or the
            // default environment without the user's say-so. CallFromExcel itself is never
            // installed here: it must match the add-in, which installme() keeps in step.
            // projectfile and temporary are set before the try, so the catch can see them: in a
            // file, a variable first assigned inside a try block is local to it.
            return "projectfile = Base.active_project()\n" +
                   // As comms.jl's temporary_project: --project, else the JULIA_PROJECT variable.
                   "temporary = (Base.JLOptions().project != C_NULL ? unsafe_string(Base.JLOptions().project) : get(ENV, \"JULIA_PROJECT\", \"\")) == \"@temp\"\n" +
                   "try\n" +
                   "    if temporary && VERSION < v\"1.11\"\n" +
                   "        theerror = \"" + TemporaryNeedsJuliaStart + " or later, but this is Julia $VERSION. " +
                   "Choose a different Project folder in Launch Settings, or change the --project option in JuliaLaunch's CommandLineOptions.\"\n" +
                   "        @error theerror\n" +
                   "        write(" + errorFile + ", theerror)\n" +
                   "    elseif projectfile !== nothing && !isdir(dirname(projectfile))\n" +
                   "        theerror = \"" + ProjectMissingStart + " $(dirname(projectfile)) doesn't exist. " +
                   "Check the Project folder in Launch Settings, or the --project option in JuliaLaunch's CommandLineOptions.\"\n" +
                   "        @error theerror\n" +
                   "        write(" + errorFile + ", theerror)\n" +
                   "    else\n" +
                   "        if temporary\n" +
                   "            toinstall = filter(p -> Base.find_package(p) === nothing, " + userPackagesVector + ")\n" +
                   "            if !isempty(toinstall)\n" +
                   "                println(\"Installing $(join(toinstall, \", \")) in the temporary environment...\")\n" +
                   "                import Pkg\n" +
                   "                Pkg.add(toinstall; preserve=Pkg.PRESERVE_TIERED_INSTALLED)\n" +
                   "            end\n" +
                   "        end\n" +
                   usingStatements +
                   "        juliaversion = string(pkgversion(CallFromExcel))\n" +
                   "        if juliaversion != " + addInVersion + "\n" +
                   "            @error \"CallFromExcel $juliaversion doesn't match the Excel add-in, version \" * " + addInVersion + " * \". Excel says how to fix this.\"\n" +
                   "            folder = temporary || projectfile === nothing || projectfile == Base.load_path_expand(\"@v#.#\") ? \"\" : dirname(projectfile)\n" +
                   "            write(" + errorFile + ", \"" + VersionMismatchMarker + "\\n$juliaversion\\n$folder\")\n" +
                   "        else\n" +
                   "            CallFromExcel.setxlpid(" + pid + ")\n" +
                   "            CallFromExcel.comms_folder(" + JuliaLiteral.MakeJuliaLiteral(commsFolderX) + ")\n" +
                   "            println(\"Julia $VERSION, using " + PackageName + " to serve Excel running as process ID " + pid + ".\")\n" +
                   "            println(\"Julia launched with command: " + literalCommand + " \")\n" +
                   "            CallFromExcel.start_server(" + port + ")\n" +
                   "        end\n" +
                   "    end\n" +
                   "    rm(" + JuliaLiteral.MakeJuliaLiteral(flagFileX) + ")\n" +
                   "catch e\n" +
                   // A package that isn't installed - CallFromExcel or one of the user's Packages:
                   // clearer than Julia's own message, as it names the environment, and mentions
                   // --project. Under @temp only CallFromExcel can be missing (the others were
                   // installed above), and it's the default environment that needs it. Other
                   // errors as Julia's own showerror text.
                   "    notfound = e isa ArgumentError ? match(r\"^Package (\\S+) not found\", e.msg) : nothing\n" +
                   "    theerror = if notfound !== nothing && temporary\n" +
                   "        pkg = notfound.captures[1]\n" +
                   "        \"The $pkg package is not installed in your default environment, which a temporary environment (--project=@temp) takes its packages from. " +
                   "Install it there, in Julia started without --project, with: import Pkg; Pkg.add(\\\"$pkg\\\")\"\n" +
                   "    elseif notfound !== nothing\n" +
                   "        pkg = notfound.captures[1]\n" +
                   "        \"The $pkg package is not installed in Julia's active environment ($(Base.active_project())). " +
                   "Install it with: import Pkg; Pkg.add(\\\"$pkg\\\") " +
                   "- or, if it is installed in a different environment, choose that as the Project folder in Launch Settings, " +
                   "or add --project=<its folder> to JuliaLaunch's CommandLineOptions.\"\n" +
                   // Some errors, e.g. Pkg's ResolverError, have no showerror method of their own, so
                   // would print as ResolverError("..."): their message is clearer.
                   "    elseif hasproperty(e, :msg) && e.msg isa AbstractString\n" +
                   "        e.msg\n" +
                   "    else\n" +
                   "        sprint(showerror, e)\n" +
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
