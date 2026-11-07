# Do not run this file directly in an interactive REPL (e.g. VS Code's "Julia: Execute File in
# REPL") - it calls exit() on failure, which is correct for the short-lived, disposable Julia
# process Pkg.build runs this in, but kills your actual REPL session stone dead if run there
# directly. To install/re-install manually, use `using CallFromExcel; CallFromExcel.installme()`
# instead (or just `Pkg.build("CallFromExcel")`, which runs this file the way it's meant to be run).
#
# Launches dotnet-addin/installer/Install.ps1 to register the C#/Excel-DNA CallFromExcel add-in,
# so that `Pkg.add("CallFromExcel")` alone is enough to install it - no separate manual call to
# CallFromExcel.installme() needed. Pkg runs this script automatically right after add/build.
#
# Doesn't `using CallFromExcel` here - the package may not be safely loadable during its own
# build, so the handful of lines from installme() (src/comms.jl) are duplicated rather than
# shared; installme() stays as-is for manual re-install/repair, where throwing on failure is
# the right behaviour, unlike here.
#
# Skipped entirely (does nothing) when:
#   - not running on Windows - the installer registers a Windows Excel add-in via COM/PowerShell
#   - running under CI - there's no Excel to install into there
if Sys.iswindows() && get(ENV, "CI", "false") != "true"
    installscript = normpath(joinpath(@__DIR__, "..", "dotnet-addin", "installer", "Install.ps1"))
    exefile = "C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
    if !isfile(exefile) || !isfile(installscript)
        @error "CallFromExcel: could not find powershell.exe or the install script; run `using CallFromExcel; CallFromExcel.installme()` manually instead."
        exit(1)
    end

    println("CallFromExcel: running the Excel add-in installer...")
    try
        # Blocking (not wait=false): Pkg.build runs this script in its own short-lived Julia
        # process, and on Windows a spawned child is killed when that process's job object
        # closes (i.e. as soon as this process exits) - a non-blocking run here would launch
        # powershell.exe only to have it killed moments later, before it can finish.
        # -ExecutionPolicy Bypass: this is a one-off script launch, not something that should
        # depend on (or need to change) whatever execution policy is set on the user's machine.
        run(`$exefile -ExecutionPolicy Bypass -NoProfile -File $installscript`)
    catch e
        # exit(1) rather than error(...): a non-zero exit here is enough for Pkg to mark the
        # build as failed. Note this doesn't avoid a stacktrace being shown - Pkg detects the
        # failed exit code and throws its own PkgError from deep inside its own internals
        # (Operations.build_versions), which is what actually produces the stacktrace visible
        # to the user; that happens regardless of whether we exit(1) or throw here ourselves.
        # It's an accepted tradeoff: a scary-looking but immediate, hard-to-miss failure beats
        # a clean warning that only reaches a log file nobody will look at.
        @error "CallFromExcel: the add-in installer did not complete ($e). Close Excel if it's running, then run `import Pkg; Pkg.build(\"CallFromExcel\")` to try again."
        exit(1)
    end
end
