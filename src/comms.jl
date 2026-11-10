"""
    setxlpid(pid::Int64)
Set the process id of the instance of Excel that the current Julia process is serving.
"""
function setxlpid(pid::Int64)
    xlpid[] = pid
    settitle()
    println("xlpid set to $pid")
    nothing
end

"""
    getxlpid()
Returns the process id of the instance of Excel that the current Julia process is serving.
"""
function getxlpid()
    xlpid[] == 0 && throw("setxlpid has not been called in this Julia session, it must be \
        called to set the process id of the active Excel session")
    xlpid[]
end

"""
    serve_xl(show_results::Bool=true, pid::Integer=0)

Attaches this Julia session to an Excel process, so it responds to `JuliaCall`/`JuliaEval`
requests from that Excel session exactly as if it had been launched by the `JuliaLaunch()`
worksheet function.

`show_results` is passed to `display_results` - it defaults to `true` here, since attaching
an interactive session like this is exactly when seeing each call and its result echoed to
the REPL is most useful.

`pid` identifies the Excel process to attach to. Leave it at its default of `0` to attach
to the single running Excel process automatically - if none is found, or more than one is,
throws an error explaining how to call `serve_xl` with `pid` given explicitly. Get `pid`
from Excel's `JuliaExcelPID()` worksheet function.

Given explicitly, `pid` is checked against the running Excel processes (Windows only - this check
is skipped, not enforced, when Julia itself is running under WSL) and an error thrown if it doesn't
match one, rather than silently starting a server nothing will ever connect to.

Also throws if Excel already has another Julia session listening for it (e.g. one launched via
`JuliaLaunch`, or attached earlier with `serve_xl` itself) - starting a second one wouldn't do
anything useful, since Excel would keep talking to whichever session it's already connected to.
Stop that session first (e.g. call `stop_server()` there, or close it) before attaching this one.
"""
function serve_xl(show_results::Bool=true, pid::Integer=0)
    if pid == 0
        pids = _running_excel_pids()
        if isempty(pids)
            throw("No running Excel process was found. Open Excel and try again, or call \
                serve_xl(show_results, pid) directly, passing the process id from Excel's \
                JuliaExcelPID() worksheet function.")
        elseif length(pids) > 1
            throw("Found $(length(pids)) running Excel processes (process ids: $(join(pids, ", "))) - \
                call serve_xl(show_results, pid) directly instead, passing the process id of the \
                Excel session to attach to, from its JuliaExcelPID() worksheet function.")
        end
        pid = only(pids)
    elseif Sys.iswindows() && !(pid in _running_excel_pids())
        throw("No running Excel process has process id $pid. Check the value returned by Excel's \
            JuliaExcelPID() worksheet function.")
    end
    _attach_to_excel(show_results, pid)
end

"""
    _attach_to_excel(show_results::Bool, pid::Integer)
Does the actual work of attaching to Excel process `pid`, once `serve_xl` has decided (or
validated) it. Split out from `serve_xl` so it can be tested directly, without needing a real
Excel process for `serve_xl`'s own validation to accept.

Throws if a Julia session is already listening on the port last recorded for `pid` - starting a
second server in that case wouldn't do anything useful: Excel would keep talking to whichever
session it's already connected to, and this one would just sit there receiving nothing. Stop the
other session first (e.g. call `stop_server()` there, or close it) before attaching this one.
"""
function _attach_to_excel(show_results::Bool, pid::Integer)
    setxlpid(Int64(pid))
    comms_folder(_default_commsfolder())
    if isfile(portfile())
        lines = readlines(portfile())   # port, then session token
        existing_port = isempty(lines) ? nothing : tryparse(Int, strip(lines[1]))
        if existing_port !== nothing && existing_port > 0 && _port_is_listening(existing_port)
            throw("Excel (process id $pid) already has a Julia session listening on port \
                $existing_port - stop that session first (e.g. call stop_server() there, or \
                close it), then call serve_xl again.")
        end
    end
    display_results(show_results)
    start_server()
end

"""
    _port_is_listening(port::Integer)::Bool
Returns whether something is currently listening on the given local port. Used by
`_attach_to_excel` to detect a live Julia session already serving the target Excel process before
starting a new one.
"""
function _port_is_listening(port::Integer)::Bool
    try
        close(Sockets.connect("127.0.0.1", port))
        true
    catch
        false
    end
end

"""
    _running_excel_pids()::Vector{Int}
Returns the process ids of all running Excel.exe processes, found via the Windows `tasklist`
utility. Used by `serve_xl` to find the single Excel process to attach to automatically.
"""
function _running_excel_pids()::Vector{Int}
    Sys.iswindows() || throw("Automatic Excel process detection needs Windows - call \
        serve_xl(show_results, pid) directly instead, passing the process id \
        from Excel's JuliaExcelPID() worksheet function.")
    pids = Int[]
    for line in readlines(`tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV /NH`)
        fields = split(line, "\",\"")
        length(fields) >= 2 && push!(pids, parse(Int, fields[2]))
    end
    pids
end

"""
    display_results(switch::Bool)
Switch on or off display in the REPL of both the incoming expression/function call from
Excel and the value returned to Excel, for calls via `JuliaCall` and `JuliaEval`.

"""
function display_results(switch::Bool)
    _display_results[] = switch
    "Results from JuliaCall/JuliaEval $(switch ? "will" : "will not") display in REPL"
end

"""
    display_results()
Returns whether display in the REPL of results of calls from Excel is currently switched on.
"""
display_results() = _display_results[]

"""
    comms_folder()
Returns the name of the comms folder used by CallFromExcel. See also `comms_folder(folder)`.
"""
function comms_folder()
    commsfolder[] == "" && throw("commsfolder has not been set")
    commsfolder[]
end

"""
    comms_folder(folder::String)
Sets the name of the comms folder used by CallFromExcel to `folder`. See also `comms_folder()`.
"""
comms_folder(folder::String) = (commsfolder[] = folder)

"""
    _default_commsfolder()::String
The comms folder the add-in uses (FileUtils.LocalTemp in the C# code): "@CallFromExcel" in the
folder .NET's `Path.GetTempPath()` returns. Used by `serve_xl`, which, unlike `JuliaLaunch`, isn't
told the folder by Excel.

On Windows, Julia's `tempdir()` calls the same Win32 function as `Path.GetTempPath()`
(GetTempPathW, which reads TMP before TEMP), so the two always agree. Under WSL, asks Windows
PowerShell for `Path.GetTempPath()` and converts the result with `wslpath`.
"""
function _default_commsfolder()
    if Sys.iswindows()
        joinpath(tempdir(), "@CallFromExcel")
    elseif Sys.islinux()
        wintemp = try
            readchomp(`powershell.exe -NoProfile -NonInteractive -Command "[IO.Path]::GetTempPath()"`)
        catch
            throw("Cannot find Windows' temp folder: running powershell.exe failed. serve_xl \
                needs Julia running on Windows, or under WSL with Windows executables on the PATH.")
        end
        joinpath(readchomp(`wslpath -u $(strip(wintemp))`), "@CallFromExcel")
    else
        throw("operating system not supported")
    end
end

function installme()
    Sys.iswindows() || throw("CallFromExcel.installme (which installs a Microsoft Excel \
        addin) can only be run from Julia on Windows")
    installscript = normpath(joinpath(@__DIR__, "..", "dotnet-addin", "installer", "Install.ps1"))
    exefile = "C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
    isfile(exefile) || throw("Cannot find PowerShell at '$exefile'")
    isfile(installscript) || throw("Cannot find install script at '$installscript'")
    # Blocking (unlike the old xlam-era installer this replaced, which ran non-blocking so an
    # interactive REPL session stayed responsive while the user clicked through its dialogs) -
    # the new installer is non-interactive and fast, so there's nothing to stay responsive for,
    # and blocking lets any failure (e.g. the add-in not built yet, or Excel still running)
    # surface immediately here rather than silently in a detached background process.
    run(`$exefile -ExecutionPolicy Bypass -NoProfile -File $installscript`)
    nothing
end

portfile() = joinpath(comms_folder(), "Port_$(getxlpid()).txt")

"""
    _encode_result_for_xl(result)::String
Encode `result` for return to Excel, shared by `srv_eval_inner` and `srv_call_inner`. If
`result` itself can't be encoded (e.g. it's a type with no `encode_for_xl` method), reports
that to the Julia console and returns an encoded error string describing the problem
instead.

Callers should invoke this via `Base.invokelatest`: if `result` is a value just defined by
an `eval` earlier in the same request (e.g. `JuliaEval("f(x)=x^2")` returns the function
`f` itself), showing its type here can require looking up a global binding that didn't
exist in the world the caller's own method was compiled in, which Julia 1.12+ reports as a
"world age" warning.
"""
function _encode_result_for_xl(result)::String
    try
        encode_for_xl(result)
    catch e
        println("")
        @error "Result of type $(typeof(result)) could not be encoded for return to Excel."
        encode_for_xl("#Expression evaluated to a variable of type $(typeof(result)), \
            which cannot be returned to Excel because: $(e)!")
    end
end

"""
    srv_eval_inner(expression::String)::String
Evaluate a Julia expression and return the encoded result as a string.
Called by the HTTP request handler in `start_server` for requests to `/eval`, originating
from the add-in's JuliaEval and JuliaEvalVBA.

`expression` may contain several newline-separated top-level statements - the add-in's
`ConcatenateExpressions` (JuliaFunctions.cs) reassembles a one-column range of cells by joining them with `\\n`,
one cell per line, exactly as the user typed them. `Meta.parseall` (rather than `Meta.parse`,
which only ever returns a single top-level form) is what `include_string` itself uses to turn
such a multi-statement string into one `:toplevel` expression, so it already handles `#`
comments, `#= =#` block comments and multi-line `function...end`/`begin...end` bodies spread
across cells correctly - `Main.eval` of that `:toplevel` expression evaluates each statement in
turn and yields the value of the last one.
"""
function srv_eval_inner(expression::String)::String
    global last_question = expression
    if _display_results[]
        printstyled("question> ", color=:green)
        println(expression)
    end
    success = true
    global last_answer = try
        Main.eval(Meta.parseall(expression))
    catch e
        success = false
        printstyled("Something went wrong evaluating the expression: ", color=:red)
        println(expression)
        friendly_error(e)
    end
    if _display_results[] && success
        printstyled("answer> ", color=:green)
        try
            display(last_answer)
        catch e
            printstyled("(could not display result of type $(typeof(last_answer)): $e)\n", color=:red)
        end
        println("")
    end
    Base.invokelatest(_encode_result_for_xl, last_answer)
end

"""
    friendly_error(e)::String

Reports an error to both the Julia console and back to Excel, to be called from a `catch e`
block in `srv_eval_inner`/`srv_call_inner`. Prints the full error to `stdout`, including `e`'s
stacktrace via `catch_backtrace()`, for inspection in the Julia REPL - then returns a short summary
string to send back to Excel as the encoded result.

The returned summary is deliberately short: only the first two lines of `showerror(io, e)` (the
exception's own message - `showerror` only appends a stacktrace when explicitly passed one, which
this call doesn't), further capped at 500 characters. Even without a stacktrace, a `MethodError`'s
"Closest candidates are:" section can run long for a heavily overloaded function (e.g. `+`), so
without this cap a single error could still flood the cell. "Julia REPL has more details and
stacktrace!" points the user at where the full detail actually lives.
"""
function friendly_error(e)
    print("\n")
    showerror(stdout, e, catch_backtrace())
    io = IOBuffer()
    showerror(io, e)
    print("\n\n")
    lines = split(String(take!(io)), '\n')
    error_desc = join(filter(!isempty, strip.(first(lines, 2))), ' ')
    if length(error_desc) > 500
        error_desc = truncate(error_desc, 500) * "..."
    end
    return "#$error_desc Julia REPL has more details and stacktrace!"
end

"""
    split_named_args(args::AbstractVector)

Split the arguments of a call from `JuliaCallNamedArgs`/`JuliaCallNamedArgsVBA` into positional
arguments and named (keyword) arguments, returning a tuple `(positional, named)` where
`positional` is a vector and `named` a `NamedTuple`.

The arguments take the form `p1, ..., pk, ";", name1, value1, ..., namem, valuem`, mirroring
the `;` that separates positional and keyword arguments in Julia's own call syntax. The
separator is found by searching backwards from the third-from-last argument in steps of 2 -
the positions it can occupy given that pairs follow it - and taking the first `";"` found.
Searching backwards means a `";"` among the positional arguments (e.g. the delimiter in a call
to `split`) is never mistaken for the separator. The cost is that a named argument whose
*value* is `";"` must be given as the last pair, since a `";"` value in any earlier pair would
be found first.

Throws an `ArgumentError` if there is no separator, if it is not followed by a positive even
number of arguments (the most likely mistake, which would otherwise surface as a misleading
`MethodError` for the function being called), or if a name is not a valid Julia identifier or is
given more than once (Julia itself would silently use the last of repeated names when splatting).
"""
function split_named_args(args::AbstractVector)
    issep(x) = x isa AbstractString && x == ";"
    sep = 0
    for i in length(args)-2:-2:1
        if issep(args[i])
            sep = i
            break
        end
    end
    if sep == 0
        last_sep = findlast(issep, args)
        last_sep === nothing && throw(ArgumentError("No \";\" argument found to mark where the \
            named arguments start. Use JuliaCall or JuliaCallVBA to call a function without \
            named arguments"))
        n = length(args) - last_sep
        throw(ArgumentError(n == 0 ?
            "No named arguments follow the \";\" - give them after it as name-value pairs" :
            "Named arguments must be given as name-value pairs after the \";\", but an odd number \
            ($n) of arguments follow it"))
    end
    names = args[sep+1:2:end]
    values = args[sep+2:2:end]
    for name in names
        (name isa AbstractString && Base.isidentifier(name)) || throw(ArgumentError(
            "Named argument name $(repr(name)) is not a valid Julia identifier"))
    end
    for name in unique(names)
        count(==(name), names) > 1 && throw(ArgumentError(
            "Named argument `$name` is given more than once"))
    end
    args[1:sep-1], NamedTuple{Tuple(Symbol.(names))}(Tuple(values))
end

"""
    srv_call_inner(payload::String; named_args::Bool=false)::String

Decode `payload` (in the CallFromExcel wire format - a 1D array whose first element is a
function name and remaining elements are its arguments), call the named function, and
return the encoded result as a string. Called by the HTTP request handler in `start_server`
for requests to `/call`, originating from the add-in's JuliaCall and JuliaCallVBA,
and with `named_args=true` for requests to `/callkw`, originating from JuliaCallNamedArgs and
JuliaCallNamedArgsVBA. In that case the arguments are split into positional and named
arguments by `split_named_args`, the named ones being saved in `kwargs_from_xl`.

Avoids the `Meta.parseall` of a literal expression that `srv_eval_inner` requires, which is
slow for large arrays of arguments.

A trailing "." on the function name (e.g. "f.") requests broadcasting, matching Julia's
`f.(...)` call syntax. That syntax is a transform on the call site itself (it lowers to
`broadcast(f, ...)`) rather than a property of a standalone function reference, so "f."
can't just be handed to `Meta.parse` as-is - the dot is stripped before resolving the
function, and `broadcast` is called explicitly instead of a plain call. As with `f.(x; k=v)`
in Julia, named arguments are passed unchanged to every call, not broadcast over.
"""
function srv_call_inner(payload::String; named_args::Bool=false)::String
    fn_name = "<unknown>"
    global args_from_xl = ["<unknown>"]
    global kwargs_from_xl = NamedTuple()
    broadcasting = false
    success = true
    kwpart = named_args ? "; kwargs_from_xl..." : ""
    global last_answer = try
        decoded = decode_from_xl(payload)
        fn_name = decoded[1]::String
        broadcasting = endswith(fn_name, ".")
        broadcasting && (fn_name = chop(fn_name))
        global last_question = broadcasting ? "$fn_name.(args_from_xl...$kwpart)" : "$fn_name(args_from_xl...$kwpart)"
        if _display_results[]
            printstyled("question> ", color=:green)
            println(last_question)
        end
        fn_to_call = Main.eval(Meta.parse(fn_name))             # fast: parses only the short function name
        args_from_xl = decoded[2:end]
        named_args && ((args_from_xl, kwargs_from_xl) = split_named_args(args_from_xl))
        if broadcasting
            broadcast((xs...) -> fn_to_call(xs...; kwargs_from_xl...), args_from_xl...)
        else
            fn_to_call(args_from_xl...; kwargs_from_xl...)
        end
    catch e
        success = false
        global last_question = broadcasting ? "$fn_name.(args_from_xl...$kwpart)" : "$fn_name(args_from_xl...$kwpart)"
        printstyled("Something went wrong calling the Julia function $fn_name", color=:red)
        print(" from Excel against\narguments saved in args_from_xl", named_args ? " and kwargs_from_xl" : "")
        print(" (overwritten by the next call),")
        print(" so\nthe error should be reproducible from here with '$last_question'.\n\n")
        friendly_error(e)
    end
    if _display_results[] && success
        printstyled("answer> ", color=:green)
        try
            display(last_answer)
        catch e
            printstyled("(could not display result of type $(typeof(last_answer)): $e)\n", color=:red)
        end
        println("")
    end
    Base.invokelatest(_encode_result_for_xl, last_answer)
end

"""
    answer_again()

Evaluates `last_question` again, and you can wrap it with `@enter` or your own debugging tools.
Useful for interactively debugging a `JuliaCall`/`JuliaEval` that failed (or just misbehaved) when
called from Excel, without needing to repeat the call from there.

Uses `Meta.parseall` rather than `Meta.parse` since `last_question` set by `srv_eval_inner` can be
a multi-statement, newline-separated string (see its docstring) - `Meta.parseall` also handles the
single-statement string that `srv_call_inner` sets `last_question` to, so this is correct either way.

See also `last_question`, `last_answer`, `args_from_xl`.
"""
answer_again() = Main.eval(Meta.parseall(last_question))

"""
    start_server(start::Int=2700)
Start an HTTP server on a free local port that handles evaluation requests from Excel,
trying up to 100 candidate ports starting from `start`. Requests to `/eval` go to
`srv_eval_inner`, and to `/call` and `/callkw` to `srv_call_inner`. Any other path gets a 404
response. Tries the real bind directly rather
than probing with a separate test-listener first: a probe-then-release check has a gap
between "verified free" and the real bind moments later, during which another process (e.g.
an orphaned Julia session from a previously-closed Excel session) could still be holding
the port - or, since HTTP.jl's server binds with `reuseaddr=true`, could appear to succeed
even though something else is already listening there. Retrying on the real bind failure
avoids relying on that separate check being reliable.

Writes the chosen port to the port file so the add-in can discover it during JuliaLaunch.

Closes any server already started by a previous call in this session first - otherwise, calling
this (or `serve_xl`) more than once would leave each earlier listener still running, each holding
its own port open indefinitely.
"""
function start_server(start::Int=2700)
    _server[] !== nothing && close(_server[])
    # A secret for this session, written to the port file (in the user's own temp folder) for
    # Excel to read and send with every request - see _reject_request.
    session_token[] = bytes2hex(rand(RandomDevice(), UInt8, 16))
    port = start
    while true
        try
            bound_port = port
            _server[] = HTTP.serve!("127.0.0.1", port) do req
                rejection = _reject_request(req, bound_port)
                rejection === nothing || return rejection
                # /ping has no side effects: Excel uses it to check that the Julia on a port is
                # alive and serving it, without disturbing last_question/last_answer.
                req.target == "/ping" && return _text_response("CallFromExcel")

                body = String(req.body)
                # No default endpoint: the add-in always names one, so anything else is a bug worth
                # reporting rather than guessing.
                result = req.target == "/eval" ? srv_eval_inner(body) :
                         req.target == "/call" ? srv_call_inner(body) :
                         req.target == "/callkw" ? srv_call_inner(body; named_args=true) :
                         return HTTP.Response(404, "Unknown CallFromExcel endpoint $(req.target)")
                _text_response(result)
            end
            break
        catch
            port += 1
            port > start + 100 && rethrow()
        end
    end
    mkpath(comms_folder())
    open(portfile(), "w") do f
        write(f, string(port, "\n", session_token[]))
    end
    xlport[] = port
    settitle()
    println("CallFromExcel HTTP server listening on port $port")
    nothing
end

# The explicit Content-Length matters for performance, not just correctness: without it, HTTP.jl
# falls back to chunked transfer-encoding, which clients reassemble less efficiently.
_text_response(s::String) = HTTP.Response(200, ["Content-Type" => "text/plain; charset=utf-8",
    "Content-Length" => string(ncodeunits(s))], s)

"""
    _reject_request(req, port) -> Union{HTTP.Response,Nothing}
Returns an error response for a request that must not be served, or `nothing` to serve it.

- Only POST is accepted.
- A request carrying an `Origin` header comes from a web page (browsers add one to every
  cross-origin POST; Excel's add-in never sends one). Without this check, any web page open in a
  browser could POST code to /eval, since a text/plain POST needs no CORS preflight.
- The `Host` header must name this machine's loopback address, which defeats DNS rebinding.
- The `X-CallFromExcel-Token` header must equal this session's token, which only the user's own
  Excel can read (from the port file in the user's temp folder). This stops other local
  processes, including other users' on a shared machine, and it means an Excel whose own Julia
  has gone away never silently talks to another Excel's Julia that has since taken over the same
  port: that Excel's port file holds a different token. Excel treats this 409 as "Julia not
  reachable".
"""
function _reject_request(req, port::Integer)
    req.method == "POST" || return HTTP.Response(405, "CallFromExcel accepts only POST requests")
    HTTP.hasheader(req, "Origin") &&
        return HTTP.Response(403, "CallFromExcel does not accept requests from web pages")
    HTTP.header(req, "Host") in ("127.0.0.1", "127.0.0.1:$port", "localhost", "localhost:$port") ||
        return HTTP.Response(403, "CallFromExcel accepts only requests addressed to 127.0.0.1")
    _tokens_equal(HTTP.header(req, "X-CallFromExcel-Token"), session_token[]) ||
        return HTTP.Response(409, "Request refused: not from the Excel session this Julia serves")
    nothing
end

# Constant-time comparison, so response timing reveals nothing about the token.
function _tokens_equal(a::AbstractString, b::AbstractString)
    x, y = codeunits(a), codeunits(b)
    (isempty(y) || length(x) != length(y)) && return false
    diff = 0x00
    for i in eachindex(x, y)
        diff |= xor(x[i], y[i])
    end
    diff == 0x00
end

"""
    stop_server()
Stops this Julia session's HTTP server, if one is running, so it no longer responds to
`JuliaCall`/`JuliaEval` requests from Excel. No-op if no server is currently running.

Also clears the port from this session's window title (via `settitle`) - otherwise the title would
keep claiming a port this session no longer serves, which could wrongly look like a match for
whichever port a *different* session (e.g. one started with `serve_xl`) happens to pick up next.

Leaves the port file on disk untouched, so Excel will still try the now-dead port on its next
request - and get a clean "no connection" error rather than reaching this session again.
"""
function stop_server()
    if _server[] !== nothing
        close(_server[])
        _server[] = nothing
        xlport[] = 0
        settitle()
    end
    nothing
end

"""
    server_status()
Returns a `NamedTuple` `(running, pid, port, display_results)` describing this Julia session's HTTP
server: `running` is `true` if a server is currently listening; `pid` is the Excel process id it's
set up to serve (0 if `setxlpid` has never been called), regardless of whether the server is
currently running; `port` is the port it's listening on, or 0 if not running; `display_results` is
the current setting of `display_results()`.
"""
function server_status()
    running = _server[] !== nothing && isopen(_server[])
    (running=running, pid=xlpid[], port=running ? xlport[] : 0, display_results=display_results())
end

"""
    setvar(name::String, arg)
Set a variable in global scope. Called by the add-in's JuliaSetVar (via the /call endpoint).
"""
function setvar(name::String, arg)

    if Base.isidentifier(name)
        # Assign the value itself, not its printed form re-parsed as source code, which
        # fails for strings ("hello" would become the variable name hello) and for values
        # with no parseable repr, such as ExcelError.
        Core.eval(Main, Expr(:global, Expr(:(=), Symbol(name), QuoteNode(arg))))

        thesize = ()
        thetype = typeof(arg)
        try
            thesize = size(arg)
        catch
        end

        numdims = length(thesize)
        if numdims == 0
            sizedesc = ""
        elseif numdims == 1
            sizedesc = "$(thesize[1])-element "
        elseif numdims > 1
            sizedesc = join(thesize, "x") * " "
        end
        "Set global variable `$name` to $sizedesc$thetype"

    else
        "#`$name` is not an allowed variable name in Julia!"
    end
end

# https://docs.microsoft.com/en-us/windows/terminal/tutorials/tab-title
function settitle()
    if Sys.islinux()
        os = "Linux"
    elseif Sys.iswindows()
        os = "Windows"
    end

    portpart = xlport[] == 0 ? "" : ", port $(xlport[])"
    print("\033]0;Julia $VERSION on $os serving Excel PID $(getxlpid())$portpart\a")
end

"""
    truncate(x::String, maxlength::Int)
Abbreviate a string to show only `maxlength` characters.
"""
function truncate(x::String, maxlength::Int)
    if (length(x)) > maxlength
        first(x, maxlength ÷ 2) * " … " * last(x, maxlength - (maxlength ÷ 2) - 1)
    else
        x
    end
end
