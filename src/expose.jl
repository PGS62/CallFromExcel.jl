#=
SPIKE (branch spike/named-functions): named worksheet functions, see "Named worksheet
functions" in Enhancements.md.

expose records a Julia function to be registered in Excel under its own name, by default
"jl." followed by the Julia name. The add-in fetches the list with exposed_functions and
registers each as a worksheet function that forwards to JuliaCall's code.
=#

const _exposed = Dict{String,Any}()   # Excel name => (julia, description, argnames, argdescriptions)

# Separates the arguments' names, and their descriptions, in one string for the add-in.
const _ARGSEP = "\x1f"

# Counts calls to expose. The add-in compares exposed_version() with the version it last
# registered after JuliaLaunch, JuliaInclude and JuliaEval, and registers again if it differs.
const _exposed_count = Ref(0)

"""
    exposed_version()

For the add-in: `"<session>#<count>"`, which changes whenever `expose` is called. The session
tag (as in handles) tells a relaunched Julia apart, whose count could match the old one's.
"""
exposed_version() = _session_tag() * "#" * string(_exposed_count[])

"""
    expose(fn_name; name=nothing, description="", argdescriptions=String[])

SPIKE. Make the Julia function named `fn_name` callable from Excel as a worksheet function,
by default `jl.<fn_name>`. Returns the Excel name. The add-in registers exposed functions
after `JuliaLaunch`, `JuliaInclude` and `JuliaEval`, and when the ribbon's Register
Functions button is clicked (for functions exposed from the REPL).

The Excel function's arguments are those of the method with the most positional arguments
(not counting varargs methods). Arguments with default values (`f(a, b=2)`) are optional in
Excel: left out at the end of a call, they take their defaults.

Descriptions, shown by Excel's Function Wizard and IntelliSense, come from the function's
docstring unless given: the first paragraph after the signature describes the function, and
an `# Arguments` list (``- `name`: description``) its arguments. `argdescriptions` replaces
the docstring's, position by position, where not empty.
"""
function expose(fn_name::AbstractString; name=nothing, description::AbstractString="",
                argdescriptions=String[])
    f = Main.eval(Meta.parse(fn_name))
    m, nrequired = _excel_method(f)
    argnames = [string(a) for a in Base.method_argnames(m)[2:end]]
    excel_name = name === nothing ? "jl." * fn_name : String(name)
    _check_excel_name(excel_name, name === nothing)

    doc = _docstring(f, m)
    desc = isempty(description) ? _doc_description(doc) : String(description)
    docargs = _doc_arguments(doc)
    argdescs = map(enumerate(argnames)) do (k, a)
        d = k <= length(argdescriptions) && !isempty(argdescriptions[k]) ? String(argdescriptions[k]) :
            get(docargs, a, "")
        k > nrequired && !startswith(lowercase(d), "optional") && (d = rstrip("Optional. " * uppercasefirst(d)))
        _fit(d)
    end
    _exposed[excel_name] = (julia=String(fn_name), description=_fit(desc),
        argnames=argnames, argdescriptions=argdescs, nrequired=nrequired)
    _exposed_count[] += 1
    excel_name
end

# The method Excel's function follows, and how many of its arguments are required: the
# non-varargs method with the most arguments (the first such, if several), and the fewest
# arguments among the methods its default values generate - those share its file and line,
# and their argument names begin its own.
function _excel_method(f)
    ms = collect(methods(f))
    isempty(ms) && throw(ArgumentError("`$f` has no methods."))
    fixed = filter(m -> !m.isva, ms)
    candidates = isempty(fixed) ? ms : fixed
    m = candidates[argmax([c.nargs for c in candidates])]
    names = Base.method_argnames(m)[2:end]
    nrequired = length(names)
    for other in fixed
        n = other.nargs - 1
        if other.file == m.file && other.line == m.line && n < nrequired &&
           Base.method_argnames(other)[2:end] == names[1:n]
            nrequired = n
        end
    end
    m, nrequired
end

# The raw text of f's docstring - the one for method m if there are several - or "".
function _docstring(f, m)
    binding = Base.Docs.Binding(parentmodule(f), nameof(f))
    sig = Base.tuple_type_tail(m.sig)
    found = String[]
    for mod in Base.Docs.modules
        meta = Base.Docs.meta(mod)
        haskey(meta, binding) || continue
        for (docsig, ds) in meta[binding].docs
            text = join(string.(ds.text))
            sig <: docsig && return text
            push!(found, text)
        end
    end
    length(found) == 1 ? only(found) : ""
end

# Markdown to one line of plain text: links keep their text, code and bold lose their marks.
function _plain(s::AbstractString)
    s = replace(s, r"\[([^\]]*)\]\([^)]*\)" => s"\1")
    s = replace(s, "`" => "", "**" => "")
    strip(replace(s, r"\s+" => " "))
end

# The docstring's first paragraph after its signature (indented or fenced code).
function _doc_description(doc::AbstractString)
    lines = split(doc, '\n')
    i = 1
    skip_blank() = while i <= length(lines) && isempty(strip(lines[i])); i += 1; end
    skip_blank()
    if i <= length(lines) && startswith(lines[i], "```")
        i += 1
        while i <= length(lines) && !startswith(lines[i], "```"); i += 1; end
        i += 1
    else
        while i <= length(lines) && (startswith(lines[i], "    ") || startswith(lines[i], "\t")); i += 1; end
    end
    skip_blank()
    para = String[]
    while i <= length(lines) && !isempty(strip(lines[i])) && !startswith(lines[i], "#")
        push!(para, lines[i]); i += 1
    end
    _plain(join(para, " "))
end

# The docstring's "# Arguments" list: argument name => description.
function _doc_arguments(doc::AbstractString)
    result = Dict{String,String}()
    lines = split(doc, '\n')
    start = findfirst(l -> occursin(r"^#+\s*Arguments\s*$", l), lines)
    start === nothing && return result
    item = r"^\s*[-*+]\s+`([^`:=]+)[^`]*`\s*[:\-]?\s*(.*)$"
    current = nothing
    for l in lines[start+1:end]
        if (c = match(item, l)) !== nothing
            current = String(strip(c[1]))
            result[current] = c[2]
        elseif current !== nothing && !isempty(strip(l)) && (startswith(l, " ") || startswith(l, "\t"))
            result[current] *= " " * l                 # an item continued on the next line
        elseif startswith(l, "#") || !isempty(strip(l))
            break                                      # the list has ended
        end
    end
    Dict(k => _plain(v) for (k, v) in result)
end

# Excel-DNA's limit for descriptions is 255 characters; 253 as for the add-in's own (see
# the test "Function and argument descriptions fit Excel's limit").
_fit(s::AbstractString) = length(s) <= 253 ? String(s) : first(s, 250) * "..."

# Excel's rules for a function name, as for a defined name: a letter, _ or \ first, then
# letters, digits, _ and . (any script's letters: jl.σ registers).
const _EXCEL_NAME = r"^[\p{L}_\\][\p{L}\p{N}_.]*$"
# Names Excel reads as a cell reference: A1 style (any column letters 1-3 long, any row - a
# little stricter than Excel, which allows XFE1), and R1C1 style (R, C, R1, C1, RC, R1C1...).
# Applied to the part before the first "." too: Excel turned =R1C1.x(1,2) into =$A$1.x(1,2),
# a data type's field, giving #FIELD! (checked 2026-11-25; XFD1.x and RC.x happened to work).
const _CELL_REFERENCE = r"^([A-Za-z]{1,3}[0-9]+|[RrCc]|[Rr][0-9]*[Cc][0-9]*|[Rr][0-9]+|[Cc][0-9]+)$"i

# Throws if Excel can't register excel_name, or if registering it would replace a function
# other than the one exposed under that name before. Called before anything is recorded.
function _check_excel_name(excel_name::String, is_default::Bool)
    how = is_default ? " Give a valid Excel name with `name=`." : ""
    if length(excel_name) > 255
        throw(ArgumentError("`$excel_name` is too long for an Excel function name (more than 255 characters).$how"))
    elseif !occursin(_EXCEL_NAME, excel_name)
        throw(ArgumentError("`$excel_name` isn't a valid Excel function name: it may contain only letters, digits, `_` and `.`, and must start with a letter or `_`.$how"))
    elseif occursin(_CELL_REFERENCE, first(split(excel_name, '.')))
        throw(ArgumentError("`$excel_name` can't be an Excel function name: Excel would read it, or its part before the first `.`, as a cell reference."))
    elseif startswith(lowercase(excel_name), "julia")
        throw(ArgumentError("`$excel_name` can't be an Excel function name: names beginning with `Julia` are reserved for the add-in's own functions."))
    end
    for other in keys(_exposed)
        if other != excel_name && lowercase(other) == lowercase(excel_name)
            throw(ArgumentError("`$excel_name` would replace `$other` in Excel, which ignores case in function names. Give a different name with `name=`."))
        end
    end
end

"""
    exposed_functions()

SPIKE. The exposed functions, for the add-in: one row per function, with columns Excel name,
Julia name, description, argument names, argument descriptions (each joined by "\\x1f") and
the number of required arguments. `nothing` when there are none: Excel has no empty
arrays, so the add-in can't decode one.
"""
function exposed_functions()
    isempty(_exposed) && return nothing
    names = sort!(collect(keys(_exposed)))
    result = Matrix{Any}(undef, length(names), 6)
    for (i, n) in enumerate(names)
        e = _exposed[n]
        result[i, :] = [n, e.julia, e.description, join(e.argnames, _ARGSEP), join(e.argdescriptions, _ARGSEP), e.nrequired]
    end
    result
end
