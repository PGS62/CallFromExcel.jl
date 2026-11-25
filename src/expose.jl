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
by default `jl.<fn_name>`. Its arguments are named as in the function's first method.
Returns the Excel name. The add-in registers exposed functions after `JuliaLaunch`,
`JuliaInclude` and `JuliaEval`, and when the ribbon's Register Functions button is clicked
(for functions exposed from the REPL).
"""
function expose(fn_name::AbstractString; name=nothing, description::AbstractString="",
                argdescriptions=String[])
    f = Main.eval(Meta.parse(fn_name))
    m = first(methods(f))
    argnames = [string(a) for a in Base.method_argnames(m)[2:end]]
    excel_name = name === nothing ? "jl." * fn_name : String(name)
    _check_excel_name(excel_name, name === nothing)
    _exposed[excel_name] = (julia=String(fn_name), description=String(description),
        argnames=argnames, argdescriptions=collect(String, argdescriptions))
    _exposed_count[] += 1
    excel_name
end

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
Julia name, description, argument names and argument descriptions (each joined by "\\x1f").
`nothing` when there are none: Excel has no empty arrays, so the add-in can't decode one.
"""
function exposed_functions()
    isempty(_exposed) && return nothing
    names = sort!(collect(keys(_exposed)))
    result = Matrix{Any}(undef, length(names), 5)
    for (i, n) in enumerate(names)
        e = _exposed[n]
        result[i, :] = [n, e.julia, e.description, join(e.argnames, _ARGSEP), join(e.argdescriptions, _ARGSEP)]
    end
    result
end
