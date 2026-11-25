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

"""
    expose(fn_name; name=nothing, description="", argdescriptions=String[])

SPIKE. Make the Julia function named `fn_name` callable from Excel as a worksheet function,
by default `jl.<fn_name>`. Its arguments are named as in the function's first method.
Returns the Excel name. The add-in registers exposed functions when asked (spike: the
JuliaRegisterFunctions command).
"""
function expose(fn_name::AbstractString; name=nothing, description::AbstractString="",
                argdescriptions=String[])
    f = Main.eval(Meta.parse(fn_name))
    m = first(methods(f))
    argnames = [string(a) for a in Base.method_argnames(m)[2:end]]
    excel_name = name === nothing ? "jl." * fn_name : String(name)
    _exposed[excel_name] = (julia=String(fn_name), description=String(description),
        argnames=argnames, argdescriptions=collect(String, argdescriptions))
    excel_name
end

"""
    exposed_functions()

SPIKE. The exposed functions, for the add-in: one row per function, with columns Excel name,
Julia name, description, argument names and argument descriptions (each joined by "\\x1f").
"""
function exposed_functions()
    names = sort!(collect(keys(_exposed)))
    result = Matrix{Any}(undef, length(names), 5)
    for (i, n) in enumerate(names)
        e = _exposed[n]
        result[i, :] = [n, e.julia, e.description, join(e.argnames, _ARGSEP), join(e.argdescriptions, _ARGSEP)]
    end
    result
end
