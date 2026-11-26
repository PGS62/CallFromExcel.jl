#=
Object handles - see "Object handles" in the documentation.

The add-in's JuliaCallHandle and JuliaEval(..., TRUE) call handle_create/handle_eval, which store the
result here and return a handle string "jl:<label>#<id>@<session>". The add-in recognises a
handle by that pattern (HandlePattern.cs) and sends it to Julia as the wire-format type
"O<id>@<session>" (see decode_from_xl), which is replaced by the stored object. Excel-DNA's RTD
machinery tells the add-in when no formula uses a worksheet handle any more, and the add-in
then calls handle_free; VBA frees the handles it creates with JuliaHandleRelease(All).

The session tag is chosen when the first handle is created, so a handle from an earlier
Julia process (saved in a workbook, or held by an Excel that outlived its Julia) gives
"not found" rather than silently naming an unrelated object with the same id.
=#

const _handles = Dict{Int,Any}()
const _handle_counter = Ref(0)
const _handle_session = Ref("")

function _session_tag()
    isempty(_handle_session[]) && (_handle_session[] = randstring(RandomDevice(), 4))
    _handle_session[]
end

"""
    handle_label(x)

The label in a handle for `x`, `jl:<label>#<id>@<session>`, for people reading the sheet -
looking up the object ignores it. By default the name of `x`'s type, without its module or
parameters: `Array` for a `Matrix{Float64}`, `Dict` for a `Dict{String,Int}`.

Add methods for your own types to choose a more informative label, for example:

```julia
CallFromExcel.handle_label(x::YieldCurve) = "Curve_\$(x.currency)"
```

so that a handle to a US dollar curve reads `jl:Curve_USD#3@k3x9`. The label is worked out
once, when the handle is created. Any `#`, `@`, whitespace or control character in it is
replaced by `_`, since `#` and `@` mark where the id and session start.
"""
handle_label(x) = string(nameof(typeof(x)))

# The label as it appears in a handle: handle_label's result with the characters a label can't
# contain replaced by '_'. Every character left must be one HandlePattern.cs accepts in a
# label: it rejects '#', '@', whitespace (.NET's \s, which includes Unicode separators) and
# control and format characters (\p{Cc}, \p{Cf}), all of which are replaced here.
_clean_label(label) = replace(string(label), r"[#@\s\p{Z}\p{Cc}\p{Cf}]" => "_")

function _store_handle(x)
    id = (_handle_counter[] += 1)
    _handles[id] = x
    "jl:" * _clean_label(handle_label(x)) * "#" * string(id) * "@" * _session_tag()
end

"""
    handle_create(fn_name::String, args...)

Call the function named `fn_name` with `args`, store the result and return its handle. For
the add-in's JuliaCallHandle, which covers both JuliaCall and JuliaCallNamedArgs: if any
argument is the string ";", the arguments are split into positional and named ones as for
JuliaCallNamedArgs (see `split_named_args`); otherwise all are positional, as for JuliaCall.
A trailing "." on the function name broadcasts, as in `srv_call_inner`: the object stored is
the broadcast result, and named arguments are passed unchanged to every call.
"""
function handle_create(fn_name::String, args...)
    broadcasting = endswith(fn_name, ".")
    broadcasting && (fn_name = chop(fn_name))
    fn = Main.eval(Meta.parse(fn_name))
    positional, named = any(x -> x isa AbstractString && x == ";", args) ?
        split_named_args(collect(Any, args)) : (args, NamedTuple())
    result = if broadcasting
        Base.invokelatest(broadcast, (xs...) -> fn(xs...; named...), positional...)
    else
        Base.invokelatest(fn, positional...; named...)
    end
    _store_handle(result)
end

"""
    handle_eval(code::String)

Evaluate `code`, store the result and return its handle. For the add-in's
`JuliaEval(code, TRUE)`.
"""
handle_eval(code::String) = _store_handle(Main.eval(Meta.parseall(code)))

# "17@k3x9" -> (17, "k3x9")
function _parse_handle_key(key::AbstractString)
    at = findlast('@', key)
    at === nothing && error("Malformed handle reference '$key'")
    parse(Int, key[1:prevind(key, at)]), key[nextind(key, at):end]
end

# Called by decode_from_xl for the wire type "O<id>@<session>".
function _lookup_handle(key::AbstractString)
    id, session = _parse_handle_key(key)
    if session == _handle_session[] && haskey(_handles, id)
        return _handles[id]
    end
    error("Handle #$key not found - it has been released, or Julia has restarted since it was created")
end

"""
    handle_free(key::String)

Drop the object for handle key "<id>@<session>". A key from another session is ignored.
Returns the number of handles still stored.
"""
function handle_free(key::String)
    id, session = _parse_handle_key(key)
    session == _handle_session[] && delete!(_handles, id)
    length(_handles)
end

"""
    handle_session()

This Julia process's session tag, the part of its handles after "@". The add-in compares it
with its live handles' after a launch, to tell whether Julia has restarted since they were
created.
"""
handle_session() = _session_tag()

"Number of objects currently stored for handles."
handle_count() = length(_handles)

"Drop every stored object."
handle_clear() = (empty!(_handles); 0)
