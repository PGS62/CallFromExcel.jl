#=
SPIKE (branch spike/object-handles): object handles, see "Object handles" in Enhancements.md.

The add-in's JuliaCallHandle and JuliaEval(..., TRUE) call handle_create/handle_eval, which store the
result here and return a handle string "jl:<label>#<id>@<session>". The add-in keeps the set
of handles it has issued and still holds; when one of them is passed to any function it is
sent as the wire-format type "O<id>@<session>" (see decode_from_xl), which is replaced by
the stored object. Excel-DNA's RTD machinery tells the add-in when no formula uses a handle
any more, and the add-in then calls handle_free.

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

Label shown in a handle for `x`, for people reading the sheet; lookup ignores it.
"""
handle_label(x) = replace(string(nameof(typeof(x))), r"[^A-Za-z0-9_]" => "")

function _store_handle(x)
    id = (_handle_counter[] += 1)
    _handles[id] = x
    "jl:" * handle_label(x) * "#" * string(id) * "@" * _session_tag()
end

"""
    handle_create(fn_name::String, args...)

Call the function named `fn_name` with `args`, store the result and return its handle. For
the add-in's JuliaCallHandle, which covers both JuliaCall and JuliaCallNamedArgs: if any
argument is the string ";", the arguments are split into positional and named ones as for
JuliaCallNamedArgs (see `split_named_args`); otherwise all are positional, as for JuliaCall.
"""
function handle_create(fn_name::String, args...)
    fn = Main.eval(Meta.parse(fn_name))
    if any(x -> x isa AbstractString && x == ";", args)
        positional, named = split_named_args(collect(Any, args))
        _store_handle(Base.invokelatest(fn, positional...; named...))
    else
        _store_handle(Base.invokelatest(fn, args...))
    end
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
    error("Handle #$key not found - recalculate the cell that created it")
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

"Number of objects currently stored for handles."
handle_count() = length(_handles)

"Drop every stored object."
handle_clear() = (empty!(_handles); 0)
