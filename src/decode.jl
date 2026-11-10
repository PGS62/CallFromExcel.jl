#=
decode_from_xl is the inverse of encode_for_xl (in encode.jl). It decodes a string in the
CallFromExcel wire format back to a Julia value. The add-in's WireFormatEncode.Encode
(dotnet-addin, WireFormatEncode.cs) uses the same format to encode arguments before POSTing them
to the /call HTTP endpoint, handled by srv_call_inner (in comms.jl).

Type indicators:
 #   Float64  (followed by 16 hex chars, IEEE-754 bit pattern)
 S   Float32  (followed by 8 hex chars)
 £   String   (followed by the string content)
 T   Bool true
 F   Bool false
 E   Missing  (VBA Empty)
 N   Nothing  (VBA Null)
 %   Int16    (followed by decimal)
 &   Int32    (followed by decimal)
 ^   Int64    (followed by decimal)
 B   UInt8    (followed by decimal; VBA's Byte, its only unsigned type)
 D   Date     (followed by Excel serial integer: days since 1899-12-30)
 G   DateTime (followed by 16 hex chars representing Excel serial as Float64)
 !   Error    (followed by Excel error number, e.g. 2042 for #N/A; decodes to ExcelError)
 *   Array    (*<rank>,<d1>[,<d2>];<len1>,<len2>,...,;<elements> column-major)
 H   Dict     (H<count>;<key1_len>,<val1_len>,...,;<key1><val1>... pairs column-order)
 V   Array of Float64 only (V<rank>,<d1>,...;<16 hex chars per element>; see decode_xl_array_v
     below, and WireFormatEncode.TryEncodeArrayAsV in the add-in, which produces this format)
 L   Array of Int64 only - the Int64 counterpart of V, produced only by Julia's encode_for_xl

See also the add-in's WireFormatEncode.Encode, which serialises i.e. inverts this.
=#

"""
    decode_from_xl(s::String)

Decode a string in the CallFromExcel wire format and return the corresponding Julia value.
Inverse of `encode_for_xl`.
"""
function decode_from_xl(s::String)
    isempty(s) && return missing
    c = s[1]
    if c == '#'                    #Double -> Float64
        hex_to_float64(s[2:end])
    elseif c == '£'                #String -> String 
        s[nextind(s, 1):end]       # skip the 2-byte £ prefix
    elseif c == 'T'                #Boolean True -> Bool true
        true
    elseif c == 'F'                #Boolean False -> Bool false
        false
    elseif c == 'E'                #Empty -> Missing
        missing
    elseif c == 'N'                #Null -> Nothing
        nothing
    elseif c == '&'                #Long -> Int32
        parse(Int32, s[2:end])
    elseif c == '%'                #Integer -> Int16
        parse(Int16, s[2:end])
    elseif c == '^'                #LongLong -> Int64
        parse(Int64, s[2:end])
    elseif c == 'B'                #Byte -> UInt8
        parse(UInt8, s[2:end])
    elseif c == 'S'                #Single -> Float32
        hex_to_float32(s[2:end])
    elseif c == 'D'                #Date, no time component -> Date
        Dates.Date(1899, 12, 30) + Dates.Day(parse(Int, s[2:end]))
    elseif c == 'G'                #Date, with time component -> DateTime
        Dates.DateTime(1899, 12, 30) +
            Dates.Millisecond(round(Int64, hex_to_float64(s[2:end]) * 86_400_000))
    elseif c == '!'                #Error -> ExcelError
        ExcelError(parse(Int, s[2:end]))
    elseif c == '*'                #Array -> Array
        decode_xl_array(s)
    elseif c == 'H'                #Dictionary -> Dict
        decode_xl_dict(s)
    elseif c == 'V'                #Array of Doubles (compact) -> Array{Float64,N}
        decode_xl_array_v(s)
    elseif c == 'L'                #Array of LongLongs (compact) -> Array{Int64,N}
        decode_xl_array_l(s)
    else
        error("decode_from_xl: unknown type indicator $(repr(c)) in '$(first(s, 50))'")
    end
end

"""
    xl_advance(s, start, xl_len) -> Int

Starting at byte position `start` in `s`, advance `xl_len` xl-length units and return the
byte index of the first character past the advanced region. An xl-length unit is one per
BMP character (< U+10000) and two per supplementary character (≥ U+10000), matching the
add-in's `string.Length` (UTF-16 code units, as VBA's `Len()`) and Julia's `xl_length()`.

UTF-8 character-boundary finding itself is left entirely to `nextind` - the only thing
genuinely specific to this function is converting "how many bytes did that character take"
into "how many xl-units does it represent", which Julia has no builtin for.
"""
function xl_advance(s::String, start::Int, xl_len::Int)::Int
    pos = start
    remaining = xl_len
    while remaining > 0
        remaining -= codepoint(s[pos]) >= 0x10000 ? 2 : 1   # BMP: 1 xl-unit, supplementary: 2
        pos = nextind(s, pos)
    end
    pos
end

# Whether an array element starting with this byte can contain non-ASCII characters: a string
# ('£', whose UTF-8 lead byte is 0xc2), or a nested array or dictionary. Every other element type
# is pure ASCII, so its length in bytes equals its xl_length and xl_advance isn't needed.
may_be_non_ascii(b::UInt8) = b == 0xc2 || b == UInt8('*') || b == UInt8('H')

"""
    decode_xl_array(s::String)

Decode an array-encoded string (starting with `*`) from the CallFromExcel wire format.
Returns a typed array when all elements share the same type, otherwise `Array{Any}`.
"""
function decode_xl_array(s::String)
    # Format: *<rank>,<d1>[,<d2>];<len1>,<len2>,...,;<elements>
    p1 = findfirst(isequal(';'), s)::Int
    p2 = findnext(isequal(';'), s, p1 + 1)::Int

    # Parse rank and dims (all ASCII, safe to byte-index)
    parts = split(s[2:p1-1], ',')
    rank  = parse(Int, parts[1])
    dims  = tuple(parse.(Int, parts[2:end])...)
    n     = prod(dims; init=1)

    if n == 0
        return rank == 1 ? Any[] : Array{Any}(undef, dims...)
    end

    # Parse lengths section (comma-separated with trailing comma, all ASCII)
    lengths = Vector{Int}(undef, n)
    pos = p1 + 1
    for i in 1:n
        comma = findnext(isequal(','), s, pos)::Int
        lengths[i] = parse(Int, s[pos:comma-1])
        pos = comma + 1
    end

    # Decode elements from contents section
    elements = Vector{Any}(undef, n)
    pos = p2 + 1
    for i in 1:n
        # Skipping xl_advance where possible matters for a "V" element of 1,000,000 Doubles
        # (16,000,000 characters), which xl_advance would otherwise step through one by one.
        len = lengths[i]
        next_pos = len > 0 && may_be_non_ascii(codeunit(s, pos)) ? xl_advance(s, pos, len) : pos + len
        # next_pos is the byte index of the start of the *next* character (or one past the end
        # of s) - a valid character-boundary index, but next_pos-1 is not necessarily one: if
        # this element's last character is multi-byte, next_pos-1 lands on a continuation byte,
        # which String's getindex/SubString reject even though it's genuinely the last byte of a
        # complete character. prevind(s, next_pos) gives the start index of that same last
        # character instead, which is what a String range's end must be.
        elements[i] = decode_from_xl(s[pos:prevind(s, next_pos)])
        pos = next_pos
    end

    # Reshape for 2D; try to return a typed array
    arr = rank == 1 ? elements : reshape(elements, dims)
    _maybe_typed(arr)
end

"""
    decode_xl_array_v(s::String)

Decode a "V"-format array-of-Float64 string (starting with `V`) - the compact encoding the add-in's
`WireFormatEncode.TryEncodeArrayAsV` (dotnet-addin, WireFormatEncode.cs) produces for a 1- to
9-dimensional array whose elements are all Doubles: no per-element type indicator or length, since
every element is always exactly 16 hex characters. This is the format of almost any range of
numbers passed to `JuliaCall`. Unlike `decode_xl_array`, this never needs `_maybe_typed` - the
result is a genuinely typed `Array{Float64,N}` directly.

The hex is big-endian (matching `hex_to_float64`'s scalar convention), as produced by the add-in's
`Hex.WriteBulkHex`. The same convention `encode_for_xl(x::Array{Float64,N})` (encode.jl) uses for
the Julia -> Excel direction.
"""
decode_xl_array_v(s::String) = decode_hex_array(Float64, s)

"""
    decode_xl_array_l(s::String)

Decode an "L"-format array-of-Int64 string - the Int64 counterpart of "V", produced by
`encode_for_xl(x::Array{Int64,N})`. Only Julia ever produces this format, so this exists so
that `decode_from_xl` remains the inverse of `encode_for_xl`.
"""
decode_xl_array_l(s::String) = decode_hex_array(Int64, s)

# Value of each hex digit (either case), indexed by byte + 1; 0xff for any other byte.
const HEX_NIBBLE = let t = fill(0xff, 256)
    for (i, c) in enumerate(b"0123456789abcdef")
        t[c+1] = i - 1
    end
    for (i, c) in enumerate(b"ABCDEF")
        t[c+1] = i + 9
    end
    t
end

"""
    decode_hex_array(T, s::String)

Shared decoder for the "V" (`T = Float64`) and "L" (`T = Int64`) formats:
`<tag><rank>,<d1>,...;<16 big-endian hex chars per element, column-major, no delimiters>`.

Reads the hex straight from the string's bytes into the result, in one pass. This replaced
`hex2bytes` + `bswap.` + `collect`, which copied the payload three times and took ~60-100 ms
for 1,000,000 elements, against ~4 ms for this loop.
"""
function decode_hex_array(::Type{T}, s::String) where {T<:Union{Float64,Int64}}
    cu = codeunits(s)
    p1 = findfirst(==(UInt8(';')), cu)
    p1 === nothing && throw(ArgumentError("'$(s[1])'-format array has no ';' after its dimensions"))
    parts = split(SubString(s, 2, p1 - 1), ',')
    rank = parse(Int, parts[1])
    length(parts) == rank + 1 || throw(ArgumentError("'$(s[1])'-format array has rank $rank but $(length(parts) - 1) dimension(s)"))
    dims = ntuple(i -> parse(Int, parts[i+1]), rank)
    n = prod(dims; init=1)
    length(cu) - p1 == 16n || throw(ArgumentError(
        "'$(s[1])'-format array of $n element(s) should have $(16n) hex characters, but has $(length(cu) - p1)"))

    out = Vector{T}(undef, n)
    bad = 0x00   # ORs together every nibble value, so it is 0xff if and only if any byte wasn't a hex digit
    @inbounds for i in 1:n
        base = p1 + 16(i - 1)
        u = zero(UInt64)
        for j in 1:16
            nibble = HEX_NIBBLE[cu[base+j]+1]
            bad |= nibble
            u = (u << 4) | nibble
        end
        out[i] = reinterpret(T, u)
    end
    bad == 0xff && throw(ArgumentError("'$(s[1])'-format array contains a character that is not a hex digit"))
    rank == 1 ? out : reshape(out, dims)
end

"""
    decode_xl_dict(s::String)

Decode a dict-encoded string (starting with `H`) from the CallFromExcel wire format, with keys and
values decoded by `decode_from_xl`. Returns a typed `Dict{K,V}` when every key shares a concrete
type `K` and every value shares a concrete type `V` (see `_maybe_typed` below - the same treatment
`decode_xl_array` already gives arrays), otherwise `Dict{Any,Any}`.
"""
function decode_xl_dict(s::String)
    # Format: H<count>;<key1_len>,<val1_len>,...,;<key1><val1>...
    p1 = findfirst(isequal(';'), s)::Int
    p2 = findnext(isequal(';'), s, p1 + 1)::Int

    n = parse(Int, s[2:p1-1])   # number of key-value pairs
    n == 0 && return Dict{Any,Any}()

    # Parse 2n lengths (alternating: key_len, val_len, ...)
    lengths = Vector{Int}(undef, 2n)
    pos = p1 + 1
    for i in 1:2n
        comma = findnext(isequal(','), s, pos)::Int
        lengths[i] = parse(Int, s[pos:comma-1])
        pos = comma + 1
    end

    # Decode key-value pairs from contents section
    result = Dict{Any,Any}()
    pos = p2 + 1
    for i in 1:n
        key_end = xl_advance(s, pos, lengths[2i-1])
        # See the matching comment in decode_xl_array: prevind, not -1, since a multi-byte last
        # character would otherwise leave the range ending on a continuation byte.
        key     = decode_from_xl(s[pos:prevind(s, key_end)])
        val_end = xl_advance(s, key_end, lengths[2i])
        val     = decode_from_xl(s[key_end:prevind(s, val_end)])
        result[key] = val
        pos = val_end
    end
    _maybe_typed(result)
end

# Convert Array{Any} to a typed array when all elements share the same concrete type.
function _maybe_typed(a::Array{Any})
    isempty(a) && return a
    T = typeof(a[1])
    all(x -> typeof(x) === T, a) || return a
    try
        convert(Array{T}, a)
    catch
        a
    end
end

# Convert Dict{Any,Any} to a typed Dict{K,V} when every key shares a concrete type K and every
# value shares a concrete type V - mirrors _maybe_typed's treatment of arrays above.
function _maybe_typed(d::Dict{Any,Any})
    isempty(d) && return d
    ks = collect(keys(d))
    vs = collect(values(d))
    K = typeof(ks[1])
    V = typeof(vs[1])
    (all(x -> typeof(x) === K, ks) && all(x -> typeof(x) === V, vs)) || return d
    try
        Dict{K,V}(d)
    catch
        d
    end
end
