using CallFromExcel
using Dates
using SHA
using Test

# Must match the hash definition in dotnet-addin/Update-Dist.ps1 exactly.
function source_hash(projectdir)
    relpaths = String[]
    for (root, _, files) in walkdir(projectdir), f in files
        rel = replace(relpath(joinpath(root, f), projectdir), '\\' => '/')
        startswith(rel, "bin/") || startswith(rel, "obj/") || push!(relpaths, rel)
    end
    buf = IOBuffer()
    for rel in sort(relpaths)
        write(buf, rel, '\n', filter(!=(0x0d), read(joinpath(projectdir, rel))), '\n')
    end
    bytes2hex(sha256(take!(buf)))
end

@testset "dotnet-addin/dist is up to date with the add-in's source" begin
    addin = joinpath(@__DIR__, "..", "dotnet-addin")
    @test isfile(joinpath(addin, "dist", "CallFromExcel.AddIn64-packed.xll"))
    @test isfile(joinpath(addin, "dist", "CallFromExcel.AddIn.dll"))
    # On failure: run `dotnet build -c Release` in dotnet-addin/CallFromExcel.AddIn, then commit dist/.
    @test source_hash(joinpath(addin, "CallFromExcel.AddIn")) ==
          strip(read(joinpath(addin, "dist", "source-hash.txt"), String))
    # The add-in reports its version in the ribbon's Help tooltip. On failure: set
    # AssemblyInformationalVersion in AssemblyInfo.cs to Project.toml's version, and rebuild.
    assemblyinfo = read(joinpath(addin, "CallFromExcel.AddIn", "src", "AssemblyInfo.cs"), String)
    m = match(r"AssemblyInformationalVersion\(\"([^\"]+)\"\)", assemblyinfo)
    @test m !== nothing && m[1] == string(pkgversion(CallFromExcel))
end

# isequal alone is type-blind (isequal(2, 2.0) is true), so it would miss a type morph introduced
# on the wire (e.g. decode returning Int32 where the input was Int64). === is too strict the other
# way: it's identity comparison for mutable types, so it fails for every array/Dict round trip even
# when content and type both match, since decode_from_xl always builds a fresh object. Checking
# typeof equality alongside isequal gets both right - exact type, and isequal's correct handling of
# NaN/missing/nested structures for the value itself.
function round_trip(x)
    y = CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl(x))
    typeof(y) == typeof(x) && isequal(y, x)
end

@testset "CallFromExcel.jl" begin
    # The exact wire forms the add-in's WireFormatDecode.Decode expects
    @test CallFromExcel.encode_for_xl(1) == "^1"
    @test CallFromExcel.encode_for_xl(1.0) == "#3FF0000000000000"
    @test CallFromExcel.encode_for_xl(Int8(1)) == "%1"
    @test CallFromExcel.encode_for_xl(Int16(1)) == "%1"
    @test CallFromExcel.encode_for_xl(Int32(1)) == "&1"
    @test CallFromExcel.encode_for_xl(Int64(1)) == "^1"
    @test CallFromExcel.encode_for_xl(UInt8(200)) == "B200"
    @test CallFromExcel.decode_from_xl("B200") == UInt8(200)
    @test CallFromExcel.encode_for_xl(true) == "T"
    @test CallFromExcel.encode_for_xl(false) == "F"
    @test CallFromExcel.encode_for_xl("foo") == "£foo"
    @test CallFromExcel.encode_for_xl('x') == "£x"
    @test CallFromExcel.encode_for_xl(:x) == "£:x"
    @test CallFromExcel.encode_for_xl(nothing) == "E"
    @test CallFromExcel.encode_for_xl(missing) == "E"
    @test CallFromExcel.encode_for_xl(Inf) == "!2036"
    @test CallFromExcel.encode_for_xl(-Inf) == "!2036"
    @test CallFromExcel.encode_for_xl(NaN) == "!2042"
    # Complex numbers return as text in the format of Excel's COMPLEX function, for IMREAL etc.
    @test CallFromExcel.encode_for_xl(3 + 4im) == "£3+4i"
    @test CallFromExcel.encode_for_xl(3 - 4im) == "£3-4i"
    @test CallFromExcel.encode_for_xl(4im) == "£4i"
    @test CallFromExcel.encode_for_xl(im) == "£i"
    @test CallFromExcel.encode_for_xl(-im) == "£-i"
    @test CallFromExcel.encode_for_xl(1.5 + im) == "£1.5+i"
    @test CallFromExcel.encode_for_xl(3.0 + 0.0im) == "£3"
    @test CallFromExcel.encode_for_xl(complex(-0.0, -0.0)) == "£0"
    @test CallFromExcel.encode_for_xl(π + im) == "£3.14159265358979+i"
    @test CallFromExcel.encode_for_xl(complex(1e-5, 2e20)) == "£1e-05+2e+20i"
    @test CallFromExcel.encode_for_xl(ComplexF32(0.5, -0.25)) == "£0.5-0.25i"
    @test CallFromExcel.encode_for_xl(complex(Inf, 1.0)) == "!2036"
    @test CallFromExcel.encode_for_xl(complex(1.0, NaN)) == "!2036"
    @test CallFromExcel.encode_for_xl([1 + 2im, 3.0 + 0im]) == "*1,2;5,2,;£1+2i£3"
    @test CallFromExcel.encode_for_xl(CallFromExcel.ExcelError(2007)) == "!2007"
    @test CallFromExcel.decode_from_xl("!2007") == CallFromExcel.ExcelError(2007)
    @test CallFromExcel.encode_for_xl(Date("2021-11-8")) == "D44508"
    @test CallFromExcel.encode_for_xl(DateTime("2021-11-8T12:00:00")) == "G40E5BB9000000000"
    @test CallFromExcel.encode_for_xl(Int64) == "£Int64"
    @test CallFromExcel.encode_for_xl(v"1.2.3") == "£1.2.3"
    @test CallFromExcel.encode_for_xl((1, 2)) == "L1,2;00000000000000010000000000000002"
    @test CallFromExcel.encode_for_xl([1, 2, 3]) == "L1,3;000000000000000100000000000000020000000000000003"
    @test CallFromExcel.encode_for_xl(Any[1, 2, 3.0, π]) == "*1,4;2,2,17,17,;^1^2#4008000000000000#400921FB54442D18"
    @test CallFromExcel.encode_for_xl([1, true, "x"]) == "*1,3;2,1,2,;^1T£x"
    @test CallFromExcel.encode_for_xl([1, [2, 3]]) == "*1,2;2,37,;^1L1,2;00000000000000020000000000000003"
    # Dict key order isn't guaranteed by the language, and can differ between Julia versions/hash
    # seeds - accept either encoding order rather than hardcoding one.
    @test CallFromExcel.encode_for_xl(Dict("a" => 1, "b" => 2)) in
          ("H2;2,2,2,2,;£a^1£b^2", "H2;2,2,2,2,;£b^2£a^1")

    @test round_trip(1)
    @test round_trip(1.0)
  #  @test round_trip(Int8(1)) # No dedicated Int8 wire type - morphs to Int16, the smallest one
    @test round_trip(Int16(1))
    @test round_trip(Int32(1))
    @test round_trip(Int64(1))
    @test round_trip(UInt8(200))
    @test round_trip(true)
    @test round_trip(false)
    @test round_trip("foo")
  #  @test round_trip('x') # Characters morph to String
  #  @test round_trip(:x)  # Symbols morph to String
  #  @test round_trip(nothing) # Nothing morphs to missing (Empty in Excel)
     @test round_trip(missing)
  #  @test round_trip(Inf)  # Morphs to ExcelError(2036) (#NUM! in Excel)
  #  @test round_trip(-Inf) # Morphs to ExcelError(2036) (#NUM! in Excel)
  #  @test round_trip(NaN)  # Morphs to ExcelError(2042) (#N/A in Excel)
    @test round_trip(Date("2021-11-8"))
    @test round_trip(DateTime("2021-11-8T12:00:00"))
  #  @test round_trip(Int64) # Types morph to String
  #  @test round_trip(v"1.2.3") # VersionNumber morph to String
  #  @test round_trip((1, 2)) # Tuples morph to Vector
    @test round_trip([1, 2, 3])
  #  @test round_trip(Any[1, 2, 3.0, π]) #Irrational morphs to nearest Float64
    @test round_trip([1, true, "x"])
    @test round_trip([1, [2, 3]])
    @test round_trip(Dict("a"=>1, "b"=>2))
    @test typeof(CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl(Dict("a"=>1, "b"=>2)))) == Dict{String,Int64}

    # Mixed-type dict correctly declines the typed conversion, staying Dict{Any,Any} - matching
    # decode_xl_array's own dynamic fallback for a Vector{Any} with mixed element types.
    @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl(Dict("a"=>1, "b"=>"x"))) isa Dict{Any,Any}

    # Regression: a string ending in a multi-byte character, as an array/dict element, used to
    # throw a StringIndexError on decode. decode_xl_array/decode_xl_dict computed each element's
    # end as (start of next element) - 1, which lands on a UTF-8 continuation byte whenever the
    # element's own last character is multi-byte - String's getindex/SubString require both range
    # endpoints to be valid character-start indices, even though that continuation byte genuinely
    # is the element's last byte. Fixed by using prevind instead of a raw -1. "£" (2 bytes) and
    # "𝔍" (a supplementary-plane character, 4 bytes / 2 UTF-16 units) both exercise this - as the
    # last array element, a non-last array element, a dict value and a dict key.
    @test round_trip(["identity", "£"])
    @test round_trip(["£", "identity"])
    @test round_trip(["identity", "𝔍"])
    @test round_trip(Dict("a" => "£"))
    @test round_trip(Dict("£" => "x"))

    # ExcelError - unlike Inf/NaN above (which only ever produce 2036/2042 from Julia's own
    # numeric values), every documented Excel error code round-trips exactly through Julia when it
    # originates as a genuine VBA error - see ExcelError's docstring (CallFromExcel.jl) for why this is
    # a distinct type rather than a String.
    for code in keys(CallFromExcel.EXCEL_ERROR_NAMES)
        @test CallFromExcel.decode_from_xl("!$code") == CallFromExcel.ExcelError(code)
        @test round_trip(CallFromExcel.ExcelError(code))
    end

    # "V" format - compact encoding for arrays of Float64 (see encode_for_xl(::Vector{Float64})/
    # (::Matrix{Float64}) and the dynamic AbstractArray fallback in src/encode.jl).
    @test CallFromExcel.encode_for_xl([1.0, 2.0, 3.0]) == "V1,3;3ff000000000000040000000000000004008000000000000"
    @test CallFromExcel.encode_for_xl([1.0 2.0; 3.0 4.0]) == "V2,2,2;3ff0000000000000400800000000000040000000000000004010000000000000"

    # Empty array and NaN/Inf-containing arrays must fall back to the general "*" format, not "V" -
    # every element of "V" is assumed to be an ordinary finite Double, which isn't true for these.
    @test CallFromExcel.encode_for_xl(Float64[]) == CallFromExcel.encode_array_general(Float64[])
    @test CallFromExcel.encode_for_xl([1.0, NaN, 3.0]) == CallFromExcel.encode_array_general([1.0, NaN, 3.0])
    @test CallFromExcel.encode_for_xl([1.0, Inf, 3.0]) == CallFromExcel.encode_array_general([1.0, Inf, 3.0])
    @test !startswith(CallFromExcel.encode_for_xl([1.0, NaN, 3.0]), "V")
    @test !startswith(CallFromExcel.encode_for_xl([1.0, Inf, 3.0]), "V")

    # Dynamic fallback: a Vector{Any}/Matrix{Any} that's actually all Float64 at runtime should
    # still get the fast "V" encoding, byte-for-byte identical to the concrete Vector{Float64}/
    # Matrix{Float64} case - dispatch alone (the type being literally Vector{Float64}) isn't the
    # only way to reach it.
    @test CallFromExcel.encode_for_xl(Any[1.0, 2.0, 3.0]) == CallFromExcel.encode_for_xl([1.0, 2.0, 3.0])
    @test CallFromExcel.encode_for_xl(Any[1.0 2.0; 3.0 4.0]) == CallFromExcel.encode_for_xl([1.0 2.0; 3.0 4.0])
    @test startswith(CallFromExcel.encode_for_xl(Any[1.0, 2.0, 3.0]), "V")

    # ...but the dynamic check must still correctly decline "V" (falling back to the general
    # format) for a Vector{Any} containing a NaN, a non-Float64 element, or with rank > 9 (the
    # add-in's WireFormatDecode.DecodeV reads the rank as a single digit).
    @test CallFromExcel.encode_for_xl(Any[1.0, 2.0, NaN]) == CallFromExcel.encode_for_xl([1.0, 2.0, NaN])
    @test !startswith(CallFromExcel.encode_for_xl(Any[1.0, 2.0, NaN]), "V")
    @test CallFromExcel.encode_for_xl(Any[1.0, 2, 3.0]) == CallFromExcel.encode_array_general(Any[1.0, 2, 3.0])
    @test !startswith(CallFromExcel.encode_for_xl(Any[1.0, 2, 3.0]), "V")
    @test CallFromExcel.encode_for_xl(Any[]) == CallFromExcel.encode_array_general(Any[])

    # Rank 3-9: one method (encode_for_xl(x::Array{Float64,N})) handles every rank, since
    # Vector{Float64}/Matrix{Float64} are just Array{Float64,1}/Array{Float64,2} and the bulk
    # byte-reinterpret doesn't care how many dimensions the buffer is viewed as. Decoded by the
    # add-in's WireFormatDecode.DecodeV, whose rank 3-9 branch builds the array with ParseDims and
    # AdvanceIndex.
    let a3d = reshape(collect(1.0:24.0), 2, 3, 4)
        @test startswith(CallFromExcel.encode_for_xl(a3d), "V3,2,3,4;")
        @test round_trip(a3d)
        @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl(a3d)) isa Array{Float64,3}
        # dynamic AbstractArray fallback also generalises to rank 3-9, not just 1-2
        @test CallFromExcel.encode_for_xl(Array{Any}(a3d)) == CallFromExcel.encode_for_xl(a3d)
    end

    # Rank > 9 must fall back to the general "*" format - the add-in's DecodeV handles only ranks 1-9.
    let a10d = reshape(collect(1.0:2.0^10), fill(2, 10)...)
        @test !startswith(CallFromExcel.encode_for_xl(a10d), "V")
        @test round_trip(a10d)
    end

    let x3d = fill(1.0, 2, 2, 2) |> a -> Array{Any}(a)
        @test startswith(CallFromExcel.encode_for_xl(x3d), "V3,")  # now gets "V", not the general format
        @test CallFromExcel.encode_for_xl(x3d) == CallFromExcel.encode_for_xl(fill(1.0, 2, 2, 2))
    end

    # decode_xl_array_v (src/decode.jl) - the Excel -> Julia direction, whose input the add-in's
    # WireFormatEncode.TryEncodeArrayAsV produces. Real round trips are possible since
    # decode_from_xl understands "V", unlike the encode-only assertions above.
    @test round_trip([1.0, 2.0, 3.0])
    @test round_trip([1.0 2.0 3.0; 4.0 5.0 6.0])  # non-square, to catch a row/column transposition bug
    @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl([1.0, 2.0, 3.0])) isa Vector{Float64}
    @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl([1.0 2.0; 3.0 4.0])) isa Matrix{Float64}

    # Hand-built big-endian hex in upper case (MSB-first per element, as the add-in's
    # Hex.WriteBulkHex produces) - confirms decode_xl_array_v works against the add-in's style of
    # input, not only against Julia's own (lower-case) encoder output.
    let be_hex(v) = uppercase(string(reinterpret(UInt64, v), base=16, pad=16))
        manual = "V1,3;" * be_hex(1.0) * be_hex(-2.5) * be_hex(3.14159265358979)
        @test CallFromExcel.decode_from_xl(manual) == [1.0, -2.5, 3.14159265358979]

        # Rank 3, exercising decode_xl_array_v's reshape path - what TryEncodeArrayAsV produces
        # for a 3D+ array passed from VBA.
        vals3d = collect(1.0:24.0)
        manual3d = "V3,2,3,4;" * join(be_hex.(vals3d))
        @test CallFromExcel.decode_from_xl(manual3d) == reshape(vals3d, 2, 3, 4)

        # Special values survive bit-for-bit, and malformed input is rejected rather than misread.
        specials = [-0.0, floatmin(Float64), nextfloat(0.0), floatmax(Float64), -1.0e300]
        decoded = CallFromExcel.decode_from_xl("V1,5;" * join(be_hex.(specials)))
        @test reinterpret(UInt64, decoded) == reinterpret(UInt64, specials)
        @test_throws ArgumentError CallFromExcel.decode_from_xl("V1,2;" * be_hex(1.0))            # too short
        @test_throws ArgumentError CallFromExcel.decode_from_xl("V1,1;" * be_hex(1.0) * "0")      # too long
        @test_throws ArgumentError CallFromExcel.decode_from_xl("V1,1;3FF000000000000G")          # not hex
        @test_throws ArgumentError CallFromExcel.decode_from_xl("V2,3;" * be_hex(1.0))            # rank/dims mismatch
        @test_throws ArgumentError CallFromExcel.decode_from_xl("L1,1;000000000000000-")          # not hex, "L" too
    end

    # "L" format - Int64 arrays, the Int64 counterpart of "V".
    @test CallFromExcel.encode_for_xl([1, -1]) == "L1,2;0000000000000001ffffffffffffffff"
    @test round_trip([1, 2, 3])
    @test round_trip([typemin(Int64) -1; 0 typemax(Int64)])
    @test round_trip(collect(reshape(1:24, 2, 3, 4)))
    @test startswith(CallFromExcel.encode_for_xl((1:3)'), "L2,1,3;")  # an AbstractArray with eltype Int64
    @test startswith(CallFromExcel.encode_for_xl(Int64[]), "*")      # empty arrays use the general format
    @test startswith(CallFromExcel.encode_for_xl(Int32[1, 2]), "*")  # only Int64 gets "L"

    # Numeric types with no VBA counterpart become the smallest VBA type holding every value of the
    # type, else Double - not their printed form as text (e.g. "0x00000003" for UInt32(3)).
    let enc = CallFromExcel.encode_for_xl
        @test enc(UInt16(3)) == enc(Int32(3))
        @test enc(UInt32(3)) == enc(Int64(3))
        @test enc(typemax(UInt32)) == enc(Int64(typemax(UInt32)))
        @test enc(Float16(1.5)) == enc(1.5f0)
        @test enc(UInt64(3)) == enc(3.0)
        @test enc(UInt128(3)) == enc(3.0)
        @test enc(big(2)^70) == enc(2.0^70)
        @test enc(big"1.5") == enc(1.5)
        @test enc(1//4) == enc(0.25)
        @test enc(UInt32[1, 2, 3]) == enc([1, 2, 3])
        @test enc(UInt64[1, 2, 3]) == enc([1.0, 2.0, 3.0])
        @test enc([1//2, 1//4]) == enc([0.5, 0.25])
        @test enc(Any[UInt32(1), "x"]) == enc(Any[1, "x"])
        @test enc(UInt32(1):UInt32(3)) == enc(1:3)
        @test CallFromExcel.decode_from_xl(enc(UInt64(1):UInt64(3))) == [1.0, 2.0, 3.0]
        @test enc(false:true) == enc(Any[false, true])   # not 0/1
    end

    # write_encoded! (used by the general array format) must write exactly encode_for_xl's output,
    # and return its xl_length, for every kind of element.
    for x in Any[0.0, -0.0, 1.5, -2.5e-300, 5e-324, floatmax(), NaN, Inf, -Inf, true, false, missing,
                 nothing, 0, -1, typemax(Int64), typemin(Int64), Int32(7), UInt8(9), "", "abc", "£€",
                 "😀x", Date(2021, 3, 11), DateTime(2021, 3, 11, 12), CallFromExcel.ExcelError(2042),
                 1.5f0, 2 + 3im, :sym, [1, 2], Any[1.0, "a"]]
        io = IOBuffer()
        n = CallFromExcel.write_encoded!(io, x)
        expected = CallFromExcel.encode_for_xl(x)
        @test String(take!(io)) == expected
        @test n == CallFromExcel.xl_length(expected)
    end

    # A DataFrame encodes column by column, but must match its header-plus-data matrix.
    let df = CallFromExcel.DataFrame(a = [1, 2], b = ["x", "£y"], c = [1.5, NaN], d = [true, missing])
        m = vcat(permutedims(Any[names(df)...]), Matrix{Any}(df))
        @test CallFromExcel.encode_for_xl(df) == CallFromExcel.encode_array_general(m)
    end
    @test CallFromExcel.encode_for_xl(CallFromExcel.DataFrame()) == "*2,1,0;;"

    # Invalid characters in strings - a lone surrogate, or invalid UTF-8 such as Latin-1 text read
    # into a String - are sent as U+FFFD, one per character. The add-in decodes UTF-8, which turned
    # them into one or more U+FFFD, so an array's element lengths came out wrong and the whole
    # result failed to decode (or, for invalid UTF-8, failed to encode).
    let lone = "a\ud800b", latin1 = String(UInt8[0x63, 0x61, 0x66, 0xe9])
        @test CallFromExcel.encode_for_xl([lone, latin1, "next"]) ==
              CallFromExcel.encode_for_xl(["a�b", "caf�", "next"])
        @test CallFromExcel.encode_for_xl(lone) == "£a�b"
        @test CallFromExcel.encode_for_xl(SubString(latin1, 1)) == "£caf�"
        @test CallFromExcel.encode_for_xl('\ud800') == "£�"
        @test CallFromExcel.encode_for_xl(Dict(lone => 1)) == CallFromExcel.encode_for_xl(Dict("a�b" => 1))
    end

    # "R" format - compact encoding for integer ranges: encodes only first/step/length, not every
    # element, so wire size is O(1) regardless of the range's length. Julia -> Excel only, so
    # decode_from_xl has no 'R' case: these are encode-only checks.
    @test CallFromExcel.encode_for_xl(1:5) == "RI,5,1,1;"
    @test CallFromExcel.encode_for_xl(5:3:47) == "RI,15,5,3;"  # non-1 step
    @test length(CallFromExcel.encode_for_xl(1:1_000_000)) < 20
    @test CallFromExcel.encode_for_xl(1:0) == CallFromExcel.encode_array_general(1:0)

    # Float ranges are sent exactly as the same values collected into a Vector. They used to be
    # sent as first/step/length, rebuilt as first + i*step, which differs in the last bit from
    # Julia's own elements for decimal steps: element 4 of 0:0.1:1 arrived as 0.30000000000000004.
    for r in (0:0.1:1, 0:0.01:1, range(0, 1, length=101), (1:1000) .* pi, 1.0:1.0:0.0, range(1.0, step=NaN, length=5))
        @test CallFromExcel.encode_for_xl(r) == CallFromExcel.encode_for_xl(collect(r))
    end

    # Defensive guard: nothing stops a third-party AbstractRange subtype implementing `step` to mean
    # something else. encode_for_xl checks last(x) == first(x) + (n-1)*step(x) before using "R",
    # here with a range whose step() claims (wrongly) it's evenly spaced by 1. Must fall back to
    # the general format, not silently send a linear approximation.
    struct LyingRange <: AbstractRange{Int64}
        data::Vector{Int64}
    end
    Base.length(r::LyingRange) = length(r.data)
    Base.first(r::LyingRange) = r.data[1]
    Base.last(r::LyingRange) = r.data[end]
    Base.step(r::LyingRange) = 1
    Base.getindex(r::LyingRange, i::Int) = r.data[i]
    Base.size(r::LyingRange) = (length(r.data),)

    let lr = LyingRange([1, 2, 100])  # not actually evenly spaced by 1
        @test !startswith(CallFromExcel.encode_for_xl(lr), "R")
        @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl(lr)) == [1, 2, 100]
    end

    # friendly_error (comms.jl) - srv_eval_inner's error path for a JuliaEval expression that
    # itself errors (as opposed to returning a value that fails to encode - see
    # _encode_result_for_xl above). Checks the exact text so a regression (a stray backtrace frame,
    # a changed cap, a wording change) is caught, not just "it's some string starting with #".
    # Julia 1.11 added a second line to MethodError's message; on 1.10 that line is empty, and
    # must not leave a double space.
    let msg = CallFromExcel.decode_from_xl(CallFromExcel.srv_eval_inner("1+\"1\""))
        if VERSION >= v"1.11"
            @test msg == "#MethodError: no method matching +(::Int64, ::String) The function `+` \
                         exists, but no method is defined for this combination of argument types. \
                         Julia REPL has more details and stacktrace!"
        else
            @test msg == "#MethodError: no method matching +(::Int64, ::String) \
                         Julia REPL has more details and stacktrace!"
        end
    end

    # display_results (comms.jl) - toggles REPL echoing of calls from Excel; check both the
    # setter's return-value contract and the getter round-trips it correctly. Ends on false (the
    # default) so it doesn't leak into any test that follows.
    @test CallFromExcel.display_results(true) == "Results from JuliaCall/JuliaEval will display in REPL"
    @test CallFromExcel.display_results() == true
    @test CallFromExcel.display_results(false) == "Results from JuliaCall/JuliaEval will not display in REPL"
    @test CallFromExcel.display_results() == false

    # last_question/last_answer/answer_again (comms.jl) - srv_eval_inner records the expression
    # text itself as last_question...
    CallFromExcel.decode_from_xl(CallFromExcel.srv_eval_inner("1+1"))
    @test CallFromExcel.last_question == "1+1"
    @test CallFromExcel.last_answer == 2
    @test CallFromExcel.answer_again() == 2

    # ...while srv_call_inner reconstructs a "fn_name(args_from_xl...)" call expression, with a
    # ".(...)" broadcasting form when the function name arrives with a trailing ".".
    let payload = CallFromExcel.encode_for_xl(Any["identity", 5])
        @test CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload)) == 5
        @test CallFromExcel.last_question == "identity(args_from_xl...)"
        @test CallFromExcel.last_answer == 5
        @test CallFromExcel.answer_again() == 5
    end
    let payload = CallFromExcel.encode_for_xl(Any["identity.", [1, 2, 3]])
        @test CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload)) == [1, 2, 3]
        @test CallFromExcel.last_question == "identity.(args_from_xl...)"
        @test CallFromExcel.answer_again() == [1, 2, 3]
    end

    # split_named_args (comms.jl) - finds the ";" separator by searching backwards from the
    # third-from-last argument in steps of 2, so a ";" among the positional arguments is never
    # taken as the separator, nor is a ";" value in the last pair. The separator is required, and
    # must be followed by a positive even number of arguments.
    let s = CallFromExcel.split_named_args
        msg(args) = try s(args); "" catch e; e.msg end
        @test startswith(msg(Any[1, 2]), "No \";\" argument found")
        @test startswith(msg(Any[]), "No \";\" argument found")
        @test startswith(msg(Any[";"]), "No named arguments follow")
        @test startswith(msg(Any[1, ";"]), "No named arguments follow")
        @test endswith(msg(Any[1, ";", "rev", false, "foo"]), "an odd number (3) of arguments follow it")
        @test endswith(msg(Any[1, ";", "rev"]), "an odd number (1) of arguments follow it")
        @test s(Any[1, ";", "a", 10]) == (Any[1], (a=10,))
        @test s(Any[";", "a", 10, "b", 20]) == (Any[], (a=10, b=20))
        @test s(Any["x;y", ";", ";", "limit", 2]) == (Any["x;y", ";"], (limit=2,))
        @test s(Any[1, ";", "a", 10, "delim", ";"]) == (Any[1], (a=10, delim=";"))
        @test_throws ArgumentError s(Any[1, ";", "not valid", 10])
        @test_throws ArgumentError s(Any[1, ";", 3.0, 10])
        @test_throws ArgumentError s(Any[1, ";", "a", 10, "a", 20])
    end

    # srv_call_inner with named_args=true (requests to /callkw, from JuliaCallNamedArgs) passes the
    # named arguments as keyword arguments, saved in kwargs_from_xl, including when broadcasting.
    let payload = CallFromExcel.encode_for_xl(Any["split", "a;b;c", ";", ";", "limit", 2])
        @test CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload; named_args=true)) == ["a", "b;c"]
        @test CallFromExcel.last_question == "split(args_from_xl...; kwargs_from_xl...)"
        @test CallFromExcel.kwargs_from_xl == (limit=2,)
        @test CallFromExcel.answer_again() == ["a", "b;c"]
    end
    let payload = CallFromExcel.encode_for_xl(Any["round.", [1.234, 5.678], ";", "digits", 1])
        @test CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload; named_args=true)) == [1.2, 5.7]
        @test CallFromExcel.last_question == "round.(args_from_xl...; kwargs_from_xl...)"
        @test CallFromExcel.answer_again() == [1.2, 5.7]
    end
    let payload = CallFromExcel.encode_for_xl(Any["round", 1.5, ";", "digitz", 1])
        @test startswith(CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload; named_args=true)), "#MethodError")
    end
    let payload = CallFromExcel.encode_for_xl(Any["round", 1.5, ";", "a", 1, "a", 2])
        @test startswith(CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload; named_args=true)),
                         "#ArgumentError: Named argument `a` is given more than once")
    end
    let payload = CallFromExcel.encode_for_xl(Any["sort", [3, 1, 2], ";", "rev", false, "foo"])
        @test startswith(CallFromExcel.decode_from_xl(CallFromExcel.srv_call_inner(payload; named_args=true)),
                         "#ArgumentError: Named arguments must be given as name-value pairs")
    end

    # _running_excel_pids (comms.jl) - used by serve_xl() to auto-detect the Excel process to
    # attach to. Just a smoke test: CI has no Excel installed, so this only checks the Windows
    # `tasklist` call runs and parses to the right type, not that it finds anything.
    if Sys.iswindows()
        @test CallFromExcel._running_excel_pids() isa Vector{Int}
    end

    # serve_xl (comms.jl) - given an explicit pid, throws unless it matches a real running Excel
    # process (Windows only) - CI has no Excel installed, so any pid is rejected there.
    if Sys.iswindows()
        @test_throws String CallFromExcel.serve_xl(true, 999999999)
    end

    # _attach_to_excel (comms.jl) - the actual attach logic serve_xl delegates to once it has a
    # validated pid; tested directly since CI has no real Excel process for serve_xl's own
    # validation to accept. Clears any stale port file left over from a previous local run first -
    # otherwise, if it happens to reference a port still genuinely listening (e.g. an orphaned
    # process from an earlier interrupted test run), the new already-listening check below would
    # spuriously reject this pid. show_results defaults to true (unlike display_results itself,
    # which defaults to false), and is passed through to display_results. Stop the server and reset
    # display_results afterwards so neither leaks into a later test.
    CallFromExcel.setxlpid(4321)
    CallFromExcel.comms_folder(CallFromExcel._default_commsfolder())
    isfile(CallFromExcel.portfile()) && rm(CallFromExcel.portfile())
    CallFromExcel._attach_to_excel(true, 4321)
    @test CallFromExcel.display_results() == true
    @test CallFromExcel.server_status() == (running=true, pid=4321, port=CallFromExcel.xlport[], display_results=true)
    CallFromExcel.stop_server()
    @test CallFromExcel.server_status().running == false
    CallFromExcel._attach_to_excel(false, 4321)
    @test CallFromExcel.display_results() == false

    # start_server dispatches on the request path, with no default: "/eval", "/call" and "/callkw"
    # reach srv_eval_inner and srv_call_inner, and any other path gets a 404.
    let url(path) = "http://127.0.0.1:$(CallFromExcel.xlport[])$path",
        token = CallFromExcel.session_token[],
        post(path, body; headers=["X-CallFromExcel-Token" => token]) =
            CallFromExcel.HTTP.post(url(path), headers, body; status_exception=false)
        # The port file holds the port, the session token, the package version (which the
        # add-in checks against its own), the project folder, "WSL" or "", and Julia's version,
        # for Excel to read.
        @test split(read(CallFromExcel.portfile(), String), "\n") ==
              [string(CallFromExcel.xlport[]), token, string(pkgversion(CallFromExcel)),
               dirname(Base.active_project()),   # Pkg.test's own environment, not the default
               Sys.iswindows() ? "" : "WSL", string(VERSION), CallFromExcel.server_token[]]
        # Replies carry the server token, proving to Excel that they're from this session; a
        # refused request's reply doesn't, so the token can't be fished for.
        @test CallFromExcel.HTTP.header(post("/ping", ""), "X-CallFromExcel-Server") == CallFromExcel.server_token[]
        @test CallFromExcel.HTTP.header(post("/eval", "1+1"), "X-CallFromExcel-Server") == CallFromExcel.server_token[]
        @test CallFromExcel.HTTP.header(post("/ping", ""; headers=["X-CallFromExcel-Token" => "wrong"]), "X-CallFromExcel-Server") == ""
        @test CallFromExcel.project_folder("C:\\p\\Project.toml", "C:\\d\\Project.toml") == "C:\\p"
        @test CallFromExcel.project_folder("C:\\d\\Project.toml", "C:\\d\\Project.toml") == ""
        @test CallFromExcel.project_folder(nothing, "C:\\d\\Project.toml") == ""
        @test CallFromExcel.project_folder("C:\\t\\Project.toml", "C:\\d\\Project.toml", true) == ""   # @temp
        @test length(token) == 32
        @test CallFromExcel.decode_from_xl(String(post("/eval", "1+1").body)) == 2
        # A function defined by one request can be called by the next, on the same kept-alive
        # connection (HTTP.jl 2 serves those from one task, so this needs invokelatest).
        post("/eval", "world_age_test(x) = 2x")
        @test CallFromExcel.decode_from_xl(String(post("/call", CallFromExcel.encode_for_xl(Any["world_age_test", 21])).body)) == 42
        # A function defined by the request itself, as an anonymous function is when its name is
        # evaluated - so the call needs a newer world than the request started in.
        @test CallFromExcel.decode_from_xl(String(post("/call", CallFromExcel.encode_for_xl(Any["(x->2x)", 21])).body)) == 42
        @test vec(CallFromExcel.decode_from_xl(String(post("/call", CallFromExcel.encode_for_xl(Any["(x->2x).", [1, 2, 3]])).body))) == [2, 4, 6]
        @test CallFromExcel.decode_from_xl(String(post("/callkw", CallFromExcel.encode_for_xl(Any["((x; k) -> x + k)", 1, ";", "k", 2])).body)) == 3
        @test CallFromExcel.decode_from_xl(String(post("/call", CallFromExcel.encode_for_xl(Any["+", 1, 2])).body)) == 3
        @test CallFromExcel.decode_from_xl(String(post("/callkw",
            CallFromExcel.encode_for_xl(Any["round", 1.25, ";", "digits", 1])).body)) == 1.2
        @test post("/", "1+1").status == 404
        @test post("/evil", "1+1").status == 404

        # Requests without this session's token are refused with 409, which the add-in treats as
        # "Julia not reachable" (e.g. another Excel, or another local user); a web page (Origin
        # header), a DNS-rebound host name, or a non-POST is refused outright.
        @test post("/eval", "1+1"; headers=[]).status == 409
        @test post("/eval", "1+1"; headers=["X-CallFromExcel-Token" => reverse(token)]).status == 409
        @test post("/ping", ""; headers=[]).status == 409
        @test post("/eval", "1+1"; headers=["X-CallFromExcel-Token" => token, "Origin" => "https://example.com"]).status == 403
        @test post("/eval", "1+1"; headers=["X-CallFromExcel-Token" => token, "Host" => "evil.example.com"]).status == 403
        @test CallFromExcel.HTTP.get(url("/eval"); status_exception=false).status == 405

        # /ping has no side effects.
        post("/eval", "40+2")
        @test post("/ping", "").status == 200
        # start_server's warm-up ping (made before Excel is told the server is ready) is answered,
        # and on a port nothing listens on it fails quietly, as it only saves time.
        @test CallFromExcel.warm_up(CallFromExcel.xlport[]) === nothing
        @test CallFromExcel.warm_up(1) === nothing
        @test CallFromExcel.last_question == "40+2"
    end

    # _attach_to_excel refuses to start a second, useless listener when one is already serving the
    # same pid - starting a fresh session again for 4321 without stopping the previous one first.
    @test_throws String CallFromExcel._attach_to_excel(true, 4321)
    CallFromExcel.stop_server()

    # Another program on the port of a session that has ended (e.g. another user's, on a shared
    # machine) isn't taken for Julia: the add-in's own client, given a port file pointing at
    # such a program - an HTTP server that answers anything with 200 - reports no connection,
    # and the program never receives the real request, only the /ping that checks it first.
    # Given the real session's port file, the same client gets its answer.
    if Sys.iswindows()
        @testset "Only this Excel's Julia is believed" begin
            dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
            CallFromExcel.setxlpid(4321)
            CallFromExcel.comms_folder(CallFromExcel._default_commsfolder())
            CallFromExcel.start_server()
            real = read(CallFromExcel.portfile(), String)
            seen = String[]
            fakeport = CallFromExcel.xlport[] + 50
            fake = CallFromExcel.HTTP.serve!("127.0.0.1", fakeport) do req
                push!(seen, req.target)
                CallFromExcel.HTTP.Response(200, "Not Julia")
            end
            try
                realfile, impostorfile = tempname(), tempname()
                write(realfile, real)
                write(impostorfile, replace(real, r"^\d+" => string(fakeport)))
                oldfile = tempname()   # as written by an older package: no server token
                write(oldfile, join(split(read(impostorfile, String), "\n")[1:6], "\n"))
                ps1 = tempname() * ".ps1"
                write(ps1, """
                    [void][Reflection.Assembly]::LoadFrom('$dll')
                    \$file = Join-Path ([CallFromExcel.AddIn.FileUtils]::LocalTemp()) ("Port_" + \$PID + ".txt")
                    function Ask(\$portfile) {
                        [CallFromExcel.AddIn.JuliaPort]::SetJuliaPort(0)
                        Copy-Item \$portfile \$file
                        try { [CallFromExcel.AddIn.JuliaHttpClient]::Post('1+1', '/eval') }
                        catch { 'ERROR: ' + \$_.Exception.GetBaseException().Message }
                    }
                    function Listening(\$portfile) {
                        [CallFromExcel.AddIn.JuliaPort]::SetJuliaPort(0)
                        Copy-Item \$portfile \$file
                        [CallFromExcel.AddIn.JuliaHttpClient]::IsListening([CallFromExcel.AddIn.JuliaPort]::ReadJuliaPortFromFile())
                    }
                    Ask('$realfile')
                    Ask('$impostorfile')
                    Listening('$impostorfile')
                    Listening('$oldfile')
                    [CallFromExcel.AddIn.JuliaPort]::SetJuliaPort(0)
                    """)
                out = readlines(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ps1`)
                rm(ps1)
                @test length(out) == 4
                @test !startswith(out[1], "ERROR") && occursin("2", out[1])
                @test startswith(out[2], "ERROR: There is no connection between Excel and Julia")
                @test all(==("/ping"), seen)   # never a real request
                # IsListening, the "is my Julia there?" check: not the impostor - unless the port
                # file has no server token, as from a CallFromExcel package older than the add-in,
                # whose session counts as running, so that its version mismatch is what's reported.
                @test out[3] == "False"
                @test out[4] == "True"
            finally
                close(fake)
                CallFromExcel.stop_server()
            end
            # When Julia exits, its port file is deleted (atexit), so Excel finds no session
            # rather than trying the freed port - but only if it's still this session's.
            portfile = CallFromExcel.portfile()
            @test isfile(portfile)   # stop_server leaves it
            CallFromExcel._remove_portfile()
            @test !isfile(portfile)
            write(portfile, "2700\nanothersessionstoken\n")
            CallFromExcel._remove_portfile()
            @test isfile(portfile)
            rm(portfile)
            # And it really is run at exit: in a Julia of its own, which starts a server, then exits.
            dir = mktempdir()
            code = "using CallFromExcel; CallFromExcel.setxlpid(98765); " *
                   "CallFromExcel.comms_folder($(repr(dir))); CallFromExcel.start_server(47950); " *
                   "print(\"written: \", isfile(CallFromExcel.portfile()))"
            output = read(`$(Base.julia_cmd()) --project=$(dirname(Base.active_project())) --startup-file=no -e $code`, String)
            @test endswith(output, "written: true")
            @test !isfile(joinpath(dir, "Port_98765.txt"))
        end
    end

    # When the Excel being served exits, the server stops and its port file goes, so a later
    # Excel given the same process id can't connect to this session. A short-lived Julia stands
    # in for Excel, and the checks are made every 0.2s rather than every 5s. Also: a watcher for
    # an earlier server stands down, rather than stopping a newer one (here serving made-up pid
    # 4321, which isn't watched).
    if Sys.iswindows()
        @testset "Stop serving when Excel closes" begin
            CallFromExcel._watch_interval[] = 0.2
            dummy() = run(`$(Base.julia_cmd()) --startup-file=no -e "sleep(120)"`; wait=false)
            stopped(timeout) = timedwait(() -> !CallFromExcel.server_status().running, timeout) == :ok
            try
                excel = dummy()
                pid = Int64(getpid(excel))   # while it runs: getpid fails once it has exited
                CallFromExcel.setxlpid(pid)
                CallFromExcel.start_server()
                portfile = CallFromExcel.portfile()
                @test isfile(portfile)
                @test !stopped(1.0)   # still serving while "Excel" runs
                kill(excel); wait(excel)
                @test stopped(5.0)
                @test timedwait(() -> !isfile(portfile), 2.0) == :ok
                @test CallFromExcel._excel_closed[]
                # The title no longer says "serving", which the add-in looks for.
                title = mktemp() do path, io
                    redirect_stdout(CallFromExcel.settitle, io)
                    close(io)
                    read(path, String)
                end
                @test occursin("Excel PID $pid has closed", title) && !occursin("serving", title)

                # A port file that is no longer this session's - as if Windows had given the
                # closed Excel's id to a new Excel, which has launched its own Julia - is left.
                excel = dummy()
                CallFromExcel.setxlpid(Int64(getpid(excel)))
                CallFromExcel.start_server()
                portfile = CallFromExcel.portfile()
                write(portfile, "2700\nanothersessionstoken\n")
                kill(excel); wait(excel)
                @test stopped(5.0)
                @test isfile(portfile)
                rm(portfile)

                excel = dummy()
                CallFromExcel.setxlpid(Int64(getpid(excel)))
                CallFromExcel.start_server()
                CallFromExcel.setxlpid(4321)
                CallFromExcel.start_server()
                kill(excel); wait(excel)
                @test !stopped(1.5)
            finally
                CallFromExcel._watch_interval[] = 5.0
                CallFromExcel.stop_server()
                CallFromExcel.setxlpid(4321)
            end
        end
    end

    # setvar assigns the value itself. It used to re-parse the value's printed form as code, which
    # failed for strings ("hello" became the variable name hello) and for ExcelError values.
    for v in Any["hello", "a\"b\$c", 1.5, [1.0, 2.0], Any[1.0, missing, CallFromExcel.ExcelError(2042)],
                 CallFromExcel.ExcelError(2007), Dict("a" => 1)]
        CallFromExcel.setvar("setvar_test_value", v)
        got = invokelatest(getglobal, Main, :setvar_test_value)
        @test isequal(got, v) && typeof(got) == typeof(v)
    end
    @test CallFromExcel.setvar("setvar_test_value", [1.0, 2.0]) == "Set global variable `setvar_test_value` to 2-element Vector{Float64}"
    @test startswith(CallFromExcel.setvar("1bad", 1), "#")

end

# The worksheet functions' descriptions (ExcelFunction/ExcelArgument attributes) are the only
# source of their IntelliSense and Function Wizard help. Excel-DNA cuts any longer than Excel's
# limit when registering them - 253 characters for a function's and its last argument's
# (it adds ". "), 255 for the others - with only a log warning, and IntelliSense shows the cut
# text too. So all are kept within 253.
@testset "Function and argument descriptions fit Excel's limit" begin
    for file in ("JuliaFunctions.cs", "JuliaCallDiag.cs")
        src = read(joinpath(@__DIR__, "..", "dotnet-addin", "CallFromExcel.AddIn", "src", file), String)
        for m in eachmatch(r"Description = \"((?:[^\"\\]|\\.)*)\"", src)
            text = replace(m.captures[1], r"\\(.)" => s"\1")
            length(text) <= 253 || @error "Description too long ($(length(text)) characters) in $file" text
            @test length(text) <= 253
        end
    end
end

# `using CallFromExcel` in the REPL reminds the user to run installme() when the version file
# Install.ps1 writes is missing or names a different version.
# installme() refuses when the session's CallFromExcel isn't the one the environment now has
# (updated after it was loaded), as it would install the loaded version's add-in.
@testset "installme() spots a package updated after it was loaded" begin
    old = joinpath("C:", "depot", "packages", "CallFromExcel", "Mm96J", "src", "CallFromExcel.jl")
    new = joinpath("C:", "depot", "packages", "CallFromExcel", "PHo8V", "src", "CallFromExcel.jl")
    msg = CallFromExcel.stale_session_message(old, new)
    @test occursin("Mm96J", msg) && occursin("PHo8V", msg) && occursin("restart Julia", msg)
    @test CallFromExcel.stale_session_message(new, new) === nothing
    @test CallFromExcel.stale_session_message(old, nothing) === nothing
    # As this test session is: the loaded package is the one its environment has.
    @test CallFromExcel.stale_session_message(pathof(CallFromExcel),
              Base.locate_package(Base.PkgId(CallFromExcel))) === nothing
end

# For the object handles tests: a type with a custom label, as a user's own package might define.
struct HandleTestCurve
    currency::String
end
CallFromExcel.handle_label(x::HandleTestCurve) = "Curve_$(x.currency)"
struct HandleTestLabelled
    label::String
end
CallFromExcel.handle_label(x::HandleTestLabelled) = x.label

@testset "Object handles in Julia" begin
    C = CallFromExcel
    # Mirrors the add-in's HandlePattern.cs; the key is "<id>@<session>".
    pattern = r"^jl:[^#@\s\p{Z}\p{Cc}\p{Cf}]*#([0-9]+@[A-Za-z0-9]+)$"
    key(h) = String(match(pattern, h)[1])

    h = C.handle_eval("[1.0 2.0; 3.0 4.0]")
    @test startswith(h, "jl:Array#") && occursin(pattern, h)
    @test occursin(r"@[A-Za-z0-9]{4}$", h)                   # the session tag: 4 letters or digits
    @test C.decode_from_xl("O" * key(h)) == [1.0 2.0; 3.0 4.0]
    @test startswith(C.handle_eval("Dict(\"a\" => 1)"), "jl:Dict#")

    # Custom labels, and the characters a label can't contain.
    h = C.handle_create("HandleTestCurve", "USD")
    @test startswith(h, "jl:Curve_USD#") && occursin(pattern, h)
    @test C.decode_from_xl("O" * key(h)).currency == "USD"
    h = C.handle_create("HandleTestLabelled", "EUR/USD @ 1#2\tΣ​")
    @test startswith(h, "jl:EUR/USD___1_2_Σ_#") && occursin(pattern, h)
    @test startswith(C.handle_create("HandleTestLabelled", ""), "jl:#")
    @test occursin(pattern, C.handle_eval("x -> 2x"))              # a closure's type name has '#'

    # Freed, or from another session: not found.
    @test C.handle_free(key(h)) isa Int
    @test_throws ErrorException C.decode_from_xl("O" * key(h))
    @test_throws ErrorException C.decode_from_xl("O1@zzzz")
end

@testset "expose_to_excel checks Excel names" begin
    C = CallFromExcel
    Core.eval(Main, :(expose_test_f(a, b) = a + b; Expose_Test_F(x) = x; expose_test!(v) = v))
    msg(f) = try f(); "" catch e; e.msg end
    empty!(C._exposed)
    @test C.exposed_functions() === nothing                        # Excel has no empty arrays

    # The version the add-in compares: this session's tag, and a count every expose_to_excel bumps.
    v0 = C.exposed_version()
    @test startswith(v0, C.handle_session() * "#")
    @test C.expose_to_excel("expose_test_f") == "jl.expose_test_f"
    @test C.exposed_version() != v0
    v1 = C.exposed_version()
    @test (try C.expose_to_excel("expose_test!") catch end; C.exposed_version() == v1)   # a refusal doesn't
    @test C.exposed_functions()[1, 4] == "a\x1fb"
    @test C.expose_to_excel("expose_test_f") == "jl.expose_test_f"           # exposing again replaces
    @test C.expose_to_excel("expose_test_f"; name="PRICE_SWAP") == "PRICE_SWAP"
    @test C.expose_to_excel("expose_test_f"; name="jl.σ") == "jl.σ"
    @test C.expose_to_excel("expose_test_f"; name="abc.123") == "abc.123"
    @test C.expose_to_excel("expose_test_f"; name="\\abc") == "\\abc"

    # Names Excel can't register.
    @test occursin("isn't a valid Excel function name", msg(() -> C.expose_to_excel("expose_test!")))
    @test occursin("with `name=`", msg(() -> C.expose_to_excel("expose_test!")))
    @test C.expose_to_excel("expose_test!"; name="jl.expose_test") == "jl.expose_test"
    @test occursin("isn't a valid", msg(() -> C.expose_to_excel("expose_test_f"; name="1abc")))
    @test occursin("isn't a valid", msg(() -> C.expose_to_excel("expose_test_f"; name="a b")))
    @test occursin("too long", msg(() -> C.expose_to_excel("expose_test_f"; name="a"^256)))
    for ref in ["abc123", "A1", "xfd1048576", "R", "c", "R1", "C22", "RC", "R1C1", "r2c", "R1C1.x", "A1.b.c"]
        @test occursin("cell reference", msg(() -> C.expose_to_excel("expose_test_f"; name=ref)))
    end
    for reserved in ["JuliaCall", "juliaprice", "JULIA"]
        @test occursin("reserved", msg(() -> C.expose_to_excel("expose_test_f"; name=reserved)))
    end

    # Excel ignores case: jl.Expose_Test_F would replace jl.expose_test_f.
    @test occursin("ignores case", msg(() -> C.expose_to_excel("Expose_Test_F")))
    @test occursin("ignores case", msg(() -> C.expose_to_excel("expose_test_f"; name="price_swap")))
    @test C.expose_to_excel("Expose_Test_F"; name="jl.Expose_Test_F2") == "jl.Expose_Test_F2"

    # A rejected name records nothing.
    @test sort(C.exposed_functions()[:, 1]) ==
        sort(["jl.expose_test_f", "PRICE_SWAP", "jl.σ", "abc.123", "\\abc", "jl.expose_test", "jl.Expose_Test_F2"])
    empty!(C._exposed)
end

@testset "expose_to_excel reads docstrings and optional arguments" begin
    C = CallFromExcel
    include_string(Main, """
        \"\"\"
            expose_doc_swap(notional, fixed_rate, years=10, spread=0.0)

        Price a toy swap, with **notional** in [USD](https://example.com)
        and `years` whole.

        More detail, not shown in Excel.

        # Arguments
        - `notional`: the notional amount.
        - `fixed_rate::Float64`: the fixed rate,
          as a decimal.
        - `years=10` - the term.

        # Examples
        - `notanarg`: not an argument.
        \"\"\"
        expose_doc_swap(notional, fixed_rate, years=10, spread=0.0) = notional * (fixed_rate + spread) * years

        \"\"\"
        ```julia
        expose_doc_fenced(x)
        ```
        Doubles x.
        \"\"\"
        expose_doc_fenced(x) = 2x

        expose_doc_none(a, b) = a
        expose_doc_methods(a) = 1
        expose_doc_methods(a, b, c) = 3
        expose_doc_methods(a, b) = 2
        expose_doc_methods(a, b, c, d...) = 4
        """)
    rows() = Dict(r[1] => r for r in eachrow(C.exposed_functions()))
    empty!(C._exposed)

    C.expose_to_excel("expose_doc_swap")
    r = rows()["jl.expose_doc_swap"]
    @test r[3] == "Price a toy swap, with notional in USD and years whole."
    @test r[4] == join(["notional", "fixed_rate", "years", "spread"], "\x1f")
    @test split(r[5], "\x1f") == ["the notional amount.", "the fixed rate, as a decimal.",
                                  "Optional. The term.", "Optional."]
    @test r[6] == 2

    # Given descriptions win, position by position where not empty.
    C.expose_to_excel("expose_doc_swap"; description="Mine", argdescriptions=["", "Rate"])
    r = rows()["jl.expose_doc_swap"]
    @test r[3] == "Mine"
    @test split(r[5], "\x1f")[1:3] == ["the notional amount.", "Rate", "Optional. The term."]
    C.expose_to_excel("expose_doc_swap"; argdescriptions=["", "", "Optional, default 10"])
    @test split(rows()["jl.expose_doc_swap"][5], "\x1f")[3] == "Optional, default 10"

    C.expose_to_excel("expose_doc_fenced")
    @test rows()["jl.expose_doc_fenced"][3] == "Doubles x."
    C.expose_to_excel("expose_doc_none")
    r = rows()["jl.expose_doc_none"]
    @test r[3] == "" && r[5] == "\x1f" && r[6] == 2

    # Several methods: the non-varargs one with most arguments, all required (separate lines).
    C.expose_to_excel("expose_doc_methods")
    r = rows()["jl.expose_doc_methods"]
    @test r[4] == join(["a", "b", "c"], "\x1f") && r[6] == 3

    # Long descriptions are cut to Excel's limit.
    @test length(C._fit("x"^300)) == 253 && endswith(C._fit("x"^300), "...")
    @test C._fit("short") == "short"
    empty!(C._exposed)
end

@testset "@excel, and broadcasting in named functions" begin
    C = CallFromExcel
    empty!(C._exposed)
    call(name, args...) = C._call_exposed(name, args...)
    msg(f) = try f(); "" catch e; e.msg end
    include_string(Main, """
        \"\"\"
            excel_test_scale(x, factor)

        Scales x.
        \"\"\"
        @excel broadcast=true excel_test_scale(x, factor) = x * factor

        excel_test_m(a) = 1
        @excel name="jl.excel_test_m1" excel_test_m(a::Int) = 2
        excel_test_m(a, b, c) = 3

        @excel broadcast=true function excel_test_wavg(x, w::AbstractVector)
            x * sum(w) / length(w)
        end
        @excel broadcast=(:b,) excel_test_list(a, b) = (a, b)
        @excel excel_test_size(a) = size(a)
        @excel broadcast=:x excel_test_pair(x) = [x, x]
        excel_test_plain(a, b) = a * length(b)
        """)

    # @excel defines, exposes and documents.
    r = Dict(r[1] => r for r in eachrow(C.exposed_functions()))
    @test r["jl.excel_test_scale"][3] == "Scales x."
    @test Main.excel_test_scale(2, 3) == 6
    # The method @excel is attached to, not the one with most arguments.
    @test r["jl.excel_test_m1"][4] == "a"
    @test C._exposed["jl.excel_test_m1"].f === Main.excel_test_m

    # Scalars: a plain call. A range broadcasts; a column against a row gives a table.
    @test call("jl.excel_test_scale", 2.0, 3.0) == 6.0
    @test call("jl.excel_test_scale", [1.0; 2.0;;], 10.0) == [10.0; 20.0;;]
    @test call("jl.excel_test_scale", [1.0; 2.0;;], [1.0 10.0 100.0]) == [1.0 10.0 100.0; 2.0 20.0 200.0]
    @test occursin("x is 3x1 and factor is 2x1", msg(() -> call("jl.excel_test_scale", [1.0; 2; 3;;], [1.0; 2;;])))

    # Declared ::AbstractVector: passed whole, as a vector, from a cell, row or column.
    @test call("jl.excel_test_wavg", [1.0; 2.0;;], [1.0 3.0]) == [2.0; 4.0;;]
    @test call("jl.excel_test_wavg", 1.0, 5.0) == 5.0
    @test call("jl.excel_test_wavg", 1.0, [2.0; 4.0;;]) == 3.0

    # broadcast= names the arguments that broadcast; by default (false), none do.
    @test call("jl.excel_test_list", [1 2], [3 4]) == [([1 2], 3) ([1 2], 4)]
    @test call("jl.excel_test_size", [1 2 3]) == (1, 3)
    @test occursin("isn't an argument", msg(() -> C.expose_to_excel("excel_test_plain"; broadcast=(:z,))))
    @test C.expose_to_excel("excel_test_plain"; broadcast=(:a,)) == "jl.excel_test_plain"
    @test call("jl.excel_test_plain", [5 6], [1 2 3]) == [15 18]
    @test C.expose_to_excel(Main.excel_test_plain; name="jl.plain2") == "jl.plain2"     # a function, not a name
    @test call("jl.plain2", [5 6], [1 2]) == [10 12]                      # by default, nothing broadcasts

    # An array for each element can't spill; an unknown name (a relaunched Julia).
    @test occursin("returns an array for each element", msg(() -> call("jl.excel_test_pair", [1 2])))
    @test call("jl.excel_test_pair", 1) == [1, 1]
    @test occursin("isn't exposed in this Julia session", msg(() -> call("jl.nonesuch", 1)))

    # Through the server (as =jl.f(...) arrives), the call is shown, saved and repeatable in the
    # REPL as a call of the function itself, with args_from_xl as the function receives them.
    srv(args...) = C.decode_from_xl(C.srv_call_inner(C.encode_for_xl(Any["CallFromExcel._call_exposed", args...])))
    @test srv("jl.excel_test_scale", 2.0, 3.0) == 6.0
    @test C.last_question == "excel_test_scale(args_from_xl...)"
    @test C.args_from_xl == [2.0, 3.0] && C.answer_again() == 6.0
    @test vec(srv("jl.excel_test_scale", [1.0; 2.0;;], 10.0)) == [10.0, 20.0]
    @test C.last_question == "excel_test_scale.(args_from_xl...)"
    @test C.args_from_xl[2] isa Ref && vec(C.answer_again()) == [10.0, 20.0]
    @test srv("jl.excel_test_wavg", 1.0, [2.0; 4.0;;]) == 3.0
    @test C.args_from_xl[2] == [2.0, 4.0] && C.answer_again() == 3.0         # the vector w receives
    @test occursin("x is 3x1 and factor is 2x1", srv("jl.excel_test_scale", [1.0; 2; 3;;], [1.0; 2;;]))
    @test C.last_question == "excel_test_scale.(args_from_xl...)"
    @test occursin("returns an array for each element", srv("jl.excel_test_pair", [1.0 2.0]))
    @test occursin("isn't exposed in this Julia session", srv("jl.nonesuch", 1.0))

    # Code from JuliaEval is all at none:1, as Meta.parseall gives here. Which method @excel
    # defined, and which arguments are optional, can't be judged by file and line.
    jleval(code) = Main.eval(Meta.parseall(code))
    jleval("excel_test_g(a) = a")
    jleval("@excel excel_test_g(a, b) = a + b")
    @test C._exposed["jl.excel_test_g"].nrequired == 2                  # b isn't optional
    jleval("excel_test_h(x::Int) = 1")
    jleval("@excel excel_test_h(x::AbstractVector) = sum(x)")
    @test C._exposed["jl.excel_test_h"].asvector == [true]              # the method just defined
    @test call("jl.excel_test_h", [1.0 2.0 3.0]) == 6.0
    jleval("@excel excel_test_d(a, b=2, c=length(a)) = a")
    @test C._exposed["jl.excel_test_d"].nrequired == 1
    jleval("@excel excel_test_d(a, b=2, c=length(a)) = a")              # defined again, as re-including does
    @test C._exposed["jl.excel_test_d"].nrequired == 1
    jleval("excel_test_e(a) = excel_test_e(a, 5); excel_test_e(a, b) = a + b")
    C.expose_to_excel("excel_test_e")                                   # a hand-written method isn't a default
    @test C._exposed["jl.excel_test_e"].nrequired == 2

    # An unnamed argument gets a name; more arguments from Excel than Julia now takes (an
    # older registration) is an error saying what to do.
    jleval("@excel excel_test_u(::Int, y) = y")
    @test C._exposed["jl.excel_test_u"].argnames == ["arg1", "y"]
    @test occursin("now takes 2 arguments in Julia, but Excel passed 3", msg(() -> call("jl.excel_test_g", 1, 2, 3)))

    # Bad uses of the macro.
    @test_throws Exception include_string(Main, "@excel nonsense=1 excel_test_bad(x) = x")
    @test_throws Exception include_string(Main, "@excel 42")
    empty!(C._exposed)
end

@testset "Reminder to run installme()" begin
    mktempdir() do dir
        versionfile = joinpath(dir, "version.txt")
        @test occursin("is not installed", CallFromExcel.addin_install_reminder(versionfile, v"0.1.1"))
        write(versionfile, "0.1.0")
        msg = CallFromExcel.addin_install_reminder(versionfile, v"0.1.1")
        @test occursin("version 0.1.0", msg) && occursin("version 0.1.1", msg) && occursin("installme()", msg)
        # An add-in newer than the package: update the package, don't downgrade the add-in.
        msg = CallFromExcel.addin_install_reminder(versionfile, v"0.0.9")
        @test occursin("Pkg.update(\"CallFromExcel\")", msg) && occursin("restart Julia", msg)
        @test CallFromExcel.addin_install_reminder(versionfile, v"0.1.0") === nothing
        write(versionfile, "0.1.0\r\n")   # tolerate a trailing newline
        @test CallFromExcel.addin_install_reminder(versionfile, v"0.1.0") === nothing
    end
end

# The add-in finds julia.exe by searching PATH. Windows accepts PATH entries in double quotes, but
# .NET Framework's Path.Combine throws on them, which used to make every JuliaLaunch fail with
# "Illegal characters in path." whenever such an entry came before Julia's folder. Calls the
# built add-in's search directly (Windows PowerShell runs the same .NET Framework as Excel-DNA),
# with PATH set to a quoted missing folder, an invalid entry, then Julia's own folder, quoted.
if Sys.iswindows()
    @testset "The add-in finds julia.exe despite quoted or invalid PATH entries" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        path = "\"C:\\No Such Folder\";C:\\bad<chars>;\"$(Sys.BINDIR)\""
        script = "\$env:PATH = '$path'; [void][Reflection.Assembly]::LoadFrom('$dll'); " *
                 "[CallFromExcel.AddIn.JuliaExeLocation]::Find(\$false)"
        found = readchomp(`powershell -NoProfile -NonInteractive -Command $script`)
        @test lowercase(found) == lowercase(joinpath(Sys.BINDIR, "julia.exe"))
    end

    # JuliaLaunch rejects -L/--load and adds -i and --threads=auto,1 unless given. It used to test
    # for substrings of the whole CommandLineOptions string, so --project=C:/Users/Jean-Luc was
    # rejected as if it were -L, and --project=C:/multi-index suppressed -i.
    @testset "JuliaLaunch's CommandLineOptions checks look at options, not substrings" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        cases = [
            "" => "-i --threads=auto,1",
            "--project=C:/Users/Jean-Luc/multi-index" => "--project=C:/Users/Jean-Luc/multi-index -i --threads=auto,1",
            "--project=\"C:/My -L -i Folder\"" => "--project=\"C:/My -L -i Folder\" -i --threads=auto,1",
            "-q -i" => "-q -i --threads=auto,1",
            "-qi" => "-qi --threads=auto,1",
            "-t 4" => "-t 4 -i",
            "-t4" => "-t4 -i",
            "--threads=8 -i" => "--threads=8 -i",
            "-e \"x = 1 -i\"" => "-e \"x = 1 -i\" -i --threads=auto,1",
            "-L foo.jl" => "THROWS",
            "-Lfoo.jl" => "THROWS",
            "-qL foo.jl" => "THROWS",
            "--load foo.jl" => "THROWS",
            "--load=foo.jl" => "THROWS",
        ]
        ps1 = tempname() * ".ps1"
        write(ps1, "[void][Reflection.Assembly]::LoadFrom('$dll')\n" * join(
            ["try { [CallFromExcel.AddIn.JuliaLaunch]::NormaliseCommandLineOptions('$(c.first)') } catch { 'THROWS' }"
             for c in cases], "\n"))
        results = readlines(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ps1`)
        rm(ps1)
        @test results == last.(cases)
    end

    # The ribbon's Launch Settings turn a project folder into a --project option: quoted, with
    # forward slashes, on Windows; a single-quoted Linux path under WSL.
    @testset "The ribbon's project folder becomes a --project option" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        cases = [
            ("", false) => "",
            ("  ", true) => "",
            ("C:\\Projects\\My Project\\", false) => "\"--project=C:/Projects/My Project/\"",
            (raw"C:\Projects\Foo", true) => "--project='/mnt/c/Projects/Foo'",
            (raw"\\wsl.localhost\Ubuntu\home\me\proj", true) => "--project='/home/me/proj'",
            (raw"\\wsl$\Ubuntu\home\me", true) => "--project='/home/me'",
            ("/home/me/it's", true) => "--project='/home/me/it'\\''s'",
            ("@myenv", true) => "--project='@myenv'",
        ]
        ps1 = tempname() * ".ps1"
        quoteps(s) = "'" * replace(s, "'" => "''") * "'"
        write(ps1, "[void][Reflection.Assembly]::LoadFrom('$dll')\n" * join(
            ["'[' + [CallFromExcel.AddIn.JuliaLaunch]::ProjectOption($(quoteps(c.first[1])), \$$(c.first[2])) + ']'"
             for c in cases], "\n"))
        results = readlines(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ps1`)
        rm(ps1)
        @test results == ["[" * c.second * "]" for c in cases]
    end

    # JuliaLaunch writes the comms folder and flag/error file paths into a Julia startup script.
    # They're under the temp folder, so a '$' in it (e.g. from the user name) used to be read as
    # string interpolation, breaking the script. Generates the script with such paths and checks
    # each comes through as a plain string.
    @testset "JuliaLaunch's startup script survives a \$ in the temp folder's path" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        folder = "C:/Users/a\$b/AppData/Local/Temp/@CallFromExcel"
        errorfile, flagfile = folder * "/LoadError_1234.txt", folder * "/Flag_1234.txt"
        script = "[void][Reflection.Assembly]::LoadFrom('$dll'); " *
                 "[CallFromExcel.AddIn.JuliaLaunch]::BuildStartupScript('', 1234, '$folder', '$errorfile', '$flagfile', 'cmd', 2712)"
        startup = read(`powershell -NoProfile -NonInteractive -Command $script`, String)
        strings(x) = x isa String ? [x] : x isa Expr ? reduce(vcat, strings.(x.args); init=String[]) : String[]
        found = strings(Meta.parseall(startup))
        @test folder in found
        @test errorfile in found
        @test flagfile in found
        @test occursin("CallFromExcel.start_server(2712)", startup)
    end

    # JuliaLaunch picks the port, as the first one free on Windows, because Julia under WSL can't
    # see Windows' ports. Holds a port, and checks that the next one is chosen.
    @testset "JuliaLaunch picks a port that's free on Windows" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        free(start) = parse(Int, read(`powershell -NoProfile -NonInteractive -Command
            "[void][Reflection.Assembly]::LoadFrom('$dll'); [CallFromExcel.AddIn.JuliaLaunch]::FirstFreePort($start)"`, String))
        start = free(47800)
        server = CallFromExcel.Sockets.listen(CallFromExcel.Sockets.IPv4("127.0.0.1"), start)
        try
            @test free(start) > start
        finally
            close(server)
        end
    end

    # The startup script checks the package version before calling anything an older package may
    # lack, and reports a mismatch through the error file. Runs a generated script, with the
    # add-in's version in it replaced by another, in a new Julia using this package.
    @testset "JuliaLaunch's startup script checks the package version first" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        dir = mktempdir()
        errorfile, flagfile = joinpath(dir, "LoadError.txt"), joinpath(dir, "Flag.txt")
        touch(flagfile)
        jl(path) = replace(path, "\\" => "/")
        startupscript(packages) = read(`powershell -NoProfile -NonInteractive -Command
            "[void][Reflection.Assembly]::LoadFrom('$dll'); [CallFromExcel.AddIn.JuliaLaunch]::BuildStartupScript('$packages', 1234, '$(jl(dir))', '$(jl(errorfile))', '$(jl(flagfile))', 'cmd', 2700)"`, String)
        startup = startupscript("")
        # Empty entries in Packages (a trailing or doubled comma) are skipped: a bare `using`
        # would make the whole script unparseable.
        @test !occursin(r"using *\n", startupscript("Example,, "))
        # A Packages entry may be any `using` target: only its package name is checked as
        # installed (and installed into @temp), and the entry itself is what's loaded.
        withtargets = startupscript("Plots.PlotMeasures, LinearAlgebra: norm")
        @test occursin("String[\"Plots\", \"LinearAlgebra\"]", withtargets)
        @test occursin("using LinearAlgebra: norm", withtargets)
        # JULIA_PROJECT=@temp, without --project, is a temporary environment too, as for Julia.
        # In a new Julia, since this one was started with --project, which takes precedence.
        @test readchomp(addenv(`$(Base.julia_cmd()) --startup-file=no -e "using CallFromExcel; print(CallFromExcel.temporary_project())"`,
                               "JULIA_PROJECT" => "@temp", "JULIA_LOAD_PATH" => "$(dirname(Base.active_project()));@stdlib")) == "true"
        addinversion = "\"$(pkgversion(CallFromExcel))\""
        @test occursin(addinversion, startup)
        scriptfile = joinpath(dir, "StartUp.jl")
        write(scriptfile, replace(startup, addinversion => "\"0.0.1\""))
        run(`$(Base.julia_cmd()) --project=$(dirname(Base.active_project())) --startup-file=no $scriptfile`)
        @test readlines(errorfile) == ["CallFromExcelVersionMismatch", string(pkgversion(CallFromExcel)),
                                       dirname(Base.active_project())]
        @test !isfile(flagfile)

        # Without the package: found before any `using`, and reported by a marker and the
        # project folder, for the add-in's own message (JuliaPort.NotInstalledError). An empty
        # project, and a load path without the default environment, which may have the package.
        rm(errorfile); touch(flagfile)
        empty = mktempdir()
        write(scriptfile, startup)
        run(addenv(`$(Base.julia_cmd()) --project=$empty --startup-file=no $scriptfile`, "JULIA_LOAD_PATH" => "@;@stdlib");
            wait=true)
        @test split(read(errorfile, String), "\n") == ["CallFromExcelNotInstalled", empty]
        @test !isfile(flagfile)

        # A project whose Manifest names a package that isn't installed for this Julia (as when
        # a project used from Julia on Windows is first used from Julia under WSL): the script
        # says so and runs Pkg.instantiate. Here the package doesn't exist, so instantiate fails
        # - which shows it was tried, without needing the network to install anything.
        rm(errorfile); touch(flagfile)
        uninstantiated = mktempdir()
        uuid = "5b3c9a5e-1d2f-4e6a-9b7c-0d1e2f3a4b5c"
        write(joinpath(uninstantiated, "Project.toml"), "[deps]\nNoSuchPackageXyz = \"$uuid\"\n")
        write(joinpath(uninstantiated, "Manifest.toml"), """
            julia_version = "$VERSION"
            manifest_format = "2.0"

            [[deps.NoSuchPackageXyz]]
            git-tree-sha1 = "0123456789abcdef0123456789abcdef01234567"
            uuid = "$uuid"
            version = "1.0.0"
            """)
        output = read(addenv(`$(Base.julia_cmd()) --project=$uninstantiated --startup-file=no $scriptfile`,
                             "JULIA_LOAD_PATH" => "@;$(dirname(Base.active_project()));@stdlib"), String)
        @test occursin("which aren't installed for this Julia: NoSuchPackageXyz", output)
        @test occursin("NoSuchPackageXyz", read(errorfile, String))
        @test !isfile(flagfile)

        # A Julia older than CallFromExcel supports: checked before anything else. Simulated by
        # raising the minimum in the script. The minimum must be Project.toml's [compat] julia.
        compat = Base.parsed_toml(joinpath(@__DIR__, "..", "Project.toml"))["compat"]["julia"]
        @test occursin("if VERSION < v\"$compat\"", startup)
        rm(errorfile); touch(flagfile)
        write(scriptfile, replace(startup, "if VERSION < v\"$compat\"" => "if VERSION < v\"99.0\""))
        run(`$(Base.julia_cmd()) --project=$(dirname(Base.active_project())) --startup-file=no $scriptfile`)
        @test startswith(read(errorfile, String), "CallFromExcel needs Julia $compat or later, but this is Julia $VERSION.")
        @test !isfile(flagfile)
        write(scriptfile, startup)

        # A --project folder that doesn't exist: Julia would accept it as a new, empty project
        # and still find CallFromExcel in the default environment, so the script checks first.
        rm(errorfile); touch(flagfile)
        missingfolder = joinpath(dir, "no such folder")
        run(`$(Base.julia_cmd()) --project=$missingfolder --startup-file=no $scriptfile`)
        missingmessage = read(errorfile, String)
        @test startswith(missingmessage, "The project folder $missingfolder doesn't exist.")
        @test !isfile(flagfile)

        # --project=@temp, a temporary environment: refused by Julia 1.10, which would read it as
        # a permanent shared environment named "temp". From 1.11, the folder exists, so the script
        # gets to the version check (made to fail, as above), which names the default environment
        # (""), not the temporary folder. The load path finds CallFromExcel in this test's
        # environment.
        rm(errorfile); touch(flagfile)
        write(scriptfile, replace(startup, addinversion => "\"0.0.1\""))
        run(addenv(`$(Base.julia_cmd()) --project=@temp --startup-file=no $scriptfile`,
                   "JULIA_LOAD_PATH" => "@;$(dirname(Base.active_project()));@stdlib"))
        if VERSION < v"1.11"
            @test startswith(read(errorfile, String), "--project=@temp, a temporary environment, needs Julia 1.11 or later")
        else
            @test split(read(errorfile, String), "\n") == ["CallFromExcelVersionMismatch", string(pkgversion(CallFromExcel)), ""]
        end
        @test !isfile(flagfile)

        # One of the user's Packages that isn't installed: a message naming it and its
        # environment, starting as the add-in expects, with the Pkg.add command it can copy.
        rm(errorfile); touch(flagfile)
        write(scriptfile, startupscript("NoSuchPackageXyz"))
        run(`$(Base.julia_cmd()) --project=$(dirname(Base.active_project())) --startup-file=no $scriptfile`)
        nosuch = read(errorfile, String)
        @test startswith(nosuch, "The NoSuchPackageXyz package is not installed in Julia's active environment")
        @test occursin("import Pkg; Pkg.add(\"NoSuchPackageXyz\")", nosuch)
        @test !isfile(flagfile)

        # With --project=@temp (Julia 1.11 on), a missing one is installed into the temporary
        # environment instead: Example (the Julia org's tiny example package), which isn't on
        # this load path. The script then gets as far as the version check (made to fail, as
        # above), so it installed and loaded Example. Needs the network, so skipped without it
        # (checked with an HTTPS connection to github.com, which networks that block ping allow).
        online = try close(CallFromExcel.Sockets.connect("github.com", 443)); true catch; false end
        online || @info "No network: skipping the test of installing Packages into @temp"
        if VERSION >= v"1.11" && online
            rm(errorfile); touch(flagfile)
            write(scriptfile, replace(startupscript("Example"), addinversion => "\"0.0.1\""))
            run(addenv(`$(Base.julia_cmd()) --project=@temp --startup-file=no $scriptfile`,
                       "JULIA_LOAD_PATH" => "@;$(dirname(Base.active_project()));@stdlib"))
            @test split(read(errorfile, String), "\n") == ["CallFromExcelVersionMismatch", string(pkgversion(CallFromExcel)), ""]
            @test !isfile(flagfile)
        end
    end

    # The four version-mismatch messages: the package or the add-in older, natively or under
    # WSL, where installme() (Windows only) can't be run in the same Julia. Each comes with the
    # Julia command the ribbon's error dialog offers to copy.
    @testset "Version-mismatch messages" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        err(v, folder, wsl) = "\$e = [CallFromExcel.AddIn.JuliaPort]::MismatchError('$v', '$folder', \$$wsl); \$e.Message; \$e.JuliaCommand"
        script = "[void][Reflection.Assembly]::LoadFrom('$dll')\n" * join(
            [err("0.0.1", "", false), err("0.0.1", "C:\\p", false), err("9.0.0", "", false),
             err("0.0.1", "/home/p", true), err("9.0.0", "", true)], "\n")
        ps1 = tempname() * ".ps1"
        write(ps1, script)
        out = readlines(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ps1`)
        rm(ps1)
        m, cmd = out[1:2:end], out[2:2:end]   # each message, then its command
        update = "using Pkg; Pkg.update(\"CallFromExcel\")"
        install = "using CallFromExcel; CallFromExcel.installme()"
        @test occursin("then start Julia and run: $update; $install", m[1]) && cmd[1] == "$update; $install"
        @test occursin("start Julia with julia --project=\"C:\\p\" and run: using Pkg", m[2]) && cmd[2] == "$update; $install"
        @test occursin("Close Excel, then start Julia and run: $install", m[3]) && cmd[3] == install
        # Under WSL, whichever is older, the package is changed to the add-in's version.
        addmatching = "import Pkg; Pkg.add(name=\"CallFromExcel\", version=\"$(pkgversion(CallFromExcel))\")"
        @test occursin("start Julia by typing julia --project=\"/home/p\", and run: $addmatching - then", m[4]) && cmd[4] == addmatching
        @test occursin("start Julia by typing julia, and run: $addmatching - then", m[5]) && cmd[5] == addmatching
    end

    # The messages for a launch that fails for want of something: WSL not set up, Julia not
    # installed under WSL, or the CallFromExcel package not installed - each with the command
    # the ribbon's error dialog offers to copy.
    @testset "Messages for what a launch needs" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        show(e) = "\$e = $e; \$e.Message; \$e.JuliaCommand"
        script = "[void][Reflection.Assembly]::LoadFrom('$dll')\n" * join([
            show("[CallFromExcel.AddIn.JuliaLaunch]::WslProblemError(\$false, \$false)"),
            show("[CallFromExcel.AddIn.JuliaLaunch]::WslProblemError(\$true, \$false)"),
            show("[CallFromExcel.AddIn.JuliaPort]::NotInstalledError('', \$true)"),
            show("[CallFromExcel.AddIn.JuliaPort]::NotInstalledError('C:\\p', \$false)"),
            "[CallFromExcel.AddIn.JuliaLaunch]::WslProblemError(\$true, \$true) -eq \$null",
            show("[CallFromExcel.AddIn.JuliaExeLocation]::NotFoundError()")], "\n")
        ps1 = tempname() * ".ps1"
        write(ps1, script)
        out = readlines(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ps1`)
        rm(ps1)
        addmatching = "import Pkg; Pkg.add(name=\"CallFromExcel\", version=\"$(pkgversion(CallFromExcel))\")"
        @test occursin("WSL (Windows Subsystem for Linux) isn't set up", out[1]) && occursin("run: wsl --install - then", out[1])
        @test out[2] == "wsl --install"
        @test occursin("Julia isn't installed in it", out[3]) && out[4] == "curl -fsSL https://install.julialang.org | sh"
        @test occursin("Julia under WSL doesn't have the CallFromExcel package", out[5]) &&
              occursin("start Julia by typing julia, and run: $addmatching", out[5]) && out[6] == addmatching
        @test occursin("start Julia with julia --project=\"C:\\p\", and run: $addmatching", out[7]) && out[8] == addmatching
        @test out[9] == "True"
        winget = "winget install --name Julia --id 9NJNWW8PVKMN -e -s msstore"
        @test occursin("Julia isn't installed on this PC, or Excel can't find it", out[10]) &&
              occursin("run: $winget - then launch Julia from Excel again", out[10]) && out[11] == winget
    end

    # Under WSL, JuliaLaunch passes the startup file's path to Julia through wsl.exe and bash. An
    # apostrophe in it (a user called O'Brien) broke bash's quoting, and a '$' was expanded by
    # the shell wsl.exe ran first. Sends such a path through the real wsl.exe and bash, with
    # printf standing in for Julia. Needs a WSL distribution, so skipped in CI and without one.
    if get(ENV, "CI", "false") != "true" && (try success(`wsl.exe --exec true`) catch; false end)
        @testset "JuliaLaunch under WSL copes with ' and \$ in the startup file's path" begin
            dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
            path = "/mnt/c/Users/O'Brien/a\$b/StartUp_1234.jl"
            ps1 = tempname() * ".ps1"
            write(ps1, """
                [void][Reflection.Assembly]::LoadFrom('$dll')
                \$inner = [CallFromExcel.AddIn.JuliaLaunch]::WslJuliaCommand('printf', "'[%s]\\n'", '$(replace(path, "'" => "''"))')
                \$psi = New-Object System.Diagnostics.ProcessStartInfo
                \$psi.FileName = 'C:\\Windows\\System32\\wsl.exe'
                \$psi.Arguments = [CallFromExcel.AddIn.JuliaLaunch]::WslArguments(\$inner)
                \$psi.UseShellExecute = \$false
                \$psi.RedirectStandardOutput = \$true
                \$p = [System.Diagnostics.Process]::Start(\$psi)
                \$p.StandardOutput.ReadToEnd()
                \$p.WaitForExit()
                """)
            out = readlines(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ps1`)
            rm(ps1)
            @test filter(!isempty, out) == ["[--load]", "[$path]"]
        end
    end

    # serve_xl must write its port file to the same folder the add-in reads it from. The add-in
    # uses .NET's Path.GetTempPath(), which reads TMP before TEMP; serve_xl used ENV["TEMP"], so
    # where the two differ, Excel never found the session. Sets them to different folders and
    # compares serve_xl's folder with the built add-in's (as folders: .NET expands 8.3 short names
    # such as PHILIP~1 in the path and Julia doesn't, so the strings can differ).
    @testset "serve_xl uses the add-in's comms folder when TMP and TEMP differ" begin
        dll = joinpath(@__DIR__, "..", "dotnet-addin", "dist", "CallFromExcel.AddIn.dll")
        tmp, temp = mktempdir(), mktempdir()
        withenv("TMP" => tmp, "TEMP" => temp) do
            script = "[void][Reflection.Assembly]::LoadFrom('$dll'); [CallFromExcel.AddIn.FileUtils]::LocalTemp()"
            addin_folder = readchomp(`powershell -NoProfile -NonInteractive -Command $script`)
            @test samefile(CallFromExcel._default_commsfolder(), addin_folder)
            @test samefile(dirname(addin_folder), tmp)
        end
        rm(tmp; recursive=true); rm(temp; recursive=true)
    end
end

# The VBA-side test suite (modTest.RunTests, in workbooks\TestCallFromExcel.xlsm) exercises
# both call paths against a real Excel/Julia session (via COM automation): CallFromExcel.JuliaCall/
# JuliaEval called from VBA code (the ComServer COM class), and =JuliaCall(...)/=JuliaEval(...)
# called from a genuine worksheet formula (the XLL) - these are separate implementations with
# their own real bugs found only by testing the worksheet-formula path directly (see that suite's
# own WsTest_* functions and dotnet-addin\README.md), unlike the old xlam-only world where a bare
# JuliaCall behaved identically either way. This needs a live Excel session, which GitHub Actions
# doesn't have - so this is skipped in CI (detected via the "CI" environment variable GitHub
# Actions sets, not the OS: the CI matrix in .github/workflows/Runtests.yml currently runs on
# windows-latest, so checking Sys.iswindows() alone wouldn't distinguish CI from a local run).
# Locally, this also needs the C# add-in already built (dotnet build -c Release in
# dotnet-addin\CallFromExcel.AddIn) and registered for COM (dotnet-addin\installer\Install.ps1) -
# dotnet-addin\vba\Run-VbaTests.ps1 itself fails clearly (rather than skipping silently) if either
# precondition isn't met.
#
# Passes "--project=" & pkgdir(CallFromExcel) through to RunTests (and from there to JuliaLaunch), so
# the Julia process RunTests launches to actually serve the VBA-side tests is always this exact
# working copy - not whatever package version happens to be in the machine's default Julia
# environment, which could silently differ from the code under test here. pkgdir gives the correct
# folder on any machine, unlike a hard-coded path.
if get(ENV, "CI", "false") != "true" && Sys.iswindows()
    @testset "VBA test suite (dotnet-addin's modTest.RunTests, requires Excel + the C# add-in built/registered locally)" begin
        script = joinpath(@__DIR__, "..", "dotnet-addin", "vba", "Run-VbaTests.ps1")
        project_options = "--project=$(pkgdir(CallFromExcel))"
        proc = run(ignorestatus(`powershell -ExecutionPolicy Bypass -File $script -CommandLineOptions $project_options`))
        @test proc.exitcode == 0
    end
end
