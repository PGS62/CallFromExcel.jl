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
    #For compatibility with VBA function Unserialise (which should be exact equivalent of decode_from_excel)
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
    # format) for a Vector{Any} containing a NaN, a non-Float64 element, or with rank > 9 (matching
    # VBA's own ReDimVariantArray MAX_RANK, modUnserialise.bas).
    @test CallFromExcel.encode_for_xl(Any[1.0, 2.0, NaN]) == CallFromExcel.encode_for_xl([1.0, 2.0, NaN])
    @test !startswith(CallFromExcel.encode_for_xl(Any[1.0, 2.0, NaN]), "V")
    @test CallFromExcel.encode_for_xl(Any[1.0, 2, 3.0]) == CallFromExcel.encode_array_general(Any[1.0, 2, 3.0])
    @test !startswith(CallFromExcel.encode_for_xl(Any[1.0, 2, 3.0]), "V")
    @test CallFromExcel.encode_for_xl(Any[]) == CallFromExcel.encode_array_general(Any[])

    # Rank 3-9: one method (encode_for_xl(x::Array{Float64,N})) handles every rank, since
    # Vector{Float64}/Matrix{Float64} are just Array{Float64,1}/Array{Float64,2} and the bulk
    # byte-reinterpret doesn't care how many dimensions the buffer is viewed as. Decoded by the
    # Case 86 'V' branch's Case Else (rank 3-9) in Unserialise (modUnserialise.bas), which reuses
    # ParseDims/ReDimVariantArray/AssignByRank - the same helpers the general "*" format's own
    # >=3-D handling already used.
    let a3d = reshape(collect(1.0:24.0), 2, 3, 4)
        @test startswith(CallFromExcel.encode_for_xl(a3d), "V3,2,3,4;")
        @test round_trip(a3d)
        @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl(a3d)) isa Array{Float64,3}
        # dynamic AbstractArray fallback also generalises to rank 3-9, not just 1-2
        @test CallFromExcel.encode_for_xl(Array{Any}(a3d)) == CallFromExcel.encode_for_xl(a3d)
    end

    # Rank > 9 must fall back to the general "*" format - matches VBA's own ReDimVariantArray
    # MAX_RANK, which Case 86's Case Else would otherwise be asked to exceed.
    let a10d = reshape(collect(1.0:2.0^10), fill(2, 10)...)
        @test !startswith(CallFromExcel.encode_for_xl(a10d), "V")
        @test round_trip(a10d)
    end

    let x3d = fill(1.0, 2, 2, 2) |> a -> Array{Any}(a)
        @test startswith(CallFromExcel.encode_for_xl(x3d), "V3,")  # now gets "V", not the general format
        @test CallFromExcel.encode_for_xl(x3d) == CallFromExcel.encode_for_xl(fill(1.0, 2, 2, 2))
    end

    # decode_xl_array_v (src/decode.jl) - the VBA -> Julia direction, added alongside
    # TrySerialiseArrayAsV (modSerialise.bas). Real round trips are possible now that
    # decode_from_xl understands "V", unlike the encode-only assertions above.
    @test round_trip([1.0, 2.0, 3.0])
    @test round_trip([1.0 2.0 3.0; 4.0 5.0 6.0])  # non-square, to catch a row/column transposition bug
    @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl([1.0, 2.0, 3.0])) isa Vector{Float64}
    @test CallFromExcel.decode_from_xl(CallFromExcel.encode_for_xl([1.0 2.0; 3.0 4.0])) isa Matrix{Float64}

    # Hand-built, VBA-style big-endian hex (plain MSB-first hex per element, matching what
    # DoubleToHex in modSerialise.bas actually produces) - confirms decode_xl_array_v works against
    # genuinely VBA-encoded-style input, not only against Julia's own encoder output.
    let be_hex(v) = uppercase(string(reinterpret(UInt64, v), base=16, pad=16))
        manual = "V1,3;" * be_hex(1.0) * be_hex(-2.5) * be_hex(3.14159265358979)
        @test CallFromExcel.decode_from_xl(manual) == [1.0, -2.5, 3.14159265358979]

        # Rank 3, exercising decode_xl_array_v's general reshape path against hand-built VBA-style
        # hex - this is what TrySerialiseArrayAsV's new Case Else (modSerialise.bas) now produces
        # for the Excel -> Julia direction of a 3D+ array.
        vals3d = collect(1.0:24.0)
        manual3d = "V3,2,3,4;" * join(be_hex.(vals3d))
        @test CallFromExcel.decode_from_xl(manual3d) == reshape(vals3d, 2, 3, 4)
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
    let msg = CallFromExcel.decode_from_xl(CallFromExcel.srv_eval_inner("1+\"1\""))
        @test msg == "#MethodError: no method matching +(::Int64, ::String) The function `+` \
                     exists, but no method is defined for this combination of argument types. \
                     Julia REPL has more details and stacktrace!"
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
        # The port file holds the port, then the session token, for Excel to read.
        @test readlines(CallFromExcel.portfile()) == [string(CallFromExcel.xlport[]), token]
        @test length(token) == 32
        @test CallFromExcel.decode_from_xl(String(post("/eval", "1+1").body)) == 2
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
        @test CallFromExcel.last_question == "40+2"
    end

    # _attach_to_excel refuses to start a second, useless listener when one is already serving the
    # same pid - starting a fresh session again for 4321 without stopping the previous one first.
    @test_throws String CallFromExcel._attach_to_excel(true, 4321)
    CallFromExcel.stop_server()

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
