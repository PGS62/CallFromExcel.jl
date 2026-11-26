# Builds the documentation site. From the repository root:
#
#     julia --project=docs -e "using Pkg; Pkg.instantiate()"    (once)
#     julia --project=docs docs/make.jl
#
# then open docs/build/index.html. Or, for a preview that refreshes as pages are edited,
# with LiveServer.jl in your default environment:
#
#     julia --project=docs -e "using LiveServer; servedocs()"
#
# The images live in images/ at the repository root, where the README uses them, and are
# copied into docs/src/assets/images/ (not committed) for each build.

using Documenter
using CallFromExcel

const ROOT = dirname(@__DIR__)
const IMAGES = joinpath(@__DIR__, "src", "assets", "images")
mkpath(IMAGES)
for f in readdir(joinpath(ROOT, "images"))
    cp(joinpath(ROOT, "images", f), joinpath(IMAGES, f); force=true)
end

makedocs(;
    sitename = "CallFromExcel.jl",
    authors = "Philip Swannell",
    modules = [CallFromExcel],
    format = Documenter.HTML(;
        prettyurls = get(ENV, "CI", "false") == "true",
        canonical = "https://pgs62.github.io/CallFromExcel.jl/stable/",
        edit_link = "dev",
        size_threshold = 400_000,
    ),
    pages = [
        "Home" => "index.md",
        "Installation" => "installation.md",
        "The Julia ribbon tab" => "ribbon.md",
        "Worksheet functions" => "functions.md",
        "Object handles" => "handles.md",
        "Your own worksheet functions" => "own_functions.md",
        "Calling from VBA" => "vba.md",
        "VBA functions" => "vba_functions.md",
        "Marshalling" => "marshalling.md",
        "Attaching to a running Julia session" => "attaching.md",
        "Debugging" => "debugging.md",
        "Performance and alternatives" => "performance.md",
        "How CallFromExcel works" => "how_it_works.md",
        "Julia API" => "api.md",
    ],
    # The worksheet functions are C#, documented in these pages; only the Julia API page
    # uses docstrings, so don't require every docstring in the package to appear.
    checkdocs = :none,
    # Examples need Excel, so they aren't run.
    doctest = false,
)

# Publishes to the gh-pages branch, served by GitHub Pages at
# https://pgs62.github.io/CallFromExcel.jl/ - only when run by GitHub Actions
# (.github/workflows/Documentation.yml); run locally, it does nothing. Pushes to dev update
# /dev; release tags (made by TagBot, with the DOCUMENTER_KEY deploy key so that they
# trigger the workflow) update /stable.
deploydocs(;
    repo = "github.com/PGS62/CallFromExcel.jl.git",
    devbranch = "dev",
    push_preview = false,
)
