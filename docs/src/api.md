# Julia API

```@meta
CurrentModule = CallFromExcel
```

The Julia functions of the CallFromExcel package that you might call yourself. The worksheet functions, which are part of the Excel add-in, are documented in [Worksheet functions](functions.md) and [VBA functions](vba_functions.md).

## Installation

```@docs
installme
```

## Your own worksheet functions

```@docs
@excel
expose_to_excel
```

## Object handles

```@docs
handle_label
```

## Attaching to a running Julia session

```@docs
serve_xl
stop_server
server_status
```

## Debugging

```@docs
display_results
answer_again
```

The variables `last_question`, `last_answer`, `args_from_xl` and `kwargs_from_xl` are described in [Debugging](debugging.md).

## Types

```@docs
ExcelError
```
