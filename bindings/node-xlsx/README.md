# @betteroffice/xlsx-node

Native Node.js bindings for the BetterOffice XLSX Rust facade. The package is intended for servers, CLIs, Electron applications, and other Node.js hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, calculation behavior, and editing semantics of the Rust facade and its Python binding. Expensive work runs outside the JavaScript event loop and workbook bytes use `Buffer`.
