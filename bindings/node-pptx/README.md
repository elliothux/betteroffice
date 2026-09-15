# @betteroffice/pptx-node

Native Node.js bindings for the BetterOffice PPTX Rust facade. The package is intended for servers, CLIs, Electron applications, and other Node.js hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, collaboration behavior, and editing semantics of the Rust facade and its Python binding. Each presentation stays on one native worker because the Rust facade is thread-affine; expensive work therefore remains asynchronous without blocking the JavaScript event loop.
