# @betteroffice/pptx-native

Native Node.js bindings for the BetterOffice PPTX Rust facade. The package is intended for servers, CLIs, Electron applications, and other Node.js hosts that need the native engine without a browser or WebAssembly runtime.

The API follows Node.js conventions while preserving the data model, limits, collaboration behavior, and editing semantics of the Rust facade and its Python binding. Facade operations return promises and run in call order on one native worker. That also keeps the thread-affine Rust facade on its owning thread without blocking the JavaScript event loop.
