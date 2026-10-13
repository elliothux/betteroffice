# @betteroffice/xlsx-native

Native Rust bindings for the BetterOffice XLSX engine in Node.js, Bun and Deno.
Facade operations return promises and run in call order on one native worker.

```sh
npm install @betteroffice/xlsx-native
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@betteroffice/xlsx-native';

const workbook = await openWorkbook(await readFile('input.xlsx'));
await writeFile('output.xlsx', await workbook.save());
```

Prebuilt binaries cover macOS arm64/x64, GNU Linux arm64/x64, and Windows x64.
The package shares its version with `@betteroffice/xlsx`.
CommonJS `require()` is also supported.
