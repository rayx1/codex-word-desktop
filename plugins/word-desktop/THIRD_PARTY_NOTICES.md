# License and external components

Word Desktop's scripts, configuration, instructions and original icon are licensed under [MIT](LICENSE), copyright 2026 Rashmiranjan Behera. MIT does not license Microsoft Word, Windows, Python or Poppler.

No third-party executables/libraries are bundled:

- **Microsoft Word:** proprietary, separately licensed Microsoft software. Provide a working licensed desktop installation. Word performs document-to-PDF rendering. [Microsoft Word](https://www.microsoft.com/microsoft-365/word).
- **Windows PowerShell 5.1 / .NET Framework:** existing Windows components, subject to their Microsoft/Windows licenses.
- **Python:** separately installed under the Python Software Foundation license. The plugin renderer uses only its standard library. [Python license](https://docs.python.org/3/license.html).
- **Poppler:** external `pdfinfo` and `pdftoppm` programs. Poppler is GPL-licensed; consult the exact build's COPYING and dependency notices (commonly GPL-2.0-or-later). This ZIP includes neither Poppler binaries nor its source. The user supplies it on PATH or uses an available Codex-bundled copy. If redistributing Poppler yourself, preserve the exact build's required notices and corresponding source obligations. [Project/source](https://poppler.freedesktop.org/), [conda-forge recipe/license metadata](https://github.com/conda-forge/poppler-feedstock/blob/main/recipe/meta.yaml).

`python-docx`, PDFium, Pillow and LibreOffice are not required or distributed. Microsoft Word, Codex and OpenAI are separate products; Word Desktop is an independent integration.
