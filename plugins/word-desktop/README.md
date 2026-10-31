# Word Desktop for Codex

**Developer:** Rashmiranjan Behera · **Website:** [rbehera.in](https://rbehera.in)  
**Version:** 0.1.1 beta · **License:** [MIT](LICENSE)

This repository contains the **runnable plugin folder, installation instructions and license notices only**. Research files, test outputs and development/demo materials are excluded. The scripts needed to run the plugin are in [`plugins/word-desktop`](plugins/word-desktop).

## Plugin description

Word Desktop connects Codex on your Windows computer to your installed Microsoft Word. It lets Codex identify and read a selected open document, revise a working copy, insert native reviewer comments, make genuine tracked replacements, save the copy and export it to PDF using Word. A separate Poppler renderer creates page images for Codex to inspect.

It is designed for research papers and academic documents where preserving formatting, reviewing changes and checking the final page layout matter. It requires a running desktop Word session under your Windows account; no LibreOffice or additional AI API key is needed.

## What it solves

- **A dependable connection to desktop Word:** Codex can work with the exact document you select, including its main text, headings and tables, through a focused set of document tools.
- **Safer revisions:** edits target working copies, leaving original files protected by default. Stale document state and ambiguous passages are rejected before editing.
- **Review inside Word:** ordinary replacements, native comments and genuine tracked changes are distinct operations, so reviewers can see what was changed.
- **Layout verification beyond text:** Word exports the PDF, then Poppler renders ordered page images. Codex can inspect those images for clipping, missing content, broken tables and page-break problems. Text extraction alone does not verify layout.

This beta supports focused plain-text edits and the tested workflows described below. Complex equations, fields, mixed-format targets and other unsupported content require manual review; see **Compatibility and limits**.

## Download and verify

For ZIP import, [download the repository](https://github.com/rayx1/codex-word-desktop/archive/refs/heads/main.zip), extract it and package only `plugins/word-desktop` as explained in option 1. Review its `plugin.json`, MIT license and third-party notices before installation. For the GitHub marketplace option, Codex downloads the repository through the commands in option 2. After either installation, run the setup/connection checks below; downloading or importing a package alone does not verify Word access.

## Choose an installation option

| Option | Use it when | Updates |
| --- | --- | --- |
| **1. ZIP import in Codex desktop** | You prefer the app's **New Plugin** upload dialog. | Make a new plugin ZIP and use the app's update/reimport flow. |
| **2. GitHub marketplace** | You prefer installing and refreshing from this repository. | Refresh the marketplace, then restart Codex. |

Choose one option. Both use the same plugin and require the software below. A ZIP import does not establish GitHub synchronization. No prebuilt ZIP is stored in this repository.

## Requirements

- Windows and a licensed, working desktop Microsoft Word installation.
- Codex desktop with plugin archive upload (option 1), or a Codex client supporting Git marketplaces (option 2).
- Windows PowerShell 5.1 (`powershell.exe`, included with Windows).
- Python 3.10+, available as `python` on the PATH seen by Codex.
- Poppler's `pdfinfo` and `pdftoppm` on PATH, or an available Codex-bundled copy. **Poppler is not bundled in this ZIP.**

Word performs DOCX → PDF; Poppler separately renders PDF → page images. No LibreOffice or additional AI API key is required. See [external licenses and sources](THIRD_PARTY_NOTICES.md).

## Prerequisite software setup

Skip software that is already installed and passes the checks below. Use Windows-native applications for this integration; a WSL, remote or cloud session cannot attach to your desktop Word session.

### 1. Microsoft Word

Install the **Windows desktop version** of Word using your existing Microsoft 365/Office entitlement or your organization's installer. Start with [Microsoft's Word page](https://www.microsoft.com/en-us/microsoft-365/word) if you need the product. Word for the web is insufficient. Open Word once, complete activation/sign-in and dismiss first-run dialogs. Under **File → Account → About Word**, note the version/build. Keep Word open under the same Windows account that runs Codex.

### 2. Codex desktop and its CLI

Install/update the Windows desktop app using [OpenAI's Windows setup page](https://learn.chatgpt.com/docs/windows/windows-app), then sign in and open Codex. Select the **Windows native** agent environment. **Option 1 does not require Codex CLI commands.** For option 2, open a fresh PowerShell terminal and check:

```powershell
codex --version
codex plugin --help
codex plugin marketplace add --help
```

If `codex` is not found, follow [OpenAI's Codex CLI installation guide](https://learn.chatgpt.com/docs/codex/cli), selecting its **Windows** instructions, and ensure the executable directory is on your user PATH. Restart the terminal. The installed CLI must expose `plugin` and local `marketplace add`; update it if these commands are absent. Command syntax was checked with `codex-cli 0.166.0-alpha.5.1`; availability can vary by client. [Official local marketplace instructions](https://developers.openai.com/plugins/build/plugins).

### 3. Windows PowerShell 5.1

Use the Windows-provided `powershell.exe`, rather than substituting PowerShell 7's `pwsh.exe`. In PowerShell, check:

```powershell
powershell.exe -NoProfile -Command '$PSVersionTable.PSVersion'
```

Expect version 5.1. No extra PowerShell module or Office add-in is needed. On a managed computer, ask your administrator if this Windows component is unavailable.

### 4. Python 3.10 or newer

Check the existing interpreter first:

```powershell
python --version
where.exe python
```

If absent, install Python from [python.org](https://www.python.org/downloads/windows/), following the [official Windows installation guide](https://docs.python.org/3/using/windows.html). The current Python install manager provides the `python` command after installation; older installers offer an **Add Python to PATH** option. This plugin requires a working `python` command, even if `py` also works. Reopen the terminal and repeat both checks. If `python` opens a Store prompt or the wrong interpreter, resolve command aliases/PATH using the official troubleshooting guide. No `pip install` packages are required.

### 5. Poppler for PDF page images

First check whether the commands are already available:

```powershell
pdfinfo -v
pdftoppm -v
```

If these are missing, the plugin's diagnostic can still find an existing Codex-bundled Poppler. Download/extract the repository and run the check under **Verify setup and connection** before installing another copy. Skip installation if the diagnostic reports both tools as available.

If neither source works, one maintained Windows x64 route is [Miniforge from conda-forge](https://github.com/conda-forge/miniforge). Download its Windows x86_64 installer, follow its installation prompts and keep the Start menu shortcut. Open **Miniforge Prompt**, then run:

```text
conda install -c conda-forge poppler
where.exe pdfinfo
where.exe pdftoppm
```

An existing conda installation can run the same install command. The [conda-forge Poppler package](https://anaconda.org/conda-forge/poppler) supplies these tools. Note the actual folder reported by `where.exe` (usually the environment's `Library\bin`). In Windows, search for **Edit environment variables for your account**, edit the user **Path**, and add that folder as a new entry; retain the other entries. Keep its accompanying DLLs in place. Close/reopen Codex and the terminal, then check `pdfinfo -v` and `pdftoppm -v` again. A conda environment activated only in one terminal does not make its tools available to an already-running desktop app.

Poppler is external GPL-licensed software, supplied separately from this MIT-licensed plugin. It is **not included** in the plugin ZIP. Consult its exact build's license/dependency notices. Miniforge installation and this alternative Poppler build were documented from their maintainers' instructions; the tested renderer used the existing Codex-bundled Poppler 26.07.0.

## Option 1: ZIP import in Codex desktop

1. Select **Code → Download ZIP** on this repository, or [download the repository](https://github.com/rayx1/codex-word-desktop/archive/refs/heads/main.zip), then extract it.
2. Find `plugins/word-desktop` inside the extracted repository. Create a ZIP of **that plugin folder's contents**, including `.codex-plugin` and `.mcp.json`. The plugin ZIP must contain `plugin.json`, `mcp.json`, `server.ps1`, `skills/` and the other runtime files directly at its root (one enclosing plugin folder is also supported). Do not upload the complete repository ZIP: it contains the marketplace and extra files.

   To package all files, including hidden ones, open PowerShell in the extracted repository root and run:

   ```powershell
   Add-Type -AssemblyName System.IO.Compression.FileSystem
   $pluginFolder = (Resolve-Path -LiteralPath '.\plugins\word-desktop').Path
   $zipPath = Join-Path (Split-Path -Parent $pluginFolder) 'word-desktop-0.1.1.zip'
   [IO.Compression.ZipFile]::CreateFromDirectory($pluginFolder, $zipPath)
   ```

   This writes the ZIP beside the plugin folder and refuses to overwrite an existing ZIP. The archive does not include the repository marketplace.

3. In Codex desktop, open **Plugins → New Plugin** (the create/add control's label may vary), select **click to upload**, choose the plugin ZIP and select **Add plugin**. Review any permissions and setup messages. If your client has no archive upload, use option 2.
4. Start a new local chat and follow **Verify setup and connection** below. An upload success does not establish that the local MCP server started.

For updates, create a ZIP from the newer plugin folder and use the app's supported update/reimport control. If only **New Plugin** is available, remove the previous installation before adding the new one; keep your documents/work outputs. Avoid duplicate Word Desktop connections.

**Validation status:** the archive-upload dialog was confirmed in the developer's Codex desktop screenshot. Acceptance of this local MCP package and startup through that dialog have not yet been tested; report any import diagnostics accurately. The repository/CLI route is available separately.

## Option 2: install from the GitHub marketplace

Use a regular terminal under the **same signed-in Windows account as Word**. Do not run as a service, sandbox account, another user or an elevated administrator. Git must be available (`git --version`); if missing, install [Git for Windows](https://git-scm.com/downloads/win) and reopen the terminal.

```powershell
codex plugin marketplace add rayx1/codex-word-desktop --ref main
codex plugin add word-desktop@word-desktop-github
```

Alternatively, run the first command, restart the desktop app, open **Plugins**, choose the **Word Desktop GitHub** source, open **Word Desktop** and select **Install**. Review the connection permissions. Then restart Codex and verify the connection below.

To fetch later repository changes:

```powershell
codex plugin marketplace upgrade word-desktop-github
```

Restart Codex and check the plugin version. Marketplace refresh is not continuous live synchronization. Follow any update prompt; if the client retains an older installation, reinstall from the refreshed marketplace and verify the reported version. [Official marketplace setup and refresh instructions](https://developers.openai.com/plugins/build/plugins).

Keep only one Word Desktop connection. Before switching options, intentionally remove/disable the previous plugin or direct MCP registration; retain saved documents. An earlier extracted-folder installation may use the marketplace name `word-desktop-local`.

## Verify setup and connection

Open Word normally under the same signed-in Windows account as Codex, and resolve Office dialogs. If you downloaded the repository, set `$pluginFolder` to its `plugins/word-desktop` directory and run:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File "$pluginFolder\diagnose.ps1"
```

Fix FAIL items. Closed Word is a WARN: prerequisites may pass while the live Word connection is unverified. The diagnostic never launches Word or reads document contents. Its Codex CLI check applies to option 2; a missing CLI alone does not establish that desktop ZIP import is unavailable. `RemoteSigned` applies only to that PowerShell process; no persistent Windows/Office setting is changed. Group Policy still applies. If signing/download approval is required, follow the normal approval process; do not bypass it.

In a fresh **local** Codex chat, ask: **Use Word Desktop to check the connection and list open documents. Do not edit anything.** Confirm that the tools are present and verify the actual Word version and full document path. The helper attaches to Word already running under your desktop account. Installation alone does not verify the connection.

## Workflow and protection

List documents, explicitly select the intended full path/session ID, read it, then create a new working copy before editing. Originals cannot be edited/saved by the helper. Clean saved originals retain their open identity. Unsaved/dirty sources require explicit consent to Save As, which changes the open identity while preserving the previous disk file.

Use a unique plain-text passage and the latest state token for each write. Choose **ordinary** or **tracked** explicitly; tracked mode creates native Word revisions. Add native comments, save the copy, export with Word, then separately render with Poppler. Codex must inspect every returned page image before claiming layout verification.

Protected, read-only and rights-managed documents are not edited. Stale tokens and ambiguous passages are rejected. On errors/timeouts, inspect the working copy before retrying; mutations are not automatically retried or rolled back. The helper never closes unrelated documents, quits Word or changes Office security.

Outputs remain in `%LOCALAPPDATA%\CodexWord\work` until manually removed. Requested content is returned to Codex and processed under the host's account/settings. The helper itself has no network/upload service.

## Compatibility and limits

The disposable sample was tested with Word 16.0.17932 (WINWORD.EXE 16.0.17932.21056), Windows PowerShell 5.1 and Poppler 26.07.0. Comments, genuine tracked replacements, working-copy preservation, Word PDF export and inspection of both sample pages were verified. The 0.1.1 setup checks, renderer regression and MCP initialization were also tested.

The folder distribution was checked with an isolated Codex CLI home: marketplace registration, plugin installation and the 10-tool MCP handshake passed. The documented ZIP command produced a valid 14-file archive, including the hidden compatibility files and excluding the marketplace. These checks do not establish successful desktop ZIP import or live Word access through a new installation.

This is a beta, not certification for all Word builds or research papers. Scope: one registered Word instance, main-story text/headings/table cells, unique plain-text anchors up to 250 characters. Mixed-format targets, fields, equations and existing revisions are not edited. Structured header/footnote/text-box extraction, accepting/rejecting revisions and PDF markup export are not implemented. Long papers and complex tables require their own verification.

## Troubleshooting

- **Not connected / `0x800401E3`:** open Word in the same account and resolve dialogs, then retry.
- **`0x80070520` / sandbox account:** use the signed-in interactive desktop session. Do not change DCOM, Office security or document permissions.
- **Python/Poppler missing:** inspect `diagnose.ps1`, fix PATH and restart Codex. Diagnostic `-Python` only checks that interpreter; plugin startup still needs `python` on Codex's PATH.
- **Tools absent:** use a new local chat, check `codex plugin list`, and retain one connection only.
- **Stale state/failed edit:** re-read the copy and reassess; do not blindly retry a timed-out write.
- **Word restarted:** restart the helper/chat; document IDs are scoped to one helper session.

Diagnostic JSON (`diagnose.ps1 -Json`) may include account names/local paths; review it before sharing.

## Shutdown and removal

Disconnecting the helper leaves Word and its documents open. For option 1, remove Word Desktop using the desktop app's plugin management controls. For option 2:

```powershell
codex plugin remove word-desktop@word-desktop-github
codex plugin marketplace remove word-desktop-github
```

For a legacy extracted-folder installation, use `word-desktop@word-desktop-local` and `word-desktop-local` instead. Restart Codex. Save needed documents before manually removing a downloaded folder or work outputs. No Office add-in, service, scheduled task or firewall rule needs removal.

GitHub distribution is separate from OpenAI's public Plugin Directory. No directory approval or Microsoft/OpenAI affiliation is claimed.
