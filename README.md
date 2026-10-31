# Word Desktop for Codex

**Developer:** Rashmiranjan Behera · **Website:** [rbehera.in](https://rbehera.in)  
**Version:** 0.1.1 beta · **License:** [MIT](LICENSE)

This repository contains the **plugin ZIP, installation instructions, license notices and checksum only**. The development repository, separate source archive, research files, test outputs and demo materials are not published here. The plugin ZIP necessarily includes the scripts/configuration that run it.

## Download and verify

[Download word-desktop-0.1.1.zip](https://github.com/rayx1/codex-word-desktop/raw/refs/heads/main/word-desktop-0.1.1.zip). Compare its hash with [SHA256SUMS.txt](SHA256SUMS.txt):

```powershell
Get-FileHash -LiteralPath .\word-desktop-0.1.1.zip -Algorithm SHA256
```

## Requirements

- Windows and a licensed, working desktop Microsoft Word installation.
- Codex desktop/CLI supporting local plugin marketplaces.
- Windows PowerShell 5.1 (`powershell.exe`, included with Windows).
- Python 3.10+, available as `python` on the PATH seen by Codex.
- Poppler's `pdfinfo` and `pdftoppm` on PATH, or an available Codex-bundled copy. **Poppler is not bundled in this ZIP.**

Word performs DOCX → PDF; Poppler separately renders PDF → page images. No LibreOffice or additional AI API key is required. See [external licenses and sources](THIRD_PARTY_NOTICES.md).

## Prerequisite software setup

Skip software that is already installed and passes the checks below. Use Windows-native applications for this integration; a WSL, remote or cloud session cannot attach to your desktop Word session.

### 1. Microsoft Word

Install the **Windows desktop version** of Word using your existing Microsoft 365/Office entitlement or your organization's installer. Start with [Microsoft's Word page](https://www.microsoft.com/en-us/microsoft-365/word) if you need the product. Word for the web is insufficient. Open Word once, complete activation/sign-in and dismiss first-run dialogs. Under **File → Account → About Word**, note the version/build. Keep Word open under the same Windows account that runs Codex.

### 2. Codex desktop and its CLI

Install/update the Windows desktop app using [OpenAI's Windows setup page](https://learn.chatgpt.com/docs/windows/windows-app), then sign in and open Codex. Select the **Windows native** agent environment. Open a fresh PowerShell terminal and check:

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

If these are missing, the plugin's diagnostic can still find an existing Codex-bundled Poppler. Extract the plugin and run step 2 under **Install** before installing another copy. Skip installation if the diagnostic reports both tools as available.

If neither source works, one maintained Windows x64 route is [Miniforge from conda-forge](https://github.com/conda-forge/miniforge). Download its Windows x86_64 installer, follow its installation prompts and keep the Start menu shortcut. Open **Miniforge Prompt**, then run:

```text
conda install -c conda-forge poppler
where.exe pdfinfo
where.exe pdftoppm
```

An existing conda installation can run the same install command. The [conda-forge Poppler package](https://anaconda.org/conda-forge/poppler) supplies these tools. Note the actual folder reported by `where.exe` (usually the environment's `Library\bin`). In Windows, search for **Edit environment variables for your account**, edit the user **Path**, and add that folder as a new entry; retain the other entries. Keep its accompanying DLLs in place. Close/reopen Codex and the terminal, then check `pdfinfo -v` and `pdftoppm -v` again. A conda environment activated only in one terminal does not make its tools available to an already-running desktop app.

Poppler is external GPL-licensed software, supplied separately from this MIT-licensed plugin. It is **not included** in the plugin ZIP. Consult its exact build's license/dependency notices. Miniforge installation and this alternative Poppler build were documented from their maintainers' instructions; the tested renderer used the existing Codex-bundled Poppler 26.07.0.

## Install

Use a regular terminal under the **same signed-in Windows account as Word**. Do not run as a service, sandbox account, another user or an elevated administrator.

1. Download the ZIP into your current directory and extract it to a permanent location:

   ```powershell
   $pluginFolder = Join-Path $env:LOCALAPPDATA 'CodexPlugins\WordDesktop-0.1.1'
   Expand-Archive -LiteralPath .\word-desktop-0.1.1.zip -DestinationPath $pluginFolder
   ```

   The extracted folder must directly contain `plugin.json`, `mcp.json`, `server.ps1`, `diagnose.ps1` and `.agents`. Do not extract over an existing installation. Keep the extracted folder so the local marketplace remains available.

2. Open Word normally and resolve any Office dialogs, then check setup:

   ```powershell
   powershell.exe -NoProfile -STA -ExecutionPolicy RemoteSigned -File "$pluginFolder\diagnose.ps1"
   ```

   Fix FAIL items. Closed Word is a WARN: prerequisites may pass while the live Word connection is unverified. The diagnostic never launches Word or reads document contents. `RemoteSigned` applies only to that PowerShell process; no persistent Windows/Office setting is changed. Group Policy still applies. If signing/download approval is required, follow the normal approval process; do not bypass it.

3. Register the marketplace included in the ZIP, then install:

   ```powershell
   codex plugin marketplace add "$pluginFolder"
   codex plugin add word-desktop@word-desktop-local
   ```

   You can instead register the marketplace with the first command, restart the desktop app, open **Plugins**, choose the **Word Desktop Local** source, open **Word Desktop** and select **Install**. Review the connection permissions. This is a local installation; do not search for a public directory listing. This GitHub repository holds a ZIP, so register the **extracted folder**, not `rayx1/codex-word-desktop` as a Git marketplace.

   Keep only one Word Desktop connection. Intentionally remove/disable an older plugin or direct MCP registration before switching; keep its saved documents.

4. Restart Codex or start a fresh **local** chat. Ask: **Use Word Desktop to check the connection and list open documents. Do not edit anything.** Verify the actual Word version and full document path. Installation alone does not verify the connection. The helper attaches to Word already running under your desktop account.

## Workflow and protection

List documents, explicitly select the intended full path/session ID, read it, then create a new working copy before editing. Originals cannot be edited/saved by the helper. Clean saved originals retain their open identity. Unsaved/dirty sources require explicit consent to Save As, which changes the open identity while preserving the previous disk file.

Use a unique plain-text passage and the latest state token for each write. Choose **ordinary** or **tracked** explicitly; tracked mode creates native Word revisions. Add native comments, save the copy, export with Word, then separately render with Poppler. Codex must inspect every returned page image before claiming layout verification.

Protected, read-only and rights-managed documents are not edited. Stale tokens and ambiguous passages are rejected. On errors/timeouts, inspect the working copy before retrying; mutations are not automatically retried or rolled back. The helper never closes unrelated documents, quits Word or changes Office security.

Outputs remain in `%LOCALAPPDATA%\CodexWord\work` until manually removed. Requested content is returned to Codex and processed under the host's account/settings. The helper itself has no network/upload service.

## Compatibility and limits

The disposable sample was tested with Word 16.0.17932 (WINWORD.EXE 16.0.17932.21056), Windows PowerShell 5.1 and Poppler 26.07.0. Comments, genuine tracked replacements, working-copy preservation, Word PDF export and inspection of both sample pages were verified. The 0.1.1 setup checks, renderer regression and MCP initialization were also tested.

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

Disconnecting the helper leaves Word and its documents open. To uninstall:

```powershell
codex plugin remove word-desktop@word-desktop-local
codex plugin marketplace remove word-desktop-local
```

Restart Codex. Save needed documents before manually removing the extracted folder or work outputs. No Office add-in, service, scheduled task or firewall rule needs removal.

GitHub distribution is separate from OpenAI's public Plugin Directory. No directory approval or Microsoft/OpenAI affiliation is claimed.
