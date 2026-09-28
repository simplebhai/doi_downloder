DOI PAPER DOWNLOADER  v2.1
==========================

Downloads legally open-access PDFs for every DOI in an Excel file and saves
each one as <Record_ID>.pdf  (e.g. UCP-001.pdf, UCP-002.pdf ...).


1. ONE-TIME SETUP (Windows) - makes DOI_Paper_Downloader.exe
---------------------------------------------------------------
 a) Install Python 3.10+ from https://www.python.org/downloads/
    IMPORTANT: tick "Add python.exe to PATH" during installation.
 b) Put all these files in one folder and double-click  build.bat
    (takes 3-5 minutes, needs internet).
 c) When it says DONE, the file  DOI_Paper_Downloader.exe  is in the folder.
    Copy it anywhere (Desktop etc.). Python is no longer needed to run it.

 Want to skip the build? Double-click  run_without_building.bat  instead
 (it still needs Python installed).

 Manual build command (same as build.bat):
   pip install -r requirements.txt pyinstaller
   pyinstaller --onefile --windowed --name DOI_Paper_Downloader --hidden-import openpyxl --hidden-import xlrd --collect-all playwright paper_downloader.py

 If you already built v1: delete the old .venv folder, then run build.bat again.


2. DAILY USE
------------
 1. Excel File     -> Browse -> choose PAPER_LINK.xlsx
 2. Output Folder  -> Browse (default: "Downloaded_Papers" next to the Excel)
 3. Enter your e-mail once (Unpaywall, the main open-access source, needs it;
    it is remembered).
 4. Click  DOWNLOAD ALL PAPERS  and wait.
 5. Click  Open Download_Status.xlsx  when finished.

 OPTIONAL FREE KEYS (recommended - many of your papers are Elsevier/Wiley):
  * Elsevier API key: https://dev.elsevier.com -> "I want an API key" ->
    sign in / register (free) -> Create API Key. Paste it in the app.
    Downloads open-access Elsevier papers (Heliyon, ScienceDirect OA,
    Urban Climate OA ...) through Elsevier's official API.
  * Wiley TDM token: log in at https://onlinelibrary.wiley.com ->
    https://onlinelibrary.wiley.com/library-info/resources/text-and-datamining
    -> accept the TDM terms -> copy the token. Paste it in the app.
    Downloads open-access Wiley papers (incl. AGU / RMetS journals).
  Both are only used for papers that are open access (or, with the
  institution option ticked, papers your network is entitled to).

 FIRST TIME: click  Test Connection  - it checks every paper source and your
 browser, and says exactly what is blocked (if anything).

 WHY A BROWSER WINDOW? Many sites (ScienceDirect, PubMed Central, HAL, Wiley,
 AMS ...) give their free PDFs only to a real web browser. The app therefore
 uses your own Microsoft Edge (or Chrome) in a separate, empty profile - no
 logins or cookies of yours are used. The window is placed off-screen; tick
 "Show browser window" to watch it. It never clicks or solves a CAPTCHA; if a
 site asks "are you human?", that paper is marked Download Failed (with the
 window shown, you may complete the check yourself and the download continues).

 Columns are detected automatically (Record_ID / record id / RECORD_ID,
 DOI / doi ...). If not, pick them from the drop-down lists.
 DOIs may be written as 10.xxxx/..., https://doi.org/10.xxxx/... or doi:10.xxxx/...


3. OUTPUT FOLDER
----------------
 Downloaded_Papers\
   UCP-001.pdf, UCP-002.pdf, ...   <- named by Record_ID (never by DOI,
                                      unless Record_ID is empty)
   Download_Status.xlsx            <- one row per Excel row + "Summary" sheet
   download_log.txt                <- time | Record_ID | status | detail
   .download_state.json            <- resume memory (hidden; do not delete
                                      if you want resume to work)

 Illegal filename characters  / \ : * ? " < > |  are replaced by "_".
 Duplicate Record_IDs get _2, _3 ... so nothing is overwritten.


4. STATUS MEANINGS
------------------
 Downloaded        legal open PDF found, verified, saved as Record_ID.pdf
 Already Exists    file already in the folder - not downloaded again
 Not Open Access   no legal open version / publisher restricts access
 PDF Not Found     listed as open access, but no actual PDF file available
 Download Failed   PDF link exists but download was refused/broken (retry later)
 Server Error      site busy/down - retried automatically on next run
 Timeout           site too slow - retried automatically on next run
 Missing DOI       empty DOI cell
 DOI Invalid       malformed DOI, or not registered at doi.org
 Not Processed     you pressed Stop before this record

 The "Source" / "Source_URL" columns show exactly where each PDF came from.


5. RESUME / RE-RUN
------------------
 Just run again with the same Excel and output folder:
  - existing PDFs are skipped (Already Exists) - no re-download
  - Not Open Access / Invalid results from earlier are reused (no re-check)
  - Timeout / Server Error / Download Failed / Not Processed are retried
 Results from v1 are re-checked automatically by v2 (only real PDFs are kept).
 Options:
  [x] Re-download existing files                  -> overwrite PDFs
  [x] Re-check papers previously found not accessible -> ask the sources again
      (useful months later - papers become open access over time)


6. WHERE PDFs COME FROM (legal sources only, in this order)
-----------------------------------------------------------
 Unpaywall -> OpenAlex -> Crossref publisher links -> Europe PMC (also used
 for every PubMed Central copy) -> Semantic Scholar -> publisher page via
 doi.org -> the same links again in your browser, if plain download failed.
 Every file is checked: must start with %PDF, be > 5 KB and complete (%%EOF).
 HTML/error/login pages are never saved as PDF.

 The program never bypasses paywalls, logins, CAPTCHAs, DRM or institutional
 access. If a site asks for login or a CAPTCHA, the paper is reported as
 not accessible. Up to 3 attempts are made for temporary errors, with pauses,
 and each website is contacted at most once every 1.5 s.

 Papers from subscription journals will be "Not Open Access" - request
 those via your library or the author.

 OPTION "Also download papers my institution subscribes to": off by default.
 Tick it only on your institute's network, for papers you are entitled to.
 Publishers forbid bulk ("systematic") downloading even for subscribers and
 can block a whole campus IP range - keep "Parallel downloads" at 1-2 and
 don't run huge batches with this option.


7. TROUBLESHOOTING
------------------
 * Run  Test Connection . Lines marked BAD name the blocked site and the
   reason (DNS failed / connection reset by firewall / proxy refused ...).
   Campus firewalls often block some sites for programs - try a mobile
   hotspot or home network, or ask IT to allow the listed sites.
 * "Browser unavailable: no usable browser found" -> install Microsoft Edge
   or Google Chrome.
 * If anything unexpected happens, the full technical details are written to
   error_log.txt in the output folder (or in
   %LOCALAPPDATA%\DOI_Paper_Downloader\). Send that file for a fix.
 * "Download_Status.xlsx is open in Excel" -> close it; a copy with a time
   stamp is saved meanwhile.
 * Windows SmartScreen warns about the new .exe -> "More info" -> "Run anyway"
   (normal for self-built programs).
 * Command-line mode (optional):
     python paper_downloader.py --cli PAPER_LINK.xlsx Downloaded_Papers --email you@example.com
     python paper_downloader.py --cli --test --email you@example.com      (connection test)
