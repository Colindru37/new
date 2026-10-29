# AI prompt: converting the remaining Python bots to Blue Prism

For the team. Give this file, the `rpa-bp` skill, the Python bot with its `config.ini`, and the
GK Revoke XML files (`xml\`) to your AI, then paste the prompt in section 4. Section 3 is every
mistake the AI made when the first five bots were converted; it will make the same ones again
unless it reads them.

## 1. What is the truth

- **The GK Revoke objects and processes are the absolute reference.** The five processes and
  seven objects in `xml\` run end to end on UAT (2026-10-17). When the Python and the GK Revoke
  XML do the same thing differently (popups, waits, table reads, uploads, Word labels, failure
  handling, page names, how a process uses the queue, how it logs), the XML is right and the
  Python is wrong. The new bot copies the GK Revoke page layout, names and patterns. The AI does
  not "improve" them. If it thinks one is wrong it says so, with the reason, and still follows
  it until you agree.
- **The INI of the Python bot stays as it is.** The objects read the keys the Python read. Only
  XPath values change, and only when the real page proves the old one wrong.
- **`docs\GUIDE.md`** describes every object action, process, environment variable and INI key.
  **`SKILL.md`** and `reference\` of the skill say how to write Blue Prism XML that imports.
- The Python's **loop logic and INI keys** are what gets converted. The Python's exception
  handling, sleeps and search-box navigation are not (section 3).

Three things in the GK Revoke XML are NOT to be copied:

1. System exceptions marked **Retry = True**. The failure handler has already reassigned the
   task, so a retry cannot work (that is why the queues run with Max Attempts = 1). New
   processes: Retry = False.
2. **Row-by-row loops** nobody sized yet (BlueSpace groups and teams, RCRT rows, the Master
   queue list). Ask the row count before copying the loop; above ~50 rows read the table in one
   Execute Script call like `Get Roles` does.
3. **Alice/BlueSpace XPaths.** Copy the pattern, not the locator.

## 2. How the work is checked

The AI writes a generator (`build_<object>.py` with the skill's `bp_gen`), never XML by hand.
After every change:

```
python build_<object>.py                                  -> out\BPA Object - GK Revoke <Bot>.xml
python <skill>\lib\validate.py out --flow --estate --strict --with <folder with the GK Revoke xml>
python <skill>\lib\compile_check.py out
```

Both must end with ALL CLEAN / ALL COMPILED. Then import the XML (File > Import) and test
**one action at a time in Object Studio** with real inputs before building the next page
(`11 - Debugging step by step.txt`, part B). A full process run on a test ticket comes last,
with Stop After Next Item = True. If the AI cannot run something, it says so and names the
exact error text to look for. It never writes "should work".

## 3. Mistakes already made on this project. Do not repeat them.

### 3.1 Structure that Blue Prism rejects silently or at run time

1. **Processes cannot hold Code stages.** Import accepts them, Process Validation says nothing,
   the run dies at the first one. Every piece of logic goes into an object action
   (`Proc.code_home('GK Revoke Common', '<Process> - ')` in the generator does this).
2. **Every object sheet needs Start linked to End**, including Initialise. Seven objects shipped
   without it and "failed at Load Config" with no message.
3. **A code stage input and output with the same name is one C# parameter twice** ("duplicate"
   on import). The generator renames inputs `<name>_In`; do not bypass it.
4. **Blocks cannot nest** and Blue Prism's internal objects are referenced by class name
   (`Blueprism.Automate.clsWorkQueuesActions`), not "Work Queues". Use `work_block()` and the
   generator's `INTERNAL_OBJECTS` map.
5. **Stage names become file names in the compile gate.** A `?` in a stage name fails on Windows
   csc. Keep stage names to letters, digits, spaces.

### 3.2 Browser automation against the real sites

6. **A locator in the INI used only by commented-out Python is unused.** Get Open RITMs read a
   column the Python never read and failed on the real list. Trace each key to a live line.
7. **Open records by URL, never by the portal's search box**, and fix every sibling: the banner
   search was replaced in Read RITM Details and left in Open Task, which then failed the same
   way a day later. After any fix, grep the generator for the same locator and path.
8. **Site popups.** Alice raises two `confirm()` dialogs on every page load. The Browser accepts
   a prompt that blocks a command and re-sends it. Do not add per-page alert handling.
9. **Element click intercepted** (a cloned sticky header over a link): scroll, retry, then
   click via script. Already in `Click`; do not add waits around it.
10. **A frame that is reloading** fails the switch with a stale element or "unknown error".
    `Switch To Frame` retries for the implicit wait. The Python's `time.sleep(1)` was that.
11. **Hidden file inputs are never visible.** `Upload File` waits for presence only. The first
    version waited for visibility and timed out on the attachment dialog.
12. **Header rows.** The real Alice roles table starts with a `td` header row with no checkbox;
    the Python never hit it because it only looked at the matching row. Any "read every row"
    conversion of a "find the matching row" loop must tolerate filler rows.
13. **Size every loop before converting it.** A user had 1078 roles; row-by-row reads were
    ~3000 WebDriver calls per region. Ask the count on the real site; above ~50 rows read the
    table in one `Execute Script` call (`document.evaluate` on the configured XPath, rows
    separated by `\u001e`, fields by `\u001f`) and parse in a code stage. Get Roles is the model.
14. **Positional columns are per-account view settings.** The INI's `td[11]` was the prod maker
    bot's personalised ServiceNow list; the checker INI said `td[13]`; UAT accounts differed
    again. Locate columns by header text at run time (Update RCRT Status reads the header row
    once, calibrates on `Number`, falls back to the INI, raises a clear error when a required
    column is absent). When an audit flags a positional locator, fix it before UAT.
15. **Read-only forms.** A task opened from the RITM's related list was read-only on UAT for
    that account; the same task by its own URL was editable. Open records the way the list does.
16. **Downloads** need the Edge prefs at session start (`download.default_directory`,
    `download.prompt_for_download: false`), and enterprise policy may still override them.

### 3.3 Office, mail, files

17. **Mandatory sensitivity labels.** Word refuses the first `SaveAs` of an unlabelled document
    ("Command failed"). Label the new document before its first save. This was predicted in an
    audit and shipped the wrong way anyway.
18. **`SensitivityLabel.SetLabel(LabelInfo, Context)` takes two LabelInfo objects.** Passing the
    justification string as the second argument is `DISP_E_TYPEMISMATCH`. Justification is a
    property of the LabelInfo. For every late-bound Office call, write the VBA signature from the
    documentation in a comment before writing the C#.
19. **Name the COM member in every COM error.** Word's own messages are generic; "Build Evidence
    Docs: Type mismatch" cost a run, "Word SetLabel: Type mismatch" cost nothing.
20. **SMTP: match the estate's VBO**, anonymous relay, no TLS, no `UseDefaultCredentials`.
21. **Trackers**: whitespace written into an empty cell is a value; an open Excel fails the write;
    the `.lock` file must time out (10 min); "4711" must match a numeric cell; whole numbers come
    back without `.0`. Test the explicitly-empty case, not only the unset one.
22. **Clear Folder deletes everything in the configured folder.** A wrong `attachment_folder_path`
    deletes a user's folder.

### 3.4 Configuration

23. **A change that needs an INI or Environment Variable edit is listed first**, not in a
    footnote. The `snow_accepted_task_states` key was mentioned in a caption, the run failed on
    "state not accepted", and the failure path reassigned the ticket to the prod ops group:
    ticket lost. On UAT the ops group key must be the tester's own group.
24. **INI keys are case-insensitive.** Adding `bluespace_url` next to `blueSpace_url` is a
    duplicate and Load Config refuses the file.
25. **`{}` in an XPath is not a C# format slot.** Replace it; `string.Format` throws.
26. **Say which files to re-import.** State the list every time something changes.

### 3.5 How the AI reasons

27. **The screenshot beats the AI's memory of the code.** Three times a confident diagnosis was
    contradicted by a screenshot it had already been given. Make it re-read the picture.
28. **"Audit the other steps" means every external application the run touches**: ServiceNow,
    the portal, Excel trackers, Word, Outlook/SMTP, Edge downloads, the file system. Enumerate
    them first, then audit each.
29. **A test built from the same mental model as the code checks that the AI agrees with
    itself.** Test against the real page, in Object Studio, one action at a time.
30. **A faithful conversion is not the goal when the Python is wrong.** The Python checker failed
    the ticket on a user without an Edit link while its maker had declared the user revoked; the
    conversion judges both the same way and says so in the guide's "deliberate differences".
    Every divergence gets written down there.
31. **One-minute loops beat twenty-minute runs.** A single object action in Object Studio with
    real inputs finds a Word or tracker bug in a minute. Do that before a full run.
32. **Write the mistake down** the moment it is found (what happened, how it was discovered, the
    fix, the rule). Then look for siblings of the same mistake before testing again.

## 4. The prompt

> You are converting the Python/Selenium bot {BOT} (`{maker}.py`, `{checker}.py`, their
> `config.ini`) into Blue Prism with the attached `rpa-bp` skill.
>
> Before writing anything: read `DEV_PROMPT_python_to_bp.md` in full, then `GUIDE.md`, then
> `SKILL.md` and `reference\` of the skill, then the attached GK Revoke XML files.
>
> The attached GK Revoke objects and processes are the absolute reference. They run end to end
> in UAT. Wherever the Python and the GK Revoke XML do something differently, do it the GK
> Revoke way: page layout, names, popup handling, waits, table reads, uploads, Word labels,
> failure handling, queue use, logging. Do not improve their patterns. If you think one is
> wrong, say so and why, then follow it anyway unless I agree. Three exceptions you must not
> copy: system exceptions with Retry = True (use Retry = False); row-by-row loops over tables
> nobody has counted (ask me the count first); Alice/BlueSpace-specific XPaths.
>
> Method:
> 1. Read the Python bot end to end. List its steps, the Selenium calls per step, every INI key
>    it uses from a live (not commented) line, and every external application it touches.
> 2. Ask me, before building: the row count of every table it loops over, whether the site
>    shows popups, which ServiceNow list views the bot account sees, and for every positional
>    locator (`td[n]`, `[n]`) whether the position comes from a per-user view. If it can,
>    locate by header text instead.
> 3. Map the steps onto the existing objects (`GK Revoke Browser`, `ServiceNow`, `Tracker`,
>    `Evidence`, `Common`). Write ONE new object for the new website, `GK Revoke {Bot}`,
>    generated with `bp_gen`, pages named like `GK Revoke Alice`. Then a Maker and a Checker
>    process built like the Alice ones. Processes hold no code.
> 4. Build one object page at a time. After each: run validate.py and compile_check.py, give me
>    the XML to import, and tell me the exact inputs to test that action with in Object Studio.
>    Do not build the next page until I confirm the test.
> 5. With every delivery: the list of files to import, and every INI key or Environment
>    Variable I must add or change, listed first.
> 6. Record every mistake you find, yours or the Python's, and every deliberate difference
>    from the Python.
>
> Do not: hand-edit XML, add a DLL, add Selenium, put a Code stage in a process, hard-code a
> locator, change the INI's keys, deliver a change that needs a config edit without listing
> the edit, or say "should work" about anything you could not execute.

## 5. Fill-in table for the remaining bots

| Bot | Maker script | Checker script | Object to write | First ask |
|---|---|---|---|---|
| {BOT} | `{maker}.py` | `{checker}.py` | `GK Revoke {Bot}` | row counts, popups, task form views, label policy |

The Master already dispatches by application name from the application list; add the new queue
name to its `Which Application?` choice and a queue environment variable.
