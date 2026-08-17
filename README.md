# Windows-Admin-Tools
A collection of tools I found and/or made to administer and work with AD

## Script conventions

### Encoding: UTF-8 **with BOM**, ASCII-only where practical

Every `.ps1` in this repo must be saved as **UTF-8 with a byte-order mark**.

Windows PowerShell 5.1 falls back to the legacy ANSI codepage when a `.ps1` has no
BOM. Any non-ASCII character is then decoded as several junk characters, which
desynchronises the parser and breaks string terminators and brace matching. The
resulting errors are reported on *later lines that are perfectly valid*, so the
message never points at the real culprit, and the script will not run at all.

PowerShell 7 defaults to UTF-8 and parses the same file without complaint, so these
bugs stay invisible until a 5.1 host runs the script. Six scripts in this repo were
affected; see [CHANGELOG.md](CHANGELOG.md) for the 2026-08-17 fix.

Prefer plain ASCII in source. Reserve non-ASCII for output that genuinely needs it
(box-drawing separators, warning symbols) - and even then, only ever in a file that
has a BOM.

**Check before committing:**

```bash
# must print nothing
grep -nP '[^\x00-\x7F]' script.ps1

# must print: ef bb bf
head -c3 script.ps1 | od -An -tx1
```

```powershell
# authoritative check - the real parser, not a regex
$e = $null
[System.Management.Automation.Language.Parser]::ParseFile('script.ps1', [ref]$null, [ref]$e)
$e.Count   # must be 0
```

### Versioning

Each script or tool is versioned independently and follows
[Semantic Versioning](https://semver.org/). Scripts that carry a version keep it in
the comment-based help `.NOTES` block, together with an inline changelog. Larger
tools additionally keep their own `CHANGELOG.md`.

Git tags are **prefixed with the tool name**, because this repo holds multiple
independent scripts - for example `smtp-relay-tester-ps-v1.0.0` and
`calendar-sharing-v1.2.0`. Repository-wide changes are recorded in the root
[CHANGELOG.md](CHANGELOG.md) and are not tagged.
