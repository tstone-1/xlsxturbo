# rust_xlsxwriter: upstream defects and name validation

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

## Upstream defects belong upstream — file them

**When rust_xlsxwriter is what is wrong, report it to jmcnamara/rust_xlsxwriter.** Not
instead of the local workaround, but alongside it: work around the defect so users are not
holding a corrupt file, then file the report so the workaround has an end. This is the
standing rule for this repository, and it was learned the expensive way.

The case that established it: a `data_bar` conditional format beside a sparkline on one
worksheet made rust_xlsxwriter emit unbalanced `<ext>` elements, and Excel reported the
workbook as damaged. xlsxturbo refused the combination from 1.0.0 and **carried that guard
across two upstream releases without anyone filing an issue.** When it was finally reported
as [#185](https://github.com/jmcnamara/rust_xlsxwriter/issues/185) on 2026-08-15, jmcnamara
acknowledged it within two hours, fixed it the same evening and released 0.98.1 the next
morning. The whole cost of the workaround was the cost of not asking. He also said outright
in that thread: report anything else you hit here so it gets fixed.

What makes a report land, from the one that worked:

- **A reproducer that needs no Excel.** Counting `<ext>` opens against closes in the
  generated XML shows the defect in a `println!`, so the maintainer can see it without
  opening a workbook or trusting your description of what Excel said.
- **Controls in the same program.** Each feature alone was printed beside the broken pair.
  That is what establishes the defect is in the combination rather than in either feature,
  and it is the difference between a report and a complaint.
- **Say what you measured and stop there.** The report proposed the sibling-`<ext>`
  structure as what it took the intent to be, and said the maintainer would know better.
  Do not assert a cause in his code you have not read.

**Pin the defect in a test that fails when it is fixed**, if the workaround makes it
unreachable from Python. That was `tests/upstream_defect.rs` — driving rust_xlsxwriter
directly, asserting the bug was *still present*, with a control so a worse regression could
not be misread as the known one. It went red on the 0.98.1 bump exactly as designed, which
is what triggered the guard's removal in 1.1.2. Delete such a file with the workaround; its
whole job is to be the thing that notices.

**The second workaround is also gone, one day after it was filed:**
[#186](https://github.com/jmcnamara/rust_xlsxwriter/issues/186) (2026-08-16) —
`Workbook::define_name` panicked on a name whose local part is empty (`""` or `"Sheet1!"`),
at `defined_name.name.chars().next().unwrap()`, `workbook.rs:1578` in 0.98.1 — was fixed in
0.98.2 (2026-08-17), which returned
`ParameterError("Name '' cannot be empty in Excel")` — 0.99.0 says `Name cannot be blank`
instead, so do not quote that string from here. The screen in `apply_defined_names` is gone,
and `Cargo.toml` carries an exact floor rather than `0.98`, because the guard's absence is
what makes the floor load-bearing.

That check — whether the crate's message is as good as the guard's — is worth keeping as
the removal criterion, and here the answer had two halves. The crate reports the **local**
part, so it says `Name ''` for both `""` and `"Sheet1!"` and cannot identify which
`defined_names` key a caller got wrong. What saves it is the `map_err` already wrapping the
call, which puts the caller's own key back in front:
`Failed to define name 'Sheet1!': Parameter error: 'Name '' cannot be empty in Excel'.`
`test_empty_defined_name_error_names_the_offending_key` pins that half specifically, because
it comes from xlsxturbo and not from the crate.

An independent review of that draft before filing caught three over-claims and asked for
two controls, and every one of them was right: **the controls decide whether a report is
about the thing you say it is about.** `"Sheet1!1abc"` (invalid but non-empty local part →
proper error) is what turns "empty names panic" from an assertion into a measurement, and
`"!MyName"` separates an empty local part from an empty sheet qualifier. The cut claims were
"the one input that falls through to a panic" (never measured to be the only one) and a
statement about pyo3 panic handling made from memory rather than from a test. Review the
draft against the packet of what you actually ran, not against what you believe.

## Name validation is ours now — upstream has drawn its line

On 2026-08-21 the rust_xlsxwriter maintainer closed the validation half of
[#189](https://github.com/jmcnamara/rust_xlsxwriter/issues/189#issuecomment-5368521583): he cannot
replicate Excel's name rules maintainably, so the crate reverted to a simple rule set plus a "keep
names simple, test in Excel if in doubt" note, and said so directly — *"Other libraries that wrap
`rust_xlsxwriter` may need a stricter validation but I think that is up to them."* Duplicate-name
comparison got the same answer: no case folding, out of scope.

**So the crate's checks are a foot-gun guard, not a specification, and they are not going to move
closer to Excel.** The two screens below stop being temporary workarounds waiting for an upstream
fix and become the layer xlsxturbo owns. That is a change of intent, not of code — read the removal
criteria in their doc comments with this in mind.

**What exists today is asymmetric on purpose.** `reject_reference_shaped_name`
(`src/workbook.rs`) **refuses** a defined name Excel would read as a cell reference, because
silently renaming one leaves every formula pointing at a name the workbook no longer defines.
`sanitize_table_name` (`src/parse/tables.rs`) **rewrites** a table name instead, since nothing
references it. Preserve that split in anything new: rename what nobody points at, refuse what
somebody might.

**Taken: `Cargo.toml` pins `0.99.0` (2026-08-23), the release of every #189 fix.** The rename
`XlsxError::TableNameReused` -> `NameReused` cost nothing, as predicted — we name neither variant.
Three things the bump did change, and only the last needed code:

- The crate now refuses cell-reference-shaped *defined* names itself.
  `reject_reference_shaped_name` was kept anyway, on the criterion in its doc comment: the crate
  reports the unqualified name only (`Name error for 'Q1'` for both `"Q1"` and `"Sheet1!Q1"`),
  while ours names the `defined_names` key the caller wrote and suggests a replacement.
- Calls that used to succeed now raise — a `defined_names` key holding a character Excel forbids
  (`My-Name`, `Total$`), a logical constant, a reserved `_xlnm.*` name. Every one of them wrote a
  workbook Excel objects to, so the change is a fix, but it is still a behaviour change and the
  CHANGELOG says so.
- **The crate's message for an empty defined name moved** (`Name '' cannot be empty in Excel` ->
  `Name cannot be blank`), reddening the three tests that pin it. That is the standing cost of
  having deleted our own screen for #186: those assertions are pinned to a message no test here
  controls. Re-pin them on a bump; do not reinstate the screen.

**A collision between a table name and a defined name is now a pre-check, not a save failure.**
0.99.0 enforces Excel's rule that the two kinds must be unique against each other, but from inside
`Workbook::save` — and `save_workbook` maps everything from there to `FileFailure`, so the caller
was told their *file* had failed, in a message naming neither the sheet nor the option. Measured,
with the pre-check disabled: `FileError: Failed to save workbook to '...': Name 'Sales' has already
been used in this workbook.` The general shape is worth carrying: **a library that validates at
save time will have its errors classified by whatever the save layer assumes**, so every new
save-time rule upstream adds is a candidate for a pre-check here.

**`claimed_table_name` (`src/convert.rs`) owns the gate that decides whether a sheet claims a name
at all** — a style requested, a header row, at least one data row, `constant_memory` off. Both
pre-checks call it, because a pre-check that guessed the gate differently would refuse workbooks
that save cleanly. It sits beside `apply_worksheet_features`, the code it mirrors.

⚠ **`sanitize_table_name` silently mangled names Excel accepts — measured 2026-08-21 end to end
through the shipped 1.2.0 wheel, half fixed the same day.** It branches on `is_alphanumeric()`,
which is the predicate upstream abandoned for exactly this reason: a combining mark is not
alphanumeric, so it became `_`.

| `table_name=` | written into the workbook |
|---|---|
| `Verkäufe` in NFD (`Verka` + U+0308) | `Verka_ufe` |
| `ไม่` Thai, tone mark U+0E48 | `ไม_` |
| `हिन्दी` Hindi, virama U+094D | `हिन_दी` |
| `がくせい` in NFD (U+3099) | `か_くせい` |

Controls in the same run behaved: ASCII `Sales`, NFC `Verkäufe`, `日本語` and NFC `がくせい` all
survived byte for byte, and `Q1` became `_Q1` as designed. Excel accepts every one of the mangled
inputs, there was no warning, and NFD text reaches a caller routinely.

**The fix taken was NFC normalisation before the screen** — the smallest of the three candidates,
because it repairs the mangling without widening what is *accepted*. It closes the first and
fourth rows of that table and not the second or third: Thai and Hindi marks have no precomposed
form, so those names are still rewritten. `marks_without_a_composed_form_are_still_rewritten` pins
that, deliberately, so the remaining gap is a decision on record rather than a surprise — and so
that widening the allowlist later goes red instead of passing silently. **NFC and not NFKC**:
NFKC folds `U+FF21` FULLWIDTH A to ASCII `A`, which would turn a name Excel treats as distinct
into the cell reference `A1` and then into `_A1`. `compatibility_forms_are_not_folded` is that
control.

**Closing the rest means inverting the allowlist to a denylist, and that direction needs the
mirror audit** — not "does it still catch the bad names" but "what does it now accept that it used
to rewrite?". Upstream's own flip in that direction moved 962,590 code points from rejected to
accepted, which only an enumeration against a real validator could size. Do not take that step from
reading.

**The proptest cannot see this, and the reason is worth transplanting.**
`sanitized_table_names_are_always_valid` asserts
`sanitized.chars().all(|c| c.is_alphanumeric() || c == '_')` — the same predicate
`sanitize_table_name` branches on. It is the writer agreeing with its own reader, so it holds by
construction for every input `".*"` can generate and would keep holding if the predicate were
wrong in any direction. A property that restates the implementation is not a test of it. The
oracle a name check needs is external: Excel, or a fixed table of names measured against Excel.

**A crate bump will not fix this either**, because the mangling happens in xlsxturbo before
rust_xlsxwriter is called. `sanitize_table_name` is an allowlist, and therefore *stricter* than the
crate's new denylist for exactly the characters at issue — the crate would accept every input in
that table. Confirmed by the 0.99.0 bump, which changed none of it.

⚠ **The mirror of that gap: the allowlist must stay at least as WIDE as the crate's denylist, or a
name we rewrite is still refused.** A sanitized name goes straight to `Table::set_name`, so
anything the crate rejects that we do not repair becomes a hard error on a name the caller expected
us to fix. Two were found that way against 0.99.0 and are now screened —
Excel's logical constants, and a reference with trailing text (`R2D2`) — and the check is
mechanical: read `utility::check_name` in the crate version being pinned and account for each of
its rules. `.` and `?` and the invalid-character list already become `_`; a backslash likewise; the
reserved `_xlnm.*` names cannot survive, since the `.` is rewritten; the 255-character cap and the
empty name are handled by the sanitizer's own tail.

**If a stricter layer is built, do not derive the rules from Excel's documentation** — it is wrong
about several common characters, `?` and `€` among them. Excel's own name validator is scriptable
and answers in a fraction of a second per name, and it rejects names that a clean file load
accepts, so it is the better of the two oracles. The measured character surveys behind that claim
are in the #189 thread.
