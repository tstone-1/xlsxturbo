# The exception hierarchy

Moved verbatim from the repository-root `AGENTS.md`, which keeps a one-line index entry for each section. Read the section before working in its area.

## The Exception Hierarchy (0.19.0+)

`src/errors.rs` owns the public exception classes and is the only place in `src/` that may
construct a `PyErr` from scratch. Raise with `errors::configuration`,
`errors::configuration_type`, `errors::input_data`, `errors::file` or
`errors::workbook_validation`. The classes are built by calling the `type` metaclass rather
than with `create_exception!`, because that macro takes a single base and every class here
needs two or three.

Facts that are expensive to rediscover:

- **`OptionError` is never raised, and that is not an oversight.** It exists so
  `except OptionError` catches both configuration classes and nothing else. The guard that
  every exported class needs a working trigger still applies to it, in a different form:
  `ABSTRACT` in `tests/test_errors.py` maps it to exactly the triggered classes it must
  catch, checked in both directions. Declaring a class abstract is otherwise a way to walk
  straight past the rule that killed `UnsupportedFeatureError`.
- **`OptionError` must not take a builtin base.** A builtin there lands on *both* children,
  so a `ConfigurationTypeError` would silently also be a `ValueError` and the value/type
  split would stop meaning anything to `except`. `FORBIDDEN_BASES` pins it.
- **The second base is a compatibility contract, not decoration.** Each class inherits the
  builtin its failures raised in 0.18. Pick a new class's builtin by **grepping what the site
  raises today**, not by what the failure morally is. Getting this from taste produced a
  breaking change twice during the original implementation — once for I/O (`OSError` alone
  would have broken `except ValueError`) and once for `InputDataError` (which is a
  `ValueError`, because frame detection has always reached the boundary through the pipeline).
- **The 93 `pytest.raises(ValueError|TypeError)` assertions across `tests/` are the
  behaviour record for pre-0.19.** They are the compatibility gate. A change to this area that
  needs one of them edited to stay green is a breaking change, whatever the changelog claims.
- **A class no site raises is dead API that can never be removed.**
  `tests/test_errors.py` asserts the exported set equals the set with a working trigger, which
  is what kept `UnsupportedFeatureError` from shipping (a `constant_memory` conflict is a
  `RuntimeWarning` and the call succeeds — nothing would have raised it).
- **The boundary is `src/lib.rs` *and* `src/extract.rs`,** 50 sites, not the dozen the plan
  assumed. Everything below `extract.rs` is uniformly `Result<_, String>`.
- **`ConvertError` in `src/convert.rs` is the one seam, and it has no `From<String>`.** Two
  variants, `Config` and `File`, because `save_workbook` runs inside the pipeline and its
  failure would otherwise be indistinguishable from a bad option at the boundary.

  Until 0.21.0 a blanket `From<String>` mapped untagged failures to `Config`, so `?` compiled
  everywhere and **a new filesystem call silently blamed the caller's options.** That
  conversion is gone: every site now names its variant, and a new failure site *does not
  compile* until it chooses. Prefer that over the `Internal` fallback variant the review
  proposed — a fallback that still exists is still a default, and the default was the bug.
  `TestConvertErrorHasNoDefaultCategory` watches for it coming back, with a control asserting
  both variants are still constructed.

  Frame detection needed the same treatment and got a boundary call
  (`require_supported_dataframe`) instead, since tagging it would have pushed `ConvertError`
  down into the write layer.
- **`FileError.errno` is populated; `strerror` and `filename` must stay `None`.** Setting
  `filename` makes `OSError.__str__` switch to `[Errno n] strerror: 'filename'` and **discard
  the message**, which is where this library puts the path and the context. `errno` alone
  leaves `str()` untouched. On Windows `raw_os_error()` is a Win32 code, not an errno —
  `ERROR_PATH_NOT_FOUND` is 3, which as POSIX means "no such process" — so `errors::posix_errno`
  passes the number through on Unix and classifies via `io::ErrorKind` elsewhere.

Full reasoning, including the shapes considered and rejected: `docs/roadmap-1.0.md`
decision D6. User-facing contract: `docs/errors.md`.
