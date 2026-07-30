//! The public exception hierarchy, and the helpers that raise it.
//!
//! # Why the classes are built by calling `type()`
//!
//! Every class here has **two or three** bases: an `xlsxturbo`-specific one plus the
//! builtin the same failure raised before the hierarchy existed. That second base is
//! what makes this hierarchy shippable in a minor release -- an existing
//! `except ValueError` around a save keeps working. PyO3's `create_exception!` macro
//! takes a single base, so it cannot express that; calling the `type` metaclass with a
//! bases tuple can, and is `abi3` clean because it is a pure Python-level operation.
//!
//! # Why there is no fallback when the cell is empty
//!
//! [`register`] runs inside `#[pymodule]` initialisation, so the cell is populated
//! before any `#[pyfunction]` in this crate can be called -- there is no ordering in
//! which a raise site observes an empty cell. The raise helpers therefore `expect`
//! rather than degrading to the builtin exception. A silent degradation would be worse
//! than a loud failure: the call would raise `ValueError` instead of
//! `ConfigurationError`, which looks exactly like a correct legacy-compatible raise and
//! would pass every test that only checks the builtin base.
//!
//! See `docs/roadmap-1.0.md`, decision D6, for how this hierarchy was chosen and what
//! it deliberately does not distinguish.

use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyDict, PyTuple, PyType};

/// The name the classes report as their `__module__`.
///
/// The compiled extension is `xlsxturbo.xlsxturbo`, but `xlsxturbo/__init__.py`
/// re-exports these names, so `xlsxturbo.ConfigurationError` resolves. Pointing
/// `__module__` at the package rather than the extension makes `repr()` read the way
/// users import it, and makes the classes picklable through that path.
const PUBLIC_MODULE: &str = "xlsxturbo";

/// The exception classes, resolved once during module initialisation.
struct ErrorTypes {
    base: Py<PyType>,
    configuration: Py<PyType>,
    configuration_type: Py<PyType>,
    input_data: Py<PyType>,
    file: Py<PyType>,
    workbook_validation: Py<PyType>,
}

static ERRORS: PyOnceLock<ErrorTypes> = PyOnceLock::new();

/// Build one exception class by calling `type(name, bases, {...})`.
fn new_exception_class<'py>(
    py: Python<'py>,
    name: &str,
    bases: &[Bound<'py, PyType>],
    doc: &str,
) -> PyResult<Py<PyType>> {
    let namespace = PyDict::new(py);
    namespace.set_item("__doc__", doc)?;
    namespace.set_item("__module__", PUBLIC_MODULE)?;
    py.get_type::<PyType>()
        .call1((name, PyTuple::new(py, bases)?, namespace))?
        .cast_into::<PyType>()
        .map(Bound::unbind)
        .map_err(PyErr::from)
}

/// Create the exception hierarchy, add it to `module`, and cache it for the raise
/// helpers.
///
/// Called from `#[pymodule]` initialisation. Idempotent: a second call rebuilds the
/// module attributes but keeps the already-cached classes, so the identity a user
/// captured earlier stays valid.
pub(crate) fn register(module: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = module.py();

    let value_error = py.get_type::<pyo3::exceptions::PyValueError>();
    let type_error = py.get_type::<pyo3::exceptions::PyTypeError>();
    let os_error = py.get_type::<pyo3::exceptions::PyOSError>();

    let base = new_exception_class(
        py,
        "XlsxTurboError",
        &[py.get_type::<pyo3::exceptions::PyException>()],
        "Base class for every exception raised by xlsxturbo.\n\n\
         Catching this catches everything this library raises, and nothing else. Its \
         subclasses also inherit the builtin exception that the same failure raised \
         before this hierarchy existed, so `except ValueError` and `except TypeError` \
         keep working unchanged.",
    )?;
    let base_ty = base.bind(py).clone();

    let configuration = new_exception_class(
        py,
        "ConfigurationError",
        &[base_ty.clone(), value_error.clone()],
        "An option or argument has an invalid value.\n\n\
         Raised for an unknown option key, a value outside the accepted set (a bad \
         color, an unparseable cell reference, an unrecognised chart type), and for \
         failures inside the write pipeline that are traceable to an option. Also a \
         `ValueError`.",
    )?;
    let configuration_ty = configuration.bind(py).clone();

    let configuration_type = new_exception_class(
        py,
        "ConfigurationTypeError",
        &[base_ty.clone(), type_error],
        "An option or argument has the wrong type.\n\n\
         Raised when a value is the wrong Python type rather than the wrong value -- a \
         list where a dict is required, a number where a string is required, a bytes \
         path. Also a `TypeError`.",
    )?;

    let input_data = new_exception_class(
        py,
        "InputDataError",
        &[base_ty.clone(), value_error.clone()],
        "The object passed as data is not a supported DataFrame.\n\n\
         Raised when the argument is neither a pandas nor a polars DataFrame, or when \
         it is one but its columns cannot be read.\n\n\
         Also a `ValueError`, not a `TypeError`, because that is what this failure has \
         always raised. `TypeError` is arguably the better fit -- the argument is the \
         wrong kind of object, not a bad value -- but changing it would break existing \
         `except ValueError` handlers, so the builtin base follows history rather than \
         taste. Catch `InputDataError` to get the distinction without the ambiguity.",
    )?;

    let file = new_exception_class(
        py,
        "FileError",
        &[base_ty.clone(), os_error, value_error],
        "A filesystem read or write failed.\n\n\
         Raised when the output workbook cannot be written (missing directory, \
         permissions, no space) or when a CSV input cannot be opened. Also an \
         `OSError`, which is what makes `except OSError` work, and a `ValueError`, \
         which is what keeps pre-0.19 `except ValueError` handlers working. `errno`, \
         `strerror` and `filename` are always `None`; the path is in the message.",
    )?;

    let workbook_validation = new_exception_class(
        py,
        "WorkbookValidationError",
        &[configuration_ty],
        "The configuration is well-formed, but Excel does not permit it.\n\n\
         Raised for workbook-level rules rather than option syntax: two sheets \
         claiming the same table name, a sheet name Excel rejects. A subclass of \
         `ConfigurationError`, and so also a `ValueError`.",
    )?;

    // `get_or_init` rather than `set`: re-importing the extension must not invalidate a
    // class object a caller already holds a reference to. On a second import the classes
    // built above are discarded and the module gets the cached ones, so that `isinstance`
    // against a module attribute always agrees with what the raise helpers construct.
    let cached = ERRORS.get_or_init(py, || ErrorTypes {
        base,
        configuration,
        configuration_type,
        input_data,
        file,
        workbook_validation,
    });

    for (name, class) in [
        ("XlsxTurboError", &cached.base),
        ("ConfigurationError", &cached.configuration),
        ("ConfigurationTypeError", &cached.configuration_type),
        ("InputDataError", &cached.input_data),
        ("FileError", &cached.file),
        ("WorkbookValidationError", &cached.workbook_validation),
    ] {
        module.add(name, class.bind(py).clone())?;
    }

    Ok(())
}

/// Build a `PyErr` of the class selected by `pick`.
fn raise(pick: fn(&ErrorTypes) -> &Py<PyType>, message: String) -> PyErr {
    Python::attach(|py| {
        let types = ERRORS
            .get(py)
            .expect("errors::register runs during module initialisation");
        PyErr::from_type(pick(types).bind(py).clone(), (message,))
    })
}

/// An option or argument value is invalid. Also a `ValueError`.
pub(crate) fn configuration(message: impl Into<String>) -> PyErr {
    raise(|t| &t.configuration, message.into())
}

/// An option or argument has the wrong type. Also a `TypeError`.
pub(crate) fn configuration_type(message: impl Into<String>) -> PyErr {
    raise(|t| &t.configuration_type, message.into())
}

/// The object passed as data is not a supported DataFrame. Also a `ValueError` --
/// deliberately not a `TypeError`; see the class docstring in [`register`] for why.
pub(crate) fn input_data(message: impl Into<String>) -> PyErr {
    raise(|t| &t.input_data, message.into())
}

/// A filesystem read or write failed. Also an `OSError` and a `ValueError`.
pub(crate) fn file(message: impl Into<String>) -> PyErr {
    raise(|t| &t.file, message.into())
}

/// Well-formed configuration that Excel does not permit. Also a `ValueError`.
pub(crate) fn workbook_validation(message: impl Into<String>) -> PyErr {
    raise(|t| &t.workbook_validation, message.into())
}

impl From<crate::convert::ConvertError> for PyErr {
    /// Routes the pipeline's one internal distinction onto the public hierarchy.
    ///
    /// This is what makes the `?` operator enough at the `lib.rs` boundary: the
    /// pipeline already decided whether a failure was the filesystem, so the boundary
    /// does not have to guess from the message text.
    fn from(error: crate::convert::ConvertError) -> Self {
        match error {
            crate::convert::ConvertError::Config(message) => configuration(message),
            crate::convert::ConvertError::File(message) => file(message),
        }
    }
}
