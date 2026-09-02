//! Attaching the caller's option to an error raised by a pure parser.
//!
//! The parsers in this module take a value, not an option name: `parse_color`
//! is given `"nope"` and cannot know whether it came from `header_format`, a
//! `conditional_formats` entry or a textbox font. That is the right shape for
//! them, and it makes naming the option the caller's job.
//!
//! Most callers did not do it. Measured against 1.3.0, six of six sampled
//! failures reached Python with no option and no key --
//! `Unknown criteria 'bogus'. Valid: ...` for a `conditional_formats` dict
//! holding two entries, `Unknown color: nope` for a `header_format`. The
//! convention `AGENTS.md` states for this layer is
//! `<feature>['<key>']: <message>`, and these two combinators are how a call
//! site meets it in one line.
//!
//! Nothing here decides an error's *class*: everything below `extract.rs`
//! returns `Result<_, String>` and the boundary classifies it. These only
//! change the text, which `docs/stability.md` explicitly does not cover.

/// Prefix an error from a pure parser with the option that produced it.
pub(crate) trait WithOptionContext<T> {
    /// Prefix with an option context, e.g. `charts['D2']: <message>`.
    ///
    /// Use where the failing value *is* the option, so a key would only
    /// repeat what the message already says.
    fn in_option(self, context: &str) -> Result<T, String>;

    /// Prefix with an option context and the key inside it, e.g.
    /// `header_format: 'bg_color': <message>`.
    ///
    /// Use where the option is a dict and several of its keys can fail the
    /// same way -- `bg_color`, `font_color` and `border_color` all reject a
    /// bad colour with the same sentence, so without the key the caller is
    /// told a colour is wrong but not which one.
    fn in_field(self, context: &str, key: &str) -> Result<T, String>;
}

impl<T> WithOptionContext<T> for Result<T, String> {
    fn in_option(self, context: &str) -> Result<T, String> {
        self.map_err(|e| format!("{}: {}", context, e))
    }

    fn in_field(self, context: &str, key: &str) -> Result<T, String> {
        self.map_err(|e| format!("{}: '{}': {}", context, key, e))
    }
}

#[cfg(test)]
mod tests {
    use super::WithOptionContext;

    #[test]
    fn an_ok_result_passes_through_untouched() {
        // The control: a combinator that rewrote successes would be caught
        // here rather than by a surprised caller.
        let ok: Result<u8, String> = Ok(7);
        assert_eq!(ok.in_option("charts['D2']"), Ok(7));
        let ok: Result<u8, String> = Ok(7);
        assert_eq!(ok.in_field("header_format", "bg_color"), Ok(7));
    }

    #[test]
    fn the_option_and_key_are_prepended_in_the_documented_shape() {
        let err: Result<(), String> = Err("Unknown color: nope".to_string());
        assert_eq!(
            err.in_option("charts['D2']"),
            Err("charts['D2']: Unknown color: nope".to_string())
        );
        let err: Result<(), String> = Err("Unknown color: nope".to_string());
        assert_eq!(
            err.in_field("header_format", "bg_color"),
            Err("header_format: 'bg_color': Unknown color: nope".to_string())
        );
    }
}
