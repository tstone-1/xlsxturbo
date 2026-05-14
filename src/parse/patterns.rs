/// Check if a column name matches a wildcard pattern.
/// Supports: "prefix*", "*suffix", "*contains*", or exact match
pub(crate) fn matches_pattern(column_name: &str, pattern: &str) -> bool {
    let starts_with_star = pattern.starts_with('*');
    let ends_with_star = pattern.ends_with('*');

    match (starts_with_star, ends_with_star) {
        (true, true) => {
            // *contains* - match substring; lone "*" matches everything
            if pattern.len() <= 2 {
                return true;
            }
            let inner = &pattern[1..pattern.len() - 1];
            column_name.contains(inner)
        }
        (true, false) => {
            // *suffix - match ending
            let suffix = &pattern[1..];
            column_name.ends_with(suffix)
        }
        (false, true) => {
            // prefix* - match beginning
            let prefix = &pattern[..pattern.len() - 1];
            column_name.starts_with(prefix)
        }
        (false, false) => {
            // Exact match
            column_name == pattern
        }
    }
}
