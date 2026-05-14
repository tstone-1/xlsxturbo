use rust_xlsxwriter::Color;

/// Parse color string (hex #RRGGBB or named color) to u32
pub(crate) fn parse_color(color_str: &str) -> Result<u32, String> {
    let color = color_str.trim();
    if let Some(hex) = color.strip_prefix('#') {
        if hex.len() != 6 {
            return Err(format!(
                "Invalid hex color '{}': expected 6 characters after #, got {}",
                color,
                hex.len()
            ));
        }
        u32::from_str_radix(hex, 16).map_err(|_| format!("Invalid hex color: {}", color))
    } else {
        match color.to_lowercase().as_str() {
            "white" => Ok(0xFFFFFF),
            "black" => Ok(0x000000),
            "red" => Ok(0xFF0000),
            "green" => Ok(0x00FF00),
            "blue" => Ok(0x0000FF),
            "yellow" => Ok(0xFFFF00),
            "cyan" => Ok(0x00FFFF),
            "magenta" => Ok(0xFF00FF),
            "gray" | "grey" => Ok(0x808080),
            "silver" => Ok(0xC0C0C0),
            "orange" => Ok(0xFFA500),
            "purple" => Ok(0x800080),
            "navy" => Ok(0x000080),
            "teal" => Ok(0x008080),
            "maroon" => Ok(0x800000),
            _ => Err(format!("Unknown color: {}", color)),
        }
    }
}

/// Parse color string into a rust_xlsxwriter `Color` enum.
/// Wraps `parse_color` — used by features whose setters take `impl Into<Color>`
/// rather than a raw `u32` (shapes, charts, sparklines).
pub(crate) fn parse_color_enum(color_str: &str) -> Result<Color, String> {
    parse_color(color_str).map(Color::RGB)
}
