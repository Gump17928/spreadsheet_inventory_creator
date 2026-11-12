"""Small helpers for turning Python values into sensible cell contents."""


def number_if_plain(text):
    """'2026' -> 2026 and '3.12' -> 3.12 so Excel shows no "number stored as text" warning.

    Anything that would not survive the round trip ('3.10' would become 3.1) or is not a number
    ('2025 R2') stays text; the workbook's version check handles both.
    """
    try:
        number = float(text)
    except (TypeError, ValueError):
        return text
    converted = int(number) if number.is_integer() else number
    return converted if str(converted) == str(text) else text
