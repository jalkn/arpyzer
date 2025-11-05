from django import template

register = template.Library()

@register.filter
def split_lines(value):
    """Splits a string by newlines and returns a list of lines."""
    if isinstance(value, str):
        # Only split if it's a string; handle None/empty string gracefully
        return value.splitlines()
    return [] # Return an empty list for non-string or None values
