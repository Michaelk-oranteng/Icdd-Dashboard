# control_dashboard/templatetags/custom_filters.py

from django import template
import json

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """Get an item from a dictionary using a key."""
    if dictionary is None:
        return ''
    if isinstance(dictionary, dict):
        return dictionary.get(key, '')
    return ''

@register.filter
def default(value, default_value=''):
    """Return default value if value is None or empty."""
    if value is None or value == '':
        return default_value
    return value

@register.filter
def to_json(value):
    """Convert a value to JSON string."""
    try:
        return json.dumps(value, default=str)
    except:
        return str(value)

@register.filter
def get_type(value):
    """Get the type of a value as a string."""
    if value is None:
        return 'None'
    return type(value).__name__

@register.filter
def get_display_date(value):
    """Format a date for display."""
    if not value:
        return ''
    from datetime import datetime
    try:
        if isinstance(value, str):
            # Try to parse the string
            value = datetime.strptime(value, '%Y-%m-%d').date()
        if hasattr(value, 'strftime'):
            return value.strftime('%b %d, %Y')
    except:
        pass
    return str(value)