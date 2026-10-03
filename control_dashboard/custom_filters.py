# control_dashboard/templatetags/custom_filters.py

from django import template

register = template.Library()


@register.filter
def get_item(dictionary, key):
    """Exact-key lookup (existing behavior — do not change)."""
    if not isinstance(dictionary, dict):
        return None
    return dictionary.get(key)


@register.filter
def get_item_loose(dictionary, key):
    """
    Robust lookup that tries several normalizations of `key`
    against the dict, so headers and row keys don't need to match
    byte-for-byte.

    Tries, in order:
      1. exact key
      2. key.upper() / key.lower() / key.title()
      3. spaces → underscores (and back)
      4. all-whitespace collapsed
      5. curly quotes → straight quotes
      6. any key in the dict that normalizes to the same string
    """
    if not isinstance(dictionary, dict):
        return None

    def _norm(s):
        return (
            str(s or '')
            .replace('\u2019', "'").replace('\u2018', "'")
            .replace('\u201c', '"').replace('\u201d', '"')
            .replace('\u00a0', ' ')
        )

    # 1. exact
    if key in dictionary:
        return dictionary[key]

    key_str = str(key or '')

    # 2. case variants
    for variant in (key_str.upper(), key_str.lower(), key_str.title()):
        if variant in dictionary:
            return dictionary[variant]

    # 3. underscore / space variants
    for variant in (
        key_str.replace(' ', '_'),
        key_str.replace('_', ' '),
        key_str.replace(' ', '_').lower(),
        key_str.replace(' ', '_').upper(),
        key_str.replace(' ', '').lower(),
        key_str.replace(' ', '').upper(),
    ):
        if variant in dictionary:
            return dictionary[variant]

    # 4. normalize whitespace + curly quotes, then compare
    key_norm = _norm(key_str).strip().upper()
    for dict_key, dict_val in dictionary.items():
        if _norm(dict_key).strip().upper() == key_norm:
            return dict_val

    # 5. final fallback — compare with all whitespace stripped
    key_stripped = key_norm.replace(' ', '')
    for dict_key, dict_val in dictionary.items():
        if _norm(dict_key).strip().upper().replace(' ', '') == key_stripped:
            return dict_val

    return None