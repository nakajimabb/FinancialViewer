import re

def safe_int(s):
    try:
        return int(s)
    except (ValueError, TypeError):
        return 0

def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def natural_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]

def parse_int_or_str(s):
    try:
        return int(s)
    except ValueError:
        return s

def safe_get(lst, index):
    return lst[index] if -len(lst) <= index < len(lst) else None
