import re


def to_snake_case(string):
    changed_string = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', string)
    result = re.sub('([a-z0-9])([A-Z])', r'\1_\2', changed_string).lower()
    return result