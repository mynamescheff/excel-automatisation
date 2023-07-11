import re

character_mapping = {
    'á': 'a',
    'à': 'a',
    'â': 'a',
    'ä': 'a',
    'ã': 'a',
    'å': 'a',
    'æ': 'ae',
    'ç': 'c',
    'č': 'c',
    'é': 'e',
    'è': 'e',
    'ê': 'e',
    'ë': 'e',
    'í': 'i',
    'ì': 'i',
    'î': 'i',
    'ï': 'i',
    'ñ': 'n',
    'ó': 'o',
    'ò': 'o',
    'ô': 'o',
    'ö': 'o',
    'õ': 'o',
    'ø': 'o',
    'œ': 'oe',
    'š': 's',
    'ú': 'u',
    'ù': 'u',
    'û': 'u',
    'ü': 'u',
    'ý': 'y',
    'ÿ': 'y',
    'ž': 'z'
    # Add more mappings as needed
}

def transform_to_swift_accepted_characters(input_list):
    transformed_list = []
    for input_string in input_list:
        transformed_string = re.sub(r'\b\w+\b', lambda m: ''.join(character_mapping.get(char, char) for char in m.group()), str(input_string))
        transformed_string = re.sub(r'[.,]', '', transformed_string)  # Remove dots and commas
        transformed_list.append(transformed_string)
    return transformed_list