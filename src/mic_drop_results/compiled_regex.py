import re


# Section A: Find all
# Match valid URLs
url_pattern = re.compile(
    r'https?:\/\/(?:www\.)?[-\w@:%.\+~#=]{1,256}\.[\w()]{1,6}'
    r'\b(?:[-\w()@:%.\+~#?&\/=]*)',
    re.IGNORECASE)

# Match 'field_name' from 'text containing {field_name} with more text behind'
field_name_pattern = re.compile(r'(?<={)([\w \-]*?)(?=})')

# Match 'username' from 'C:\Users\username\directory'
username_pattern = re.compile(
    r'(?<=(?:\\|\/)Users(?:\\|\/)).+?(?=(?:\\|\/))',
    re.IGNORECASE)


# Section B: Substitute
# Match forbidden file name characters
forbidden_char_pattern = re.compile(r'[\\\/:"*?<>|]+')

# Match space characters
space_pattern = re.compile(r'\s')

# Match special characters
special_char_pattern = re.compile(r'[^a-z0-9]', re.IGNORECASE)


# Section C: Full match
# Match valid hex color code
hex_pattern = re.compile(r'^#?[0-9a-f]{6}$', re.IGNORECASE)