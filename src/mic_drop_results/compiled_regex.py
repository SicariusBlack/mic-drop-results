import re


# Section A: Find all
# Match image URLs
img_url_pattern = re.compile(
    r'(https?\:\/\/)?[\w\-\.]+\.[a-z]*\/'
    r'\S*\.(png|jpg|jpeg|gif|svg)',
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


# Section C: Full match
# Match hex color value
hex_pattern = re.compile(r'^[0-9a-f]{6}$', re.IGNORECASE)
