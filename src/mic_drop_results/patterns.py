import re


# Match image URLs
img_url_pattern = re.compile(
    r'(https?\:\/\/)?[\w\-\.]+\.[a-z]*\/'
    r'\S*\.(png|jpg|jpeg|gif|svg)',
    re.IGNORECASE)

# Match 'field_name' from 'text containing {field_name} with more text behind'
field_name_pattern = re.compile(r'(?<={)([\w \-]*?)(?=})')

# Match space characters
space_pattern = re.compile(r'\s')

# Match 'username' from 'C:\Users\username\directory'
username_pattern = re.compile(
    r'(?<=(?:\\|\/)Users(?:\\|\/)).+?(?=(?:\\|\/))',
    re.IGNORECASE)