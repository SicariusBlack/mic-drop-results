import re


# Section A: Find all
match_url = re.compile(
    r"https?:\/\/(?:www\.)?[-\w@:%.\+~#=]{1,256}\.[\w()]{1,6}"
    r"\b(?:[-\w()@:%.\+~#?&\/=]*)",
    re.IGNORECASE,
)
match_field_name = re.compile(
    r"(?<={)([\w \-]*?)(?=})"  # "field_name" from "{field_name}"
)
match_windows_username = re.compile(
    r"(?<=(?:\\|\/)Users(?:\\|\/)).+?(?=(?:\\|\/))",  # "username" from "C:\Users\username\"
    re.IGNORECASE,
)


# Section B: Substitute
match_forbidden_char = re.compile(r'[\\\/:"*?<>|]+')
match_space = re.compile(r"\s")
match_non_username_char = re.compile(r"([^a-z0-9_.]|(?<=\.)\.|(?<=#).+)", re.IGNORECASE)


# Section C: Full match
match_hex = re.compile(r"^#?[0-9a-f]{6}$", re.IGNORECASE)
