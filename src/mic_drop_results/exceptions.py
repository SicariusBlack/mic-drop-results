class DiscordAPIError(Exception):
    pass


class InvalidTokenError(DiscordAPIError):
    pass
