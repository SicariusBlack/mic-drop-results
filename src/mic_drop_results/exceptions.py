class DiscordAPIError(Exception):
    pass


class InvalidTokenError(DiscordAPIError):
    pass


class UnknownUserError(DiscordAPIError):
    pass
