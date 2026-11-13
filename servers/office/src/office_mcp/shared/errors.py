class OfficeError(Exception):
    pass


class SettingsError(OfficeError):
    pass


class PlatformError(OfficeError):
    pass
