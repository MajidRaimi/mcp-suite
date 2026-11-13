class WhatsAppError(Exception):
    pass


class ConfigurationError(WhatsAppError):
    pass


class WahaUnavailableError(WhatsAppError):
    pass


class ChatNotFoundError(WhatsAppError):
    pass


class DraftError(WhatsAppError):
    pass


class AttachmentError(WhatsAppError):
    pass


class PartialSendError(WhatsAppError):
    pass
