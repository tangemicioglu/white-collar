from .mail import MailAdapter, UnavailableMailAdapter
from .slides import SlidesAdapter, UnavailableSlidesAdapter
from .word import OoxmlWordAdapter, WordAdapter

__all__ = [
    "MailAdapter",
    "OoxmlWordAdapter",
    "SlidesAdapter",
    "UnavailableMailAdapter",
    "UnavailableSlidesAdapter",
    "WordAdapter",
]
