from .mail import MailAdapter, UnavailableMailAdapter
from .slides import PowerPointComAdapter, SlidesAdapter, UnavailableSlidesAdapter
from .word import OoxmlWordAdapter, WordAdapter
from .word_com import Win32WordComAdapter

__all__ = [
    "MailAdapter",
    "OoxmlWordAdapter",
    "SlidesAdapter",
    "UnavailableMailAdapter",
    "UnavailableSlidesAdapter",
    "WordAdapter",
    "Win32WordComAdapter",
]
