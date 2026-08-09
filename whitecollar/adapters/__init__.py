from .protocols import MailAdapter, SlidesAdapter, WordAdapter
from .outlook import OutlookComAdapter
from .slides import PowerPointComAdapter
from .word_com import Win32WordComAdapter

__all__ = [
    "MailAdapter",
    "OutlookComAdapter",
    "PowerPointComAdapter",
    "SlidesAdapter",
    "WordAdapter",
    "Win32WordComAdapter",
]
