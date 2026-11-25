# -*- coding: utf-8 -*-
"""Feature registry. Order = order of application inside one profile."""
from .background import BackgroundFeature
from .base import Context, Feature, ProfileError  # noqa: F401
from .comments import CommentsFeature
from .cover import CoverFeature
from .font import FontFeature
from .header_footer import HeaderFooterFeature
from .illustrations import IllustrationsFeature
from .lists import ListsFeature
from .page import BreaksFeature, PageFeature
from .paragraph import ParagraphFeature
from .proofing import ProofingFeature
from .tables import TablesFeature
from .text import TextFeature

# Structure/removal first, text edits next, then formatting, page-level items; cover last
# (so formatting passes don't overwrite the cover's own styling).
_ORDER = [CommentsFeature, TextFeature, BreaksFeature, ParagraphFeature, FontFeature,
          ListsFeature, TablesFeature, IllustrationsFeature, PageFeature, HeaderFooterFeature, BackgroundFeature,
          ProofingFeature, CoverFeature]

REGISTRY = {f.key: f() for f in _ORDER}


def get_feature(key):
    try:
        return REGISTRY[key]
    except KeyError:
        raise ProfileError("unknown feature %r (available: %s)" % (key, ", ".join(REGISTRY)))
