from .text_clean import clean_text
from .file_io import read_text, write_text, safe_mkdir, list_files
from .placeholder import (
    extract_block, append_below_placeholder, replace_placeholder, has_placeholder,
    extract_serviceinfo_block, update_serviceinfo_block,
)

# NEW — merged OpenLP module
from .openlp import (
    load_song,
    load_custom_slide,
    get_scripture_text,
    build_song_index,
)

# from .ccli_helpers import get_ccli_lyrics  # deprecated and archived
from .logging_utils import init_logging
