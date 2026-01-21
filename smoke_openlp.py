from scripts.utils.openlp import get_songs_db, get_bible_db
import traceback

for church in ('elkton','lb'):
    try:
        print(f'{church} songs_db: {get_songs_db(church)}')
        print(f'{church} bible_db: {get_bible_db(church)}')
    except Exception:
        traceback.print_exc()
