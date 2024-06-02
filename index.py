from pptx.util import Pt

from song_editor import SongEditor

input_songs_path = "input/Songs.pptx"
output_songs_path = "output/Songs.pptx"

song_generator = SongEditor(
    input_path=input_songs_path,
)
song_generator.update_title_font_size(Pt(48), Pt(80))
song_generator.save(output_songs_path)
