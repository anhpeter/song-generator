import json

from song_generator import SongGenerator

input_songs_path = "input/songs.json"
output_songs_path = "output/songs.pptx"
template_path = "template/template_1.pptx"

with open(input_songs_path, encoding="utf-8") as input_song_list_file:
    input_song_list = json.load(input_song_list_file)[:5]
    print(f"Number of songs: {len(input_song_list)}")
    print("Generating...")
    song_generator = SongGenerator(
        input_template_file_path=template_path,
        input_song_list=input_song_list,
        content_max_length=112,
    )
    song_generator.generate()
    song_generator.save(output_songs_path)
    print(f"Saved at {output_songs_path}")
