import copy
import math

from pptx import Presentation


class SongGenerator:

    def __init__(
        self,
        input_template_file_path: str,
        input_song_list: list,
        content_max_length=28,
    ):
        self.content_max_length = content_max_length
        self.input_template_file_path = input_template_file_path
        self.input_song_list = input_song_list
        self.template_presentation = Presentation(input_template_file_path)
        self.presentation = Presentation(input_template_file_path)

    def duplicate_slide(self, pres, index):
        template = pres.slides[index]
        try:
            blank_slide_layout = pres.slide_layouts[index]
        except:
            blank_slide_layout = pres.slide_layouts[len(pres.slide_layouts)]

        copied_slide = pres.slides.add_slide(blank_slide_layout)

        for shp in template.shapes:
            el = shp.element
            newel = copy.deepcopy(el)
            copied_slide.shapes._spTree.insert_element_before(newel, "p:extLs")

        # remove empty shape
        empty_shape_idx_list = []
        for idx, shape in enumerate(copied_slide.shapes):
            if shape.text.strip() == "":
                empty_shape_idx_list.append(idx)

        deleted_shape_count = 0
        for idx in empty_shape_idx_list:
            copied_slide.shapes.element.remove(
                copied_slide.shapes[idx - deleted_shape_count].element
            )
            deleted_shape_count += 1
        return copied_slide

    def update_placeholder_content(self, slide, placeholder, content):
        for shape in slide.shapes:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.text == f"<{placeholder}>":
                        run.text = content

    def new_slide(self):
        slide_layout = self.presentation.slide_layouts[0]
        slide = self.presentation.slides.add_slide(slide_layout)
        for shape in slide.shapes:
            slide.shapes._spTree.remove(shape._element)
        return slide

    def generate(self):
        for input_song in self.input_song_list:
            title_slide = self.duplicate_slide(self.presentation, 0)
            self.update_placeholder_content(title_slide, "type", input_song["type"])
            self.update_placeholder_content(title_slide, "title", input_song["title"])
            for input_content in input_song["content"]:
                input_words = input_content.split(" ")
                number_of_slide = math.ceil(len(input_words) / self.content_max_length)
                for i in range(number_of_slide):
                    slide_content = " ".join(
                        input_words[
                            i
                            * self.content_max_length : (i + 1)
                            * self.content_max_length
                        ]
                    )
                    print(f"content length: {len(slide_content)}")
                    content_slide = self.duplicate_slide(self.presentation, 1)
                    self.update_placeholder_content(
                        content_slide, "content", slide_content
                    )

    def save(self, path):
        self.presentation.save(path)
