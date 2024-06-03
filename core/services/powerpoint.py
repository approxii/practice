from io import BytesIO
from pptx import Presentation
from core.services.base import BaseDocumentService
from pptx.util import Inches

class PPTXService(BaseDocumentService):
    def __init__(self):
        self.presentation = None

    def load(self, file) -> None:
        self.presentation = Presentation(file)

    def update(self, data: dict) -> None:
        if not self.presentation:
            raise ValueError("PPTX файл не загружен")
        
        for key, value in data.items():
            if key.startswith('a') and key[1:].isdigit():
                slide_index = int(key[1:]) - 1
                # Убедиться что слайд есть, иначе создать
                if slide_index >= len(self.presentation.slides):
                    slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[6])  # Использовать пустой шаблон
                else:
                    slide = self.presentation.slides[slide_index]

                # Создать и заполнить фигуру в зависимости от значения
                if isinstance(value, list):
                    for index, item in enumerate(value):
                        if isinstance(item, list):
                            text = '\n'.join(f"{i+1}. {sub_item}" for i, sub_item in enumerate(item))
                            self.create_shape(slide, text, index+1)
                        else:
                            self.create_shape(slide, item, index+1)
                else:
                    self.create_shape(slide, value, index+1)

    def create_shape(self, slide, text, index):
        """Создать фигуру с текстом на слайде"""
        left = width = height = Inches(1)
        textbox = slide.shapes.add_textbox(left, Inches(index), width, height)
        tf = textbox.text_frame
        p = tf.add_paragraph()
        p.text = text

    def save_to_bytes(self) -> BytesIO:
        if not self.presentation:
            raise ValueError("PPTX файл не загружен")
        output = BytesIO()
        self.presentation.save(output)
        output.seek(0)
        return output

    def save_to_file(self, file_path: str) -> None:
        if self.presentation:
            self.presentation.save(file_path)
        else:
            raise ValueError("PPTX файл не загружен")
