from pptx import Presentation

class DataLoader:
    #Класс для загрузки данных из PPTX файлов.
    def load_presentation(self, path):
        #Загружает PPTX презентацию.
        return Presentation(path)