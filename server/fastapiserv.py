import os
from src.loader import DataLoader
from src.updater import parser
from flask import Flask
from flask_restx import Api, Resource
from werkzeug.datastructures import FileStorage

# Задать пути к папкам
current_dir = os.path.dirname(os.path.abspath(__file__))
data_dir = os.path.join(current_dir, '..', 'data')
output_dir = os.path.join(current_dir, '..', 'output')
pptx_path = os.path.join(data_dir, 'a.pptx')

#Инициализировать Flask
app = Flask(__name__)
api = Api(app, version='1.0', title='PPTX Analyze API',
          description='Анализ PPTX презентации с помощью JSON')

#Настройка Swagger модели
pptx_upload = api.parser()
pptx_upload.add_argument('pptx_file', location='files', type=FileStorage, required=True, help='PPTX файл')

ns = api.namespace('pptx', description='Анализ PPTX презентации')

@ns.route('/post')
class PptxUpdate(Resource):
    @api.expect(pptx_upload)
    def post(self):
        args = pptx_upload.parse_args()
        pptx_file = args['pptx_file']  # Получаем файл PPTX

        #Сохраняем PPTX файл
        pptx_file.save(pptx_path)

        try:
            if main_update(pptx_path) == True:
                return 'Success', 200
            else:
                return 'Failed', 400
        except Exception as e:
            return f'{e}', 500

def main_update(pptx_path):
    #Основная функция для анализа презентации.
    loader = DataLoader()
    updater = parser()

    try:
        # Загрузить презентацию
        pptx = loader.load_presentation(pptx_path)
        # Заполнить JSON
        analyzed_json = updater.analyze(pptx)
        print("Презентация успешно проанализирована.")
        return True
    except Exception as e:
        print(f"Ошибка: {e}")
        return False

# Точка входа
#if __name__ == '__main__':
    #app.run(debug=True)