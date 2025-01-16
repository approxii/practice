
## Word

### `/generate/`
- **Описание**: Принимает документ и словарь в теле запроса, возвращает новый документ.
  - **Допустимые входные данные**:
  
    file: "<binary>"  // Word файл  
      ```json
    {
        "blocks":[
            {
                "key1": "значение1",
                "key2": "значение2"
            },
            {
                "key1": "другое значение 1"
                "key2": "другое значение 2"
            }
        ],
        "newpage": "true"
    }


### `/get_bookmarks/`
- **Описание**: Принимает документ в теле запроса, возвращает словарь в Response body
- **Допустимые входные данные**:
  
  Word файл  

  - **Формат ответа**:
      ```json
      {
          "blocks":[
          {
              "key1": "значение1",
              "key2": "значение2"
          }
      ],
      "newpage": "false"
      }


### `/get_bookmarks_with_formatting/`
- **Описание**: Принимает документ в теле запроса, возвращает словарь в Response body, учитывая форматирование документа
- **Допустимые входные данные**:
  
  Word файл

  - **Формат ответа**:
    ```json
    {
          "blocks":[
          {
              "key1": {
                "value": "значение1",
                "format": {
                  "fontname": "Times New Roman",
                  "fontsize": 14,
                  "fillcolor": "yellow",
                  "textcolor": null,
                  "bold": true,
                  "italic": true,
                  "underline": true,
                  "strikethrough": false,
                  "align": "center"
                }
              }
          }
      ],
      "newpage": "false"
      }


### `/generate_with_formatting/`
- **Описание**: Принимает документ и словарь в теле запроса, возвращает новый документ. Учитывается форматирование
  - **Допустимые входные данные**:
  
    file: "<binary>"  // Word файл  
      ```json
    {
        "blocks": [
        {
            "key1": {
              "value": "измененное с форматированием 1",
              "format": {
                "fontname": "Times New Roman",
                "fontsize": 24,
                "fillcolor": "yellow",
                "textcolor": "FF0000",
                "bold": true,
                "italic": false,
                "underline": true,
                "strikethrough": true,
                "align": "center"
              }
            },
            "key2": {
              "value": "измененное с форматирование 2",
              "format": {
                "fontname": "Calibri",
                "fontsize": "11",
                "fillcolor": null,
                "textcolor": null,
                "bold": false,
                "italic": false,
                "underline": false,
                "strikethrough": false,
                "align": "center"
              }
            }
          }
        ],
        "newpage": "false"
    }


### `/add bookmarks instead of paragraph/`
- **Описание**: Принимает документ и словарь в теле запроса, возвращает новый документ(очищает нужный текст, указанный в словаре, и ставит туда закладку)
  - **Допустимые входные данные**:

    file: "<binary>"  // Word файл
      ```json
    {
        "bookmarks": [
        {
            "текст для замены 1": "key1"
            "текст для замены 2": "key2"
        }
        ]
    }

