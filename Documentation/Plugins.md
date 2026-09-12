# Плагины и расширенные функции Echos

---

## Введение

Плагины позволяют расширять функциональность Echos без изменения основного кода. Плагин — это Python-скрипт, который загружается при запуске программы и может обрабатывать сообщения, получать доступ к API и выполнять дополнительные действия.

---

## Где находятся плагины

Плагины размещаются в папке `plugins/` рядом с исполняемым файлом (или в каталоге программы). Открыть папку можно через **Настройки → Плагины → Открыть папку плагинов**.

---

## Создание первого плагина


Создайте файл с расширением `.py` в папке `plugins`. Минимальный шаблон:

```python
class Plugin:
    def on_message_send(self, text, api):
        # Вызывается перед отправкой сообщения пользователя
        # Пример: меняем temperature для каждого запроса
        api.main_window.settings["temperature"] = 0.3
        api.main_window.save_settings()
        return text

    def on_response_received(self, text, api):
        # Вызывается после получения ответа модели
        return text
```

Сохраните файл. Плагин загрузится при следующем запуске программы или сразу после нажатия Настройки → Плагины → Перезагрузить (кнопка в диалоге управления плагинами).

---

## Доступные хуки (события)

| Хук | Описание | Когда вызывается |
|-----|----------|------------------|
| `on_message_send(text, api)` | Модификация текста перед отправкой | После нажатия «Отправить», до формирования запроса к модели |
| `on_response_received(text, api)` | Модификация ответа модели | После получения полного ответа, до добавления в историю |
| `on_plugin_loaded(api)` | Выполняется после загрузки плагина | При старте программы |
| `on_settings_changed(settings, api)` | Реакция на изменение настроек | После сохранения настроек |

Методы `on_plugin_loaded` и `on_settings_changed` являются опциональными. Вы можете реализовать только необходимые.

---

## Объект PluginAPI

В хуки передаётся объект `api`, который предоставляет доступ к функциям Echos.

```python
class PluginAPI:
    def log(self, message: str):
        """Вывод сообщения в лог программы (и в консоль)."""

    def get_current_chat_title(self) -> str | None:
        """Возвращает название текущего открытого чата или None."""
```

---

## Примеры плагинов

### Автозамена текста

```python
class Plugin:
    def on_message_send(self, text, api):
        # Заменяем "бот" на "ассистент"
        return text.replace("бот", "ассистент")
```

### Логирование ответов в файл

```python
class Plugin:
    def on_response_received(self, text, api):
        with open("responses.log", "a", encoding="utf-8") as f:
            f.write(text + "\n---\n")
        return text
```

###  Использование API для получения названия чата

```python
class Plugin:
    def on_response_received(self, text, api):
        title = api.get_current_chat_title()
        api.log(f"Ответ получен в чате: {title}")
        return text
```

---

## Расширенные возможности

### Работа с базой знаний (RAG)

Плагины могут добавлять документы в RAG напрямую:

```python
class Plugin:
    def on_plugin_loaded(self, api):
        main_window = api.main_window
        if main_window.rag_manager:
            main_window.rag_manager.add_document("Полезная информация...")
            api.log("Документ добавлен в RAG")
```

### Изменение параметров генерации

```python
class Plugin:
    def on_message_send(self, text, api):
        main_window = api.main_window
        main_window.settings["temperature"] = 0.3
        main_window.save_settings()
        return text
```

### Вызов диалоговых окон

```python
from PySide6.QtWidgets import QMessageBox

class Plugin:
    def on_response_received(self, text, api):
        QMessageBox.information(None, "Плагин", "Ответ получен!")
        return text
```

---

## Обработка ошибок

Чтобы плагин не нарушал работу программы, рекомендуется оборачивать потенциально опасный код в блоки `try-except`:

```python
class Plugin:
    def on_message_send(self, text, api):
        try:
            # Пример: заменить "бот" на "ассистент"
            result = text.replace("бот", "ассистент")
            return result
        except Exception as e:
            api.log(f"Ошибка в плагине: {e}")
            return text
```

---

## Рекомендации по разработке


- Проверяйте наличие сторонних библиотек внутри кода, чтобы избежать сбоев загрузки.

- Используйте `api.log()` для отладки.

- Документируйте свои плагины, добавляйте комментарии.

- Указывайте версию в атрибутах класса (например, `self.version = "1.0.0"`), если необходимо.

---

## Публикация и обмен

Плагин — это один файл `.py`, который можно скопировать в папку `plugins`. Для распространения можно упаковать плагин с дополнительными файлами в ZIP.

> 💡 Совет: Для более сложных плагинов вы можете использовать все возможности Python и PySide6, включая создание собственных окон и взаимодействие с API LM Studio.

Если у вас есть вопросы или предложения по улучшению системы плагинов, посетите [репозиторий](https://github.com/Vecsai/Echos "Подсказка") проекта.


---

## 📜 Лицензия

Проект распространяется под лицензией **MIT**.

Кратко: вы можете свободно использовать, изменять и распространять Echos, 
включая коммерческое использование. Единственное требование — сохранить 
уведомление об авторстве (файл [LICENSE](LICENSE)).

---

## © Автор

**Vecsai** — разработчик Echos.

Copyright (c) 2026 Vecsai. Распространяется под [MIT License](LICENSE).

---










