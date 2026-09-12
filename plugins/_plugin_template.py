# Echos — Private LLM client with encryption, RAG, and chat management
# Copyright (c) 2026 Vecsai
# Distributed under the MIT License. See LICENSE file for details.
# Project: https://github.com/Vecsai/Echos

"""
Echos — шаблон плагина.

Это пример плагина для Echos. Скопируйте этот файл, переименуйте
и измените под свои нужды.

Плагины автоматически загружаются из папки plugins/ при старте
программы. Дополнительно можно перезагрузить их через
Настройки → Плагины → Перезагрузить.

Доступные хуки:
    on_message_send(text, api)     — вызывается перед отправкой
                                     сообщения модели
    on_response_received(text, api) — вызывается после получения
                                     ответа от модели

Объект api (PluginAPI) предоставляет:
    api.log(message)            — запись в лог программы
    api.get_current_chat_title() — название текущего чата
"""

# class Plugin:
#     def on_message_send(self, text, api):
#         api.log(f"Сообщение до обработки: {text}")
#         # Добавляем свой текст к сообщению
#         return text + " (изменено плагином)"
#
#     def on_response_received(self, text, api):
#         api.log(f"Ответ до обработки: {text[:50]}...")
#         # Добавляем подпись к ответу
#         return text + "\n\n[Ответ обработан плагином]"