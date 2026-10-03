# Echos — Private LLM client with encryption, RAG, and chat management
# Copyright (c) 2026 Vecsai
# Distributed under the MIT License. See LICENSE file for details.
# Project: https://github.com/Vecsai/Echos

"""Читатель сокета для консоли отладки Echos.
Собирается в отдельный Echos_reader.exe. Зависимостей нет,
кроме стандартной библиотеки Python."""
import sys
import socket
import time


def main():
    if len(sys.argv) < 2:
        print("[reader] Нет порта в аргументах")
        time.sleep(3)
        return

    try:
        port = int(sys.argv[1])
    except ValueError:
        print(f"[reader] Неверный порт: {sys.argv[1]}")
        time.sleep(3)
        return

    # UTF-8 вывод в консоль
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

    # Подключаемся к серверу с ретраями
    s = None
    for _ in range(25):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.connect(("127.0.0.1", port))
            break
        except Exception:
            try:
                if s:
                    s.close()
            except Exception:
                pass
            s = None
            time.sleep(0.2)

    if s is None:
        print("[reader] Не удалось подключиться к Echos")
        time.sleep(10)
        return

    print("=" * 60)
    print("  ECHOS - КОНСОЛЬ ОТЛАДКИ")
    print("  Закрытие окна НЕ завершает Echos.")
    print("=" * 60)
    print()

    # Читаем и печатаем поток
    try:
        while True:
            data = s.recv(4096)
            if not data:
                break
            sys.stdout.write(data.decode("utf-8", errors="replace"))
            sys.stdout.flush()
    except Exception:
        pass
    finally:
        try:
            s.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
