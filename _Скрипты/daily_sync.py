"""
NTE200 Daily Sync Agent
Запускается каждый день в 12:00 через Windows Task Scheduler.
Классифицирует новые файлы в корне проекта, раскладывает по папкам Obsidian,
создаёт .md-заметки для техотчётов, обновляет индексы и пушит на GitHub.
"""

import re
import subprocess
from datetime import datetime
from pathlib import Path

ROOT = Path("C:/Projects/nte200")
LOG_FILE = ROOT / "_Данные/sync_log.md"

# --- Правила классификации: (regex-паттерн, папка_назначения, создать_.md) ---
RULES = [
    # ECM данные
    (r"^\d+_ECM\d+_\d{2}\.\d{2}\.\d{2}\.xlsx$",       "_Данные/ECM",                  False),
    # DML NTE200
    (r"^DML[-_].*\.csv$",                               "_Данные/DML NTE200",           False),
    (r"^NTE200[A ].*\.csv$",                            "_Данные/DML NTE200",           False),
    # DML 730E
    (r"^730[EЕ].*\.csv$",                               "_Данные/DML 730E",             False),
    # Паспорта топлива
    (r"^Паспорт[^.]+\.pdf$",                            "_Данные/Паспорта топлива",     False),
    # Insite .eif
    (r"^I-\d{8}-\d+-\d{2}\.eif$",                      "_Данные/Insite",               False),
    # КАМСС GIF/скриншоты
    (r"^.*\.(GIF|gif)$",                                "_Данные/КАМСС",                False),
    # Масло
    (r"^Свод_анализов",                                 "_Данные/Масло",                False),
    (r"^Доливки",                                       "_Данные/Масло",                False),
    (r"^ПРОГРАММА.*МАСЛ",                               "_Данные/Масло",                False),
    # Топливо
    (r"^(730|NTE200)_fuel\.xlsx$",                      "_Данные/Топливо",              False),
    # ГЕ (General Electric привод — события Event 635)
    (r"^GE[_\- ]",                                      "_Данные/ГЕ",                   False),
    (r"^Ежесменный",                                    "_Данные/ГЕ",                   False),
    # Весовая (10/10/20 нагрузка)
    (r"^Весовая",                                       "_Данные/Весовая",              False),
    (r"^weight",                                        "_Данные/Весовая",              False),
    # Аналитические данные
    (r"^verify_",                                       "_Данные/Аналитика",            False),
    # Сводные данные
    (r"^(ОТЧЕТ|Сводный_анализ|ГБЦ ремонты|ECM_Analysis|ECM_对比)", "_Данные/Сводные", False),
    # Технические отчёты (с созданием .md-заметки)
    (r"^ТЕХНИЧЕСКИЙ ОТЧЁТ",                             "05 Отчёты/Технические",        True),
    (r"^(Тех отчет|Тех\.отчет|Тех оточет)",            "05 Отчёты/Технические",        True),
    (r"ТЕХНИЧЕСКИЙ ОТЧЁТ.*клапанов",                    "05 Отчёты/Технические",        True),
    (r"^\d+ ТЕХНИЧЕСКИЙ ОТЧЁТ",                         "05 Отчёты/Технические",        True),
    # Тесты ГБЦ
    (r"^ГБЦ#\d+\.pdf$",                                "05 Отчёты/Тесты ГБЦ",         False),
    (r"^.*клапан гбц.*\.pdf$",                          "05 Отчёты/Тесты ГБЦ",         False),
    (r"^Тест [LR]\d",                                   "05 Отчёты/Тесты ГБЦ",         False),
    (r"^\d+ \d+ \([0-9][LR]\)",                         "05 Отчёты/Тесты ГБЦ",         False),
    # Форсунки
    (r"[Фф]орсунк",                                     "05 Отчёты/Форсунки",           False),
    # Аналитические отчёты
    (r"^(Анализ_|АНАЛИЗ_)",                             "05 Отчёты/Аналитические",      False),
    (r"^Коренные причины",                              "05 Отчёты/Аналитические",      False),
    # Презентации
    (r"\.pptx$",                                        "05 Отчёты/Презентации",        False),
    # Архивы
    (r"\.(7z|zip|rar)$",                                "_Архивы",                      False),
    # Изображения
    (r"\.(png|webp|jpg|jpeg)$",                         "_Изображения",                 False),
    # Python-скрипты
    (r"\.py$",                                          "_Скрипты",                     False),
]

# Файлы которые никогда не трогаем в корне
EXCLUDED_NAMES = {
    ".gitignore", "README.md", "00 Главная.md",
    "_Данные", "_Скрипты", "_Архивы", "_Изображения", "_Шаблоны",
}
EXCLUDED_PREFIXES = (".", "01 ", "02 ", "03 ", "04 ", "05 ", "06 ", "07 ", "_")


def excluded(f: Path) -> bool:
    name = f.name
    if name in EXCLUDED_NAMES:
        return True
    for p in EXCLUDED_PREFIXES:
        if name.startswith(p):
            return True
    return False


def classify(name: str):
    for pattern, folder, create_note in RULES:
        if re.search(pattern, name, re.IGNORECASE):
            return folder, create_note
    return None, False


def extract_machine_number(name: str) -> str:
    m = re.search(r"[гГ]ар[№#](\d+)|[№#](\d+)", name)
    if m:
        return m.group(1) or m.group(2)
    return ""


def create_md_note(filepath: Path, machine: str, today: str):
    stem = filepath.stem
    note_name = f"{stem}.md"
    note_path = filepath.parent / note_name

    if note_path.exists():
        return

    machine_link = f"[[Реестр машин NTE200#№{machine}]]" if machine else ""
    content = f"""---
date: {today}
machine: {machine or ""}
file_type: техотчёт
tags:
  - техотчёт
  - qsk50
---

# {stem}

[[{filepath.name}|Открыть файл]]

- Машина: {machine_link}
- Проблема: [[Износ клапанов ГБЦ QSK50]]
"""
    note_path.write_text(content, encoding="utf-8")


def update_index(folder: str, filename: str):
    index_path = ROOT / folder / "_Индекс.md"
    if not index_path.exists():
        return
    content = index_path.read_text(encoding="utf-8")
    link = f"[[{filename}]]"
    if link in content or filename in content:
        return
    addition = f"\n- {link}"
    index_path.write_text(content + addition, encoding="utf-8")


def write_log(moved: list, today: str):
    if not moved:
        return
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"\n## {today}\n\nПеремещено файлов: {len(moved)}\n"]
    for name, folder, created_note in moved:
        note_mark = " + создана заметка" if created_note else ""
        lines.append(f"- `{name}` → `{folder}`{note_mark}")
    lines.append("")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write("\n".join(lines))


def git_run(args: list) -> tuple:
    try:
        subprocess.run(args, cwd=ROOT, check=True, capture_output=True)
        return True, "OK"
    except subprocess.CalledProcessError as e:
        return False, e.stderr.decode("utf-8", errors="replace") if e.stderr else str(e)


def main():
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    today = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 1. Pull ПЕРВЫМ — до любых изменений файлов
    pull_ok, pull_msg = git_run(["git", "pull", "--ff-only", "origin", "main"])
    pull_status = "OK" if pull_ok else f"FAIL: {pull_msg.strip()[:80]}"

    moved = []

    # 2. Найти файлы в корне которые нужно разложить
    candidates = [
        f for f in ROOT.iterdir()
        if f.is_file() and not excluded(f)
    ]

    # 3. Классифицировать и переместить
    for f in candidates:
        folder, create_note = classify(f.name)
        if not folder:
            continue
        dest_dir = ROOT / folder
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f.name
        if dest.exists():
            continue
        f.rename(dest)
        moved.append((f.name, folder, create_note))

    # 4. Создать .md заметки для техотчётов
    for name, folder, create_note in moved:
        if create_note:
            machine = extract_machine_number(name)
            create_md_note(ROOT / folder / name, machine, today[:10])

    # 5. Обновить индексы затронутых папок (все файлы, а не только первый)
    for name, folder, _ in moved:
        update_index(folder, name)

    # 6. Записать лог
    write_log(moved, today)

    # 7. Commit + Push
    git_run(["git", "add", "-A"])
    has_changes = subprocess.run(
        ["git", "diff", "--cached", "--quiet"], cwd=ROOT, capture_output=True
    ).returncode != 0
    if has_changes:
        git_run(["git", "commit", "-m", f"Синхронизация {today[:10]}: {len(moved)} файлов"])
    push_ok, push_msg = git_run(["git", "push", "origin", "main"])
    push_status = "OK" if push_ok else f"FAIL: {push_msg.strip()[:80]}"

    print(f"[{today}] pull:{pull_status} | push:{push_status} | перемещено:{len(moved)}")


if __name__ == "__main__":
    main()
