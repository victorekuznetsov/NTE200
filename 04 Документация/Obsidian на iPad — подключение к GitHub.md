---
title: Obsidian на iPad — подключение к GitHub
aliases: [Obsidian iPad GitHub, Синхронизация iPad]
tags: [obsidian, ipad, github, синхронизация, инструкция]
created: 2026-06-26
status: инструкция
---

# Obsidian на iPad — подключение к GitHub

> Репозиторий: `victorekuznetsov/nte200` · ветка по умолчанию: **`main`**
> В хранилище есть тяжёлые бинарники (pptx ~2.4 МБ, PDF до 6 МБ, zip) — это влияет на выбор способа.

## Шаг 0. Personal Access Token (нужен для обоих способов)

1. На GitHub: **Settings → Developer settings → Fine-grained tokens → Generate new token**.
2. Repository access → **Only select repositories** → `victorekuznetsov/nte200`.
3. Permissions → **Contents: Read and write** (этого достаточно).
4. Сохранить токен (показывается один раз). Использовать его **вместо пароля**; логин — имя пользователя GitHub.

---

## Способ A (рекомендую) — Working Copy + Obsidian

Надёжно для крупных файлов (презентации, PDF).

1. Установить приложение **Working Copy** (git-клиент для iOS).
2. **Clone repository** → войти через GitHub/токен → склонировать `victorekuznetsov/nte200`, ветка **`main`**.
3. Working Copy работает как провайдер в приложении **Файлы** — дать доступ к папке репозитория.
4. В Obsidian iOS: **«Open folder as vault»** → выбрать папку репозитория из Файлов.
5. Редактируете в Obsidian; **commit / push / pull делаете в Working Copy**.

---

## Способ B — плагин «Obsidian Git» (всё внутри Obsidian)

Поддерживает iOS, но git здесь чисто на JS (медленный, держит данные в памяти) — на тяжёлых файлах **может подвисать**.

1. Obsidian: **Настройки → Сторонние плагины → Обзор → «Git»** (автор Vinzent03) → установить, включить.
2. Команда **«Clone an existing repository»**, URL: `https://github.com/victorekuznetsov/nte200.git`, ветка **`main`**.
3. При запросе логина — имя пользователя GitHub, пароль — **токен** из Шага 0.
4. В настройках плагина указать author name/email. Дальше — Commit / Push / Pull.
5. **Совет:** отключить авто-pull при запуске, если зависает.

---

## Тяжёлые бинарники — компромисс

Для Способа B (и чтобы мобильная синхронизация была лёгкой) можно исключить крупные файлы из синхронизации, добавив в `.gitignore`:

```
*.pptx
*.pdf
*.zip
```

Тогда заметки синхронизируются быстро, но презентации/PDF **не будут попадать на iPad**. Для Способа A (Working Copy) это не требуется — он нормально тянет бинарники.

---

## Частые проблемы

| Симптом | Причина / решение |
|---|---|
| Просит пароль, аккаунтный не подходит | Нужен **токен** (Шаг 0), не пароль GitHub |
| Зависает при старте | Отключить авто-pull; для больших репо — Способ A |
| Конфликты при pull | Сначала закоммитить локальные правки, затем pull |
| «Authentication failed» | Токен без права **Contents: Read and write** или не на тот репозиторий |
