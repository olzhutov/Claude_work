#!/bin/bash
# Автоматический git push — запускается каждый день в 12:00 через launchd
# Установка: см. tools/com.decosta.git-push.plist

PROJECT="/Users/zhutovoleg/Doc/CLAUDE/CLAUDE_CODE"
LOG="$PROJECT/tools/git_push.log"
DATE=$(date '+%Y-%m-%d %H:%M')

cd "$PROJECT" || exit 1

# Проверяем есть ли изменения
if git diff --quiet && git diff --cached --quiet && [ -z "$(git ls-files --others --exclude-standard)" ]; then
    echo "[$DATE] Нет изменений — push пропущен" >> "$LOG"
    exit 0
fi

# Добавляем все файлы и делаем коммит
git add -A
git commit -m "Auto-commit $DATE"
GIT_TERMINAL_PROMPT=0 git -c credential.helper= push origin main

if [ $? -eq 0 ]; then
    echo "[$DATE] ✅ Push успешен" >> "$LOG"
else
    echo "[$DATE] ❌ Ошибка push" >> "$LOG"
fi
