# Excel Data Processor & Merger 🚀

![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue)
![Poetry](https://img.shields.io/badge/Poetry-1.6%2B-orange)
![License](https://img.shields.io/badge/License-MIT-green)
![Open Issues](https://img.shields.io/github/issues/GenyaElamkov/auto_exc)
![Stars](https://img.shields.io/github/stars/GenyaElamkov/auto_exc)

**Инструмент для автоматизации обработки Excel-файлов:** извлечение структурированных данных, многопоточная конвертация в CSV, объединение в единый XLSX-отчет с последующей очисткой временных файлов.

---

## 🔥 Ключевые возможности
- **Интеллектуальный парсинг**  
  Автоматическое извлечение данных из сложных Excel-шаблонов (финансовые ведомости, учетные записи).
- **Многопоточная обработка**  
  Оптимизация скорости через `multiprocessing.Pool` (масштабируется под CPU ядра).
- **Гибкая конфигурация**  
  Поддержка ручного и автоматического режимов работы.
- **Безопасность данных**  
  Встроенная валидация входных файлов и предупреждение о резервном копировании.
- **Эффективное управление памятью**  
  Режим `read_only` для работы с большими файлами + принудительный вызов `gc.collect()`.

---

## ⚠️ Важно!
- **Исходные файлы удаляются после обработки!**  
  Перед запуском создайте копии данных.
- Поддерживаются **только .xlsx** файлы определенного формата (пример структуры в [документации](documents\document.xlsx)).

---

## 🛠 Установка

### Вариант 1: С использованием Poetry (рекомендуется)
```bash
# Установите Poetry (если не установлен)
curl -sSL https://install.python-poetry.org | python3 -

# Клонировать репозиторий
git clone git clone https://github.com/GenyaElamkov/auto_exc.git
cd auto_exc

# Установить зависимости и создать виртуальное окружение
poetry install

# Активировать окружение
poetry shell
```

### Вариант 2: Классическая установка
```bash
pip install -r requirements.txt
```


## 📦 Зависимости
Управление зависимостями через Poetry (см. [pyproject.toml](pyproject.toml)):
```toml
[tool.poetry.dependencies]
python = "^3.9"
pandas = "^2.0.3"
openpyxl = "^3.1.2"
colorama = "^0.4.6"
```
---

## 🖥 Использование

### Автоматический режим (рекомендуется)
```bash
python auto_exc.py
```
1. Выберите `1` в меню  
2. Все файлы в директории скрипта будут:  
   ✅ Обработаны  
   ✅ Конвертированы в CSV  
   ✅ Объединены в `combined_data.xlsx`  
   ✅ Временные CSV удалены  

### Ручные операции
| Режим | Команда | Назначение |
|-------|---------|------------|
| Только конвертация | Выбрать `2` | Создание CSV без объединения |
| Только объединение | Выбрать `3` | Сбор существующих CSV в XLSX |

---

## 📂 Структура проекта
```
.
├── 00_Data/                  # Выходные данные
├── src/
│   └── auto_exc.py           # Точка входа, Ядро логики
├── requirements.txt
└── README.md
```

---

## 📊 Пример вывода
```python
Скрипт выполнялся 12.8924 секунд
Обработано файлов: 47/47
Данные сохранены в файл: 00_Data/combined_data.xlsx
```

---

## 🤝 Contributing
1. Форкните репозиторий
2. Создайте ветку: `git checkout -b feature/your-feature`
3. Сделайте коммит: `git commit -m 'Add some feature'`
4. Запушьте: `git push origin feature/your-feature`
5. Откройте Pull Request

---

## 📜 Лицензия  
MIT License. Подробнее в [LICENSE](LICENSE).

---
