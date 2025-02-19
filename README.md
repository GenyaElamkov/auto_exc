# Excel Data Processor & Merger 🚀

![Python Version](https://img.shields.io/badge/Python-3.9%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Open Issues](https://img.shields.io/github/issues/yourusername/excel-data-processor)
![Stars](https://img.shields.io/github/stars/yourusername/excel-data-processor)

**Профессиональный инструмент для автоматизации обработки Excel-файлов:** извлечение структурированных данных, многопоточная конвертация в CSV, объединение в единый XLSX-отчет с последующей очисткой временных файлов.

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
- Поддерживаются **только .xlsx** файлы определенного формата (пример структуры в [документации](ссылка)).

---

## 🛠 Установка
```bash
# Клонировать репозиторий
git clone https://github.com/yourusername/excel-data-processor.git

# Установить зависимости
pip install -r requirements.txt
```

**Требования:**
- Python 3.9+
- Библиотеки: `pandas>=1.4.0`, `openpyxl>=3.0.10`, `colorama>=0.4.4`

---

## 🖥 Использование

### Автоматический режим (рекомендуется)
```bash
python main.py
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
├── 00_Data/                 # Выходные данные
├── src/
│   ├── main.py              # Точка входа
│   └── processor.py        # Ядро логики
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

## ✉️ Контакты  
**Автор:** Иван Петров  
**Email:** data-tools@proton.me  
**Telegram:** [@excel_master](https://t.me/excel_master)  

[![Sponsor](https://img.shields.io/badge/Support-Project-blue?style=for-the-badge)](https://ko-fi.com/yourprofile)
