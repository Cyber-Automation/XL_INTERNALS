# XL\_INTERNALS (EN)

VBA component ecosystem for Microsoft Excel and Office.  
No external dependencies. Windows only (x32/x64).

📖 Documentation: [script-coding.ru](https://script-coding.ru)

\---

## Repository Structure

```
01-VBUtilities/          # Standalone Excel applications
02-Developer-Kits/       # Self-contained developer components
03-Background-Services/  # Security demonstration modules
04-Other-Components/     # Miscellaneous utilities
Documentation/
LICENSE
README.md
```

\---

## 01 · VBUtilities

|Component|Description|
|-|-|
|VBUtil\_RAT|Remote administration system; isolated execution subsystems, custom engines and services|
|VBUtil\_ExcelServicesManager|SDI Excel process scheduler|
|VBUtil\_FileEncryptor|File encryption GUI|
|VBUtil\_Information|Hardware and system configuration inspector|
|VBUtil\_MediaPlayer|Audio playback inside Excel, no third-party software|
|VBUtil\_XFormExcel|Framework for standalone GUI applications with hidden Excel interface|

\---

## 02 · Developer Kits

Each component is self-contained. Drop into any VBA project — no manual configuration required.  
Object model access is set programmatically on the first call to any public function of the component.

|Component|Description|
|-|-|
|VBD\_Kit\_BuilderUI|Runtime UserForm construction during VBA code execution|
|VBD\_Kit\_Cryptography|Cryptography suite in pure VBA|
|VBD\_Kit\_Hashing|Component for structured hashing of various data types|
|VBD\_Kit\_Interface\_SDI|Excel process manager with isolated execution|
|VBD\_Kit\_ProcessBar|Runtime progress bar generation|
|VBD\_Kit\_Properties|Excel file property control (BDP, CDP, CP)|
|VBD\_Kit\_RTTI|Runtime type identification and VBA code parsing|
|VBD\_Kit\_Security|Security audit and configuration management|
|VBD\_Kit\_UnitTest\_Framework|Unit testing framework for VBA|
|VBD\_Kit\_WMI|System information via WMI|

\---

## 03 · Background Services

Demonstrations of potentially harmful VBA capabilities.  
No working exploits are included. Each module contains detection and mitigation notes.

|Component|Description|
|-|-|
|modSDI\_Clipboard|Clipboard interception concept|
|modSDI\_DispatcherMonitor|Task Manager interaction concept|
|modSDI\_KeyLogger|Keystroke capture concept|
|modSDI\_ResetPassword|Workbook open-password vulnerability demonstration|
|modSDI\_Screenshots|Screenshot capture via WinAPI|
|modSDI\_ShutDown|Shutdown/restart interception concept|
|modSDI\_SystemNotifier|Background native Windows notification system|

\---

## 04 · Other Components

|Component|Description|
|-|-|
|modProc\_ChangeVBETheme|Programmatic VBE color theme switching|
|modProc\_QuickFileSearch|File search via WinAPI|

\---

## Requirements

* Excel 2016 or later
* Windows 10 / 11
* Office x32 or x64

\---

## License

See [LICENSE](./LICENSE).

<hr style="border: none; height: 2px; background: linear-gradient(to right, #ff6b6b, #4ecdc4); margin: 20px 0;">

# 

# XL\_INTERNALS (RU)

Экосистема VBA-компонентов для Microsoft Excel и Office.  
Внешние зависимости отсутствуют. Только Windows (x32/x64).

📖 Документация: [script-coding.ru](https://script-coding.ru)

\---

## Структура репозитория

```
01-VBUtilities/          # Standalone Excel-приложения
02-Developer-Kits/       # Самодостаточные компоненты для разработчиков
03-Background-Services/  # Демонстрационные модули по безопасности
04-Other-Components/     # Прочие утилиты
Documentation/
LICENSE
README.md
```

\---

## 01 · VBUtilities

|Компонент|Описание|
|-|-|
|VBUtil\_RAT|Система удалённого администрирования; изолированные подсистемы исполнения, собственные движки и сервисы|
|VBUtil\_ExcelServicesManager|Планировщик SDI-процессов Excel|
|VBUtil\_FileEncryptor|GUI-шифровальщик файлов|
|VBUtil\_Information|Информация об железе и конфигурации системы|
|VBUtil\_MediaPlayer|Воспроизведение аудио внутри Excel без сторонних программ|
|VBUtil\_XFormExcel|Фреймворк для standalone GUI-приложений с скрытым интерфейсом Excel|

\---

## 02 · Developer Kits

Каждый компонент самодостаточен. Достаточно вставить в проект — настройка не требуется.  
Доступ к объектной модели выставляется программно при первом запуске любой публичной функции компонента.

|Компонент|Описание|
|-|-|
|VBD\_Kit\_BuilderUI|Динамическое создание UserForm во время выполнения кода VBA.|
|VBD\_Kit\_Cryptography|Криптография на чистом VBA|
|VBD\_Kit\_Hashing|Компонент для хеширования различных данных|
|VBD\_Kit\_Interface\_SDI|Менеджер процессов Excel с изолированным исполнением|
|VBD\_Kit\_ProcessBar|Динамическое создание прогрессбаров|
|VBD\_Kit\_Properties|Управление свойствами файлов Excel (BDP, CDP, CP)|
|VBD\_Kit\_RTTI|Динамическая идентификация типов и парсинг VBA-кода|
|VBD\_Kit\_Security|Аудит и управление конфигурацией безопасности|
|VBD\_Kit\_UnitTest\_Framework|Фреймворк модульного тестирования для VBA|
|VBD\_Kit\_WMI|Информация о системе через WMI|

\---

## 03 · Background Services

Демонстрации потенциально опасных возможностей VBA.  
Рабочие эксплойты не включены. Каждый модуль содержит описание методов обнаружения и защиты.

|Компонент|Описание|
|-|-|
|modSDI\_Clipboard|Концепция перехвата буфера обмена|
|modSDI\_DispatcherMonitor|Концепция взаимодействия с диспетчером задач|
|modSDI\_KeyLogger|Концепция перехвата нажатий клавиш|
|modSDI\_ResetPassword|Демонстрация уязвимости пароля книги при открытии|
|modSDI\_Screenshots|Создание скриншотов через WinAPI|
|modSDI\_ShutDown|Концепция перехвата выключения / перезагрузки|
|modSDI\_SystemNotifier|Фоновая система нативных уведомлений Windows|

\---

## 04 · Other Components

|Компонент|Описание|
|-|-|
|modProc\_ChangeVBETheme|Программная смена цветовой схемы VBE|
|modProc\_QuickFileSearch|Быстрый поиск файлов через WinAPI|

\---

## Требования

* Excel 2016 и новее
* Windows 10 / 11
* Office x32 или x64

\---

## Лицензия

См. [LICENSE](./LICENSE).

