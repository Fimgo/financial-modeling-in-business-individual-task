# Starbucks Financial Model

Індивідуальне завдання з дисципліни **"Фінансове моделювання в бізнесі"**  
Тема: **Формування власної фінансової моделі за даними реального підприємства**

## 1. Що знаходиться в папці

У папці `Starbucks_Financial_Model` підготовлено базовий комплект для роботи:

| Файл | Призначення |
|---|---|
| `Starbucks_Financial_Model.xlsx` | Основний Excel-файл із фінансовою моделлю Starbucks |
| `starbucks_financial_model.R` | R-скрипт із логікою моделі, історичними даними, припущеннями та розрахунками |
| `build_starbucks_workbook.py` | Технічний скрипт, яким було згенеровано Excel-файл у цьому середовищі |
| `README.md` | Пояснення структури роботи, джерел, моделі та графіків |

Основний файл для здачі як фінансова модель:  
**`Starbucks_Financial_Model.xlsx`**

Основний файл із кодом моделі:  
**`starbucks_financial_model.R`**

## 2. Обране підприємство

Для моделі обрано **Starbucks Corporation**.

Starbucks Corporation - міжнародна компанія у сфері спеціалізованої кавової роздрібної торгівлі та ресторанного бізнесу. Компанія управляє мережею кав'ярень, продає кавові напої, чай, їжу, упаковану каву, готові напої та супутні товари.

Основні джерела доходів:

- продаж напоїв у власних кав'ярнях;
- продаж їжі та супутніх товарів;
- ліцензовані магазини;
- продаж упакованої кави та готових напоїв через канали роздрібної торгівлі.

Основні групи витрат:

- product and distribution costs;
- store operating expenses;
- other operating expenses;
- depreciation and amortization;
- general and administrative expenses;
- restructuring and impairments.

## 3. Джерела даних

Фінансові дані взяті з офіційних джерел Starbucks:

1. **Starbucks Fiscal 2025 Annual Report PDF**  
   https://s203.q4cdn.com/326826266/files/doc_financials/2025/ar/Starbucks-Corporation_2025-Annual-Report-Web-Ready.pdf

2. **Starbucks Investor Relations - Annual Reports**  
   https://investor.starbucks.com/financials/annual-reports/

3. **Starbucks SEC Filing Details - Form 10-K**  
   https://investor.starbucks.com/financials/sec-filings/sec-filings-details/default.aspx?FilingId=18927461

У самому Excel-файлі ці джерела винесені на окремий аркуш **`Sources`**, щоб їх можна було швидко перевірити.

## 4. Структура Excel-моделі

Файл **`Starbucks_Financial_Model.xlsx`** складається з таких аркушів:

| Аркуш | Зміст |
|---|---|
| `Cover` | Загальна інформація про модель, компанію, періоди та статус перевірки |
| `Data` | Історичні фінансові дані Starbucks за 2023-2025 роки |
| `Assumptions` | Прогнозні припущення для 2026E-2028E |
| `P&L` | Спрощений звіт про прибутки та збитки |
| `Break-even` | Розрахунок точки беззбитковості |
| `Sensitivity` | Аналіз чутливості прибутку до зміни доходів і витрат |
| `Charts` | Графіки для візуального аналізу |
| `Checks` | Перевірки коректності моделі |
| `Sources` | Джерела даних та посилання |

## 5. Де знаходяться графіки

Графіки знаходяться в Excel-файлі **`Starbucks_Financial_Model.xlsx`** на аркуші:

**`Charts`**

На цьому аркуші є 4 графіки:

1. **Revenue Dynamics** - динаміка доходів Starbucks.
2. **Operating Income** - динаміка операційного прибутку.
3. **Operating Margin** - динаміка операційної маржі.
4. **Revenue vs Total Operating Expenses** - порівняння доходів і операційних витрат.

Вимога завдання рекомендує 2-3 графіки, тому модель навіть трохи перевиконує цю частину: у файлі є 4 графіки.

## 6. Логіка фінансової моделі

Модель побудована за простою та прозорою логікою.

Історичні дані:

- 2023A;
- 2024A;
- 2025A.

Прогнозні роки:

- 2026E;
- 2027E;
- 2028E.

Основні формули:

```text
Revenue = Revenue попереднього року * (1 + Revenue growth)

Variable Costs = Revenue * Variable cost ratio

Fixed Costs = Depreciation and amortization
            + General and administrative expenses
            + Restructuring and impairments

Operating Income = Revenue
                 - Total Operating Expenses
                 + Income from equity investees

Operating Margin = Operating Income / Revenue
```

## 7. Припущення моделі

Основні припущення винесені на аркуш **`Assumptions`**.

Ключові припущення:

- виручка у 2026E зростає на 3.5%;
- виручка у 2027E та 2028E зростає на 4.0%;
- variable cost ratio поступово знижується з 76.5% до 75.5%;
- fixed cost growth становить 3.0% на рік;
- restructuring costs знижуються після підвищеного рівня 2025 року.

Ці припущення є редагованими: у Excel вони виділені синім текстом і жовтим фоном.

## 8. Break-even analysis

Точка беззбитковості розрахована на аркуші **`Break-even`**.

Формула:

```text
Contribution Margin % = 1 - Variable Cost Ratio

Break-even Revenue = Fixed Costs / Contribution Margin %
```

Для цієї моделі break-even revenue показує, який рівень доходу потрібен Starbucks, щоб покрити основні операційні витрати в прогнозному періоді.

## 9. Sensitivity analysis

Аналіз чутливості знаходиться на аркуші **`Sensitivity`**.

Він показує, як зміниться operating income у 2026E за різних комбінацій:

- revenue growth delta: `-5%`, `0%`, `+5%`;
- variable cost ratio: `74.5%`, `76.5%`, `78.5%`.

Це дозволяє побачити, наскільки фінансовий результат залежить від зміни обсягу продажів і рівня витрат.

## 10. Відповідність вимогам завдання

| Вимога | Де виконано |
|---|---|
| Опис підприємства | README, аркуш `Cover`, майбутній текстовий звіт |
| Збір і підготовка даних | Аркуші `Data` і `Sources` |
| Прогноз доходів | Аркуш `P&L` |
| Прогноз витрат | Аркуш `P&L` |
| Розрахунок прибутку | Аркуш `P&L` |
| Спрощений P&L | Аркуш `P&L` |
| Точка беззбитковості | Аркуш `Break-even` |
| Аналіз чутливості | Аркуш `Sensitivity` |
| 2-3 графіки | Аркуш `Charts`, фактично 4 графіки |
| Джерела даних | Аркуш `Sources` |
| Перевірка моделі | Аркуш `Checks` |

## 11. Висновки для ворда 

На основі моделі можна зробити такі попередні висновки:

1. Starbucks залишається компанією з великим масштабом операцій та стабільною базою доходів.
2. У 2025 році операційна прибутковість знизилася через зростання операційних витрат і restructuring and impairments.
3. Прогноз моделі передбачає поступове відновлення прибутковості у 2026-2028 роках за рахунок помірного зростання доходів і нормалізації витрат.
4. Аналіз чутливості показує, що operating income сильно залежить від variable cost ratio, тобто контроль витрат є критичним фактором для фінансового результату.
5. Break-even analysis показує, що прогнозний рівень доходів значно перевищує розраховану точку беззбитковості, отже в базовому сценарії компанія має запас фінансової стійкості.


