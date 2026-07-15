import pandas as pd
import numpy as np
from pathlib import Path

# --- Настройки ---
PERIOD_DAYS = 30          # период выгрузки продаж в днях (для расчёта среднесуточных)
ANONYMIZE = True          # True — обезличить данные перед экспортом (для публикации)
MONEY_COEF = 0.87         # коэффициент искажения сумм (один для всех, пропорции сохраняются)

# --- Пути к файлам ---
base_path = Path(r'C:\Users\Windows\Documents\WB_Project')
data_folder = base_path / "raw_new"
sales_path = data_folder / 'sales.xlsx'
output_path = base_path / "data_exports"
output_path.mkdir(parents=True, exist_ok=True)

# --- Чтение выгрузки (реальный заголовок на второй строке) ---
sales_df = pd.read_excel(sales_path, engine='openpyxl', header=1)

# --- Отсечение транзитных складов (сортировочные центры, склады поставщика) ---
# Это не места хранения, а точки транзита — в анализ запасов не входят.
sales_df = sales_df[~sales_df['Склад'].str.contains('СЦ|SC|Склад поставщика', na=False)]

# --- Отбор нужных колонок и переименование в машинные имена ---
sales_cols = {
    'Бренд': 'brand',
    'Предмет': 'category',
    'Наименование': 'product_name',
    'Артикул продавца': 'seller_sku',
    'Артикул WB': 'wb_sku',
    'Склад': 'warehouse',
    'шт.': 'unit_sales',
    'Сумма заказов минус комиссия WB, руб.': 'sales_without_coms',
    'Текущий остаток, шт.': 'current_stock',
}
sales_df = sales_df[list(sales_cols.keys())].rename(columns=sales_cols)

# --- Свёртка до уникального зерна «товар × склад» ---
# Один товар может встречаться несколькими строками (размеры/баркоды) — суммируем.
sales_df_gr = sales_df.groupby(
    ['brand', 'category', 'product_name', 'seller_sku', 'wb_sku', 'warehouse']
).agg({
    'unit_sales': 'sum',
    'sales_without_coms': 'sum',
    'current_stock': 'sum',
}).reset_index()

# --- Аналитические метрики ---

# Среднесуточные продажи за период выгрузки
sales_df_gr['avg_daily_sales'] = sales_df_gr['unit_sales'] / PERIOD_DAYS

# Дни запаса: на сколько дней хватит остатка при текущем темпе продаж.
# У товаров без продаж делить не на что → NaN (оборачиваемость не определена).
sales_df_gr['days_of_supply'] = np.where(
    sales_df_gr['avg_daily_sales'] == 0,
    np.nan,
    sales_df_gr['current_stock'] / sales_df_gr['avg_daily_sales']
)

# Статус остатка — четыре взаимоисключающих ситуации по паре товар×склад
conditions = [
    (sales_df_gr['unit_sales'] > 0) & (sales_df_gr['current_stock'] == 0),  # продавался, остаток кончился
    (sales_df_gr['unit_sales'] == 0) & (sales_df_gr['current_stock'] > 0),  # лежит, но не продаётся
    (sales_df_gr['unit_sales'] > 0) & (sales_df_gr['current_stock'] > 0),   # и продаётся, и есть запас
]
choices = ['распродано', 'неликвид', 'в обороте']
sales_df_gr['stock_status'] = np.select(conditions, choices, default='нет активности')

# Цена за единицу (из продаж). Без продаж вывести цену нельзя → NaN.
sales_df_gr['price_per_unit'] = np.where(
    sales_df_gr['unit_sales'] == 0,
    np.nan,
    sales_df_gr['sales_without_coms'] / sales_df_gr['unit_sales']
)

# Стоимость запаса (замороженный капитал) = остаток × цена за единицу.
# Замечание: у неликвида (0 продаж) цены нет, поэтому в деньгах он не оценивается.
sales_df_gr['frozen_money'] = sales_df_gr['current_stock'] * sales_df_gr['price_per_unit']

# --- Обезличивание для публикации ---
# Скрываем реальные бренды, названия и суммы конкретного продавца,
# сохраняя структуру и все пропорции — аналитическая ценность не теряется.
if ANONYMIZE:
    # Бренды → «Бренд А», «Бренд Б», ...
    brands = sorted(sales_df_gr['brand'].dropna().unique())
    brand_map = {b: f'Бренд {chr(1040 + i)}' for i, b in enumerate(brands)}  # 1040 = «А»
    sales_df_gr['brand'] = sales_df_gr['brand'].map(brand_map)

    # Артикулы WB → последовательные условные номера
    wb_ids = sorted(sales_df_gr['wb_sku'].unique())
    wb_map = {wb: 100000 + i for i, wb in enumerate(wb_ids)}
    sales_df_gr['wb_sku'] = sales_df_gr['wb_sku'].map(wb_map)

    # Названия товаров → «Товар 0001» (нумерация привязана к артикулу, чтобы
    # один товар везде назывался одинаково)
    name_map = {wb: f'Товар {i + 1:04d}' for i, wb in enumerate(wb_ids)}
    sales_df_gr['product_name'] = sales_df_gr['wb_sku'].map(
        {v: name_map[k] for k, v in wb_map.items()}
    )

    # Артикул продавца → условный код на основе нового wb_sku
    sales_df_gr['seller_sku'] = 'SKU-' + sales_df_gr['wb_sku'].astype(str)

    # Денежные колонки → умножаем на единый коэффициент.
    # Один коэффициент для всех строк и колонок → пропорции и все метрики сохраняются,
    # но абсолютные суммы перестают быть реальными.
    for col in ['sales_without_coms', 'price_per_unit', 'frozen_money']:
        sales_df_gr[col] = sales_df_gr[col] * MONEY_COEF

# --- Экспорт для Power BI (utf-8-sig — чтобы Excel корректно читал кириллицу) ---
filename = "sales_result.csv"
sales_df_gr.to_csv(output_path / filename, index=False, encoding='utf-8-sig')

print(f"Готово. Строк: {len(sales_df_gr)}. Файл: {output_path / filename}")

