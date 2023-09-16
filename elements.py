# Збереження ціни насосу у валюті:
pump_price_currency = get_pump_price(pump_info)

# Пошук групи знижок:
marker_row = None
for row in reversed(list(sheet_with_pumps.iter_rows(values_only=True))):
    if "PG" in row:
        marker_row = row
        break

if marker_row:
    group_name = marker_row[1]  # Припускаючи, що назва групи у другому стовпці
else:
    group_name = None

# Зчитування числового значення знижки:
if group_name:
    discount = None
    for row in sheet_with_discounts.iter_rows(values_only=True):
        if group_name.lower() == row[0].lower():  # Припускаємо, що назва групи в першому стовпці
            discount = row[1]  # Припускаємо, що значення знижки у другому стовпці
            break
else:
    discount = None

# Розрахунок ціни насосу у гривнях:
if pump_price_currency is not None and discount is not None:
    pump_price_uah = pump_price_currency * (1 - discount) * currency_rate
else:
    pump_price_uah = None