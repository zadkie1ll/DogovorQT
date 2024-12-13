from datetime import datetime
from dateutil.relativedelta import relativedelta

def calculate_payments(num_payments, total_amount, discount, start_date, first_payment):
    # Вычисляем оставшуюся сумму для расчета регулярных платежей
    remaining_amount = (total_amount - discount) - first_payment
    remaining_payments = num_payments - 1  # Уменьшаем количество платежей на 1

    if remaining_payments <= 0:
        raise ValueError("Количество платежей должно быть больше 1 при наличии первого платежа.")

    # Рассчитываем регулярный платеж
    payment_amount = remaining_amount / remaining_payments
    payment_amount_rounded = round(payment_amount, -2)
    
    # Корректируем последний платеж на разницу округления
    total_rounded = payment_amount_rounded * remaining_payments
    difference = remaining_amount - total_rounded

    # Преобразуем дату начала платежей
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, "%d.%m.%Y")

    table_data = []

    # Добавляем первый платеж с сегодняшней датой
    today = datetime.today()
    table_data.append([1, today.strftime("%d.%m.%Y"), f"{first_payment:.2f}"])

    # Генерируем график для оставшихся платежей
    current_date = start_date
    for i in range(remaining_payments):
        current_date += relativedelta(months=1)

        # Учет февраля
        if current_date.month == 2 and current_date.day > 28:
            current_date = current_date.replace(day=28)
        else:
            try:
                current_date = current_date.replace(day=start_date.day)
            except ValueError:
                current_date = current_date.replace(day=1) + relativedelta(day=31)

        # Учет разницы для последнего платежа
        payment = payment_amount_rounded
        if i == remaining_payments - 1:
            payment += difference

        table_data.append([i + 2, current_date.strftime("%d.%m.%Y"), f"{payment:.2f}"])

    return table_data


schedule = calculate_payments(
    num_payments=6, 
    total_amount=100400, 
    discount=5967, 
    start_date="22.01.2024", 
    first_payment=10000
)

for row in schedule:
    print(row)