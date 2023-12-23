from datetime import datetime
from typing import List, Dict, Union, Optional
from inventory_operations import (
    load_from_excel,
    save_to_excel,
    create_product,
    add_product,
    remove_product,
    edit_product,
    search_product,
    display_inventory,
    make_sale,
    generate_report,
)

def get_user_input(prompt: str, input_type: Union[type, None] = str) -> Optional[Union[str, int, float]]:
    while True:
        user_input = input(prompt)
        if not user_input:
            return None
        try:
            return input_type(user_input) if input_type else user_input
        except ValueError:
            print("Некорректный формат ввода. Пожалуйста, повторите.")

def add_new_product(inventory: dict) -> None:
    code = get_user_input("Введите код товара: ")
    name = get_user_input("Введите название товара: ")
    price = get_user_input("Введите цену товара: ", float)
    quantity = get_user_input("Введите количество товара: ", int)

    add_product(inventory, code, name, price, quantity)

def sell_products(inventory: dict) -> None:
    sale_products = []
    while True:
        product_code = get_user_input("Введите код товара (или '0' для завершения): ")
        if not product_code or product_code == '0':
            break
        quantity = get_user_input("Введите количество товара: ", int)

        product = next((p for p in inventory['products'] if p['code'] == product_code), None)
        if product and product['quantity'] >= quantity:
            sale_products.append(create_product(product['code'], product['name'], product['price'], quantity))
        else:
            print("Товар не найден или недостаточно в наличии.")

    payment_method = get_user_input("Выберите способ оплаты (наличные/безналичные): ")
    discount = get_user_input("Введите скидку (в процентах): ", float)
    tax_rate = get_user_input("Введите налог (в процентах): ", float)

    make_sale(inventory, sale_products, payment_method, discount, tax_rate)

if __name__ == "__main__":
    inventory_system = load_from_excel()

    while True:
        print("\nКоманды:")
        print("1. Добавить новый товар")
        print("2. Удалить товар")
        print("3. Редактировать товар")
        print("4. Поиск товара")
        print("5. Просмотреть инвентарь")
        print("6. Продать товары")
        print("7. Генерировать отчет")
        print("8. Выйти из программы")

        choice = get_user_input("Введите номер команды: ", int)

        if choice == 1:
            add_new_product(inventory_system)
        elif choice == 2:
            code = get_user_input("Введите код товара для удаления: ")
            remove_product(inventory_system, code)
        elif choice == 3:
            code = get_user_input("Введите код товара для редактирования: ")
            new_name = get_user_input("Введите новое название товара: ")
            new_price = get_user_input("Введите новую цену товара: ", float)
            new_quantity = get_user_input("Введите новое количество товара: ", int)
            edit_product(inventory_system, code, new_name, new_price, new_quantity)
        elif choice == 4:
            keyword = get_user_input("Введите название или код товара для поиска: ")
            search_product(inventory_system, keyword)
        elif choice == 5:
            display_inventory(inventory_system)
        elif choice == 6:
            sell_products(inventory_system)
        elif choice == 7:
            start_date_str = get_user_input("Введите начальную дату (в формате ГГГГ-ММ-ДД) или оставьте пустым: ")
            end_date_str = get_user_input("Введите конечную дату (в формате ГГГГ-ММ-ДД) или оставьте пустым: ")
            specific_product_code = get_user_input("Введите код товара для фильтрации отчета или оставьте пустым: ")
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d") if start_date_str else None
                end_date = datetime.strptime(end_date_str, "%Y-%m-%d") if end_date_str else None
            except ValueError:
                print("Некорректный формат даты.")
                continue
            generate_report(inventory_system, start_date, end_date, specific_product_code)
        elif choice == 8:
            save_to_excel(inventory_system)
            break
        else:
            print("Неверная команда. Пожалуйста, введите правильный номер команды.")
