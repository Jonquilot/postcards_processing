from openpyxl import Workbook

from postcards_lib import compilate_data, compilate_grouped_counts, count_unsufficient_tags, table_compilate


def main():
    ### Ввод пути к файлу.
    user_path = input("Введите путь к файлу корпуса: ").strip()
    if not user_path:
        return print("ОШИБКА: Путь к файлу отсутствует, инициация программы прервана")

    ### Составление словарей соответствий номеров открытки с получателями/отправителями.
    try:
        senders_dict, recipients_dict = compilate_data(path=user_path)
    except FileNotFoundError:
        return print("ОШИБКА: Путь введен некоректно или файла не существует, инициация программы прервана.")
    if not senders_dict:
        return print("ОШИБКА: Данные об отправителях не были введены или не содержат ни одного элемента, инициация программы прервана")
    elif not recipients_dict:
        return print("ОШИБКА: Данные о получателях не были введены или не содержат ни одного элемента, инициация программы прервана")
    print("Данные получены.")

    ### Составление словарей соответствий между отправителями и получателями, подсчет кол-ва соответствий.
    print('Установка соответствий "отправители -> получатели"...')
    senders_compilated = compilate_grouped_counts(senders_dict, recipients_dict)
    if not senders_compilated:
        print("ОШИБКА: Установка соответствий ""отправители -> получатели"" невозможна, инициация программы прервана.")
        return
    print("Соответствия установлены.")
    print("Подсчет тегов [нрзб] и [отсутствует]...")
    senders_missing, senders_illegible = count_unsufficient_tags(senders_dict)
    print("Подсчет тегов завершен.")
    print('Установка соответствий "получатели -> отправители"...')
    recipients_compilated = compilate_grouped_counts(recipients_dict, senders_dict)
    if not recipients_compilated:
            print("ОШИБКА: Установка соответствий ""получатели -> отправители"" невозможна, инициация программы прервана.")
            return
    print("Соответствия установлены.")
    print("Подсчет тегов [нрзб] и [отсутствует]...")
    recipients_missing, recipients_illegible = count_unsufficient_tags(recipients_dict)
    print("Подсчет тегов завершен.")

    ### Заполнение таблицы скомпилированными данными.
    result = Workbook()
    print("Составление таблицы отправителей...")
    table_compilate(
        result, 
        senders_compilated, 
        senders_missing, 
        senders_illegible, 
        "Отправители - получатели", 
        "Отправители", 
        "Получатели", 
        "Кол-во открыток"
        )
    print("Таблица отправителей составлена")
    print("Составление таблицы получателей...")
    table_compilate(
        result, 
        recipients_compilated, 
        recipients_missing, 
        recipients_illegible, 
        "Получатели - отправители", 
        "Получатели", 
        "Отправители", 
        "Кол-во открыток"
        )
    print("Таблица получателей составлена")
    # удаляем созданный по умолчанию при инициализации Workbook пустой лист
    # т.к. данные записываются в отдельные именованные листы через table_compilating
    del result["Sheet"]
    result.save("Сопоставление получателей и отправителей.xlsx")

    print('Результат сохранен в файл "Сопоставление получателей и отправителей.xlsx". Спасибо за использование!')


if __name__ == '__main__':
    main()