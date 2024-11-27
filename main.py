import json
import random
from openpyxl import Workbook


def generate_schedule_and_save(data_file, excel_file="schedule.xlsx"):
    # Завантаження даних з JSON
    data = json.load(open(data_file))

    # Пустий розклад
    schedule = {}

    # Ітерування по кожній групі
    for group in data["groups"]:
        # Ітерування по кожному дню
        for weekday in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
            # Ітерування по кожномо номеру пари
            for time in data["times"]:
                # Список доступних лекторів для цього номеру пари
                available_lectors = []

                # Перевірка чи лектор може бути призначений
                for lector, details in data["lectors_info"].items():
                    if details["subjects"]:  # Перевірка чи лектор має предмети
                        # Перевірка чи лектор не призначений у цей день
                        if not any(
                                value[0] == lector
                                for key, value in schedule.items()
                                if key[0] == group and key[1] == weekday
                        ):
                            # Перевірка чи лектор має доступні години навантаження
                            if details["hours"] < details["max_hours"]:
                                available_lectors.append(lector)

                # Сортування лекторів по залишку годин
                if not available_lectors:
                    available_lectors = sorted(
                        available_lectors,
                        key=lambda prof: data["lectors_info"][prof]["max_hours"] - data["lectors_info"][prof]["hours"]
                    )

                # Призначення лектора на цей номер пари
                for lector in available_lectors:
                    subjects = data["lectors_info"][lector]["subjects"]  # Список предметів, які може вести
                    if subjects:
                        # Випадковий предмет з списку предметів
                        subject = random.choice(subjects)

                        # Список вільних кімнат
                        available_rooms = []
                        for room in data["rooms"]:
                            # Перевірка чи кітната не задіяна уже
                            if not any(
                                    value[2] == room
                                    for key, value in schedule.items()
                                    if key[1] == weekday and key[2] == time
                            ):
                                available_rooms.append(room)

                        # Перевірка чи залишилися вільні
                        if available_rooms:
                            room = random.choice(available_rooms)  # Випадковим чином
                            schedule[(group, weekday, time)] = (lector, subject, room)
                            data["lectors_info"][lector]["hours"] += 1

    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"
    ws.append(["Group", "Weekday", "Time", "Lector", "Subject", "Room"])

    for (group, weekday, time), (lector, subject, room) in schedule.items():
        ws.append([group, weekday, time, lector, subject, room])

    wb.save(excel_file)
    print(f"Schedule saved to {excel_file}")


if __name__ == '__main__':
    generate_schedule_and_save("data.json")
