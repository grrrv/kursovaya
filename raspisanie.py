import pandas as pd
from datetime import datetime, timedelta


# Загрузка данных
def load_data_with_universities(groups_path, rooms_path, subjects_path, teachers_path, universities_path):
    """
    Загружает данные из файлов Excel, включая университеты.
    """
    groups = pd.read_excel(groups_path)
    rooms = pd.read_excel(rooms_path)
    subjects = pd.read_excel(subjects_path)
    teachers = pd.read_excel(teachers_path)
    universities = pd.read_excel(universities_path)
    return groups, rooms, subjects, teachers, universities


# Проверка доступности группы, преподавателя и аудитории
def is_available(schedule, day, time, group, teacher, room):
    """
    Проверяет доступность группы, преподавателя и аудитории.
    """
    for entry in schedule:
        if entry['День'] == day and entry['Время'] == time:
            if entry['Группа'] == group or entry['Преподаватель'] == teacher or entry['Аудитория'] == room:
                return False
    return True


# Генерация расписания
def generate_schedule_with_universities(groups, rooms, subjects, teachers, universities, start_date):
    """
    Создает сбалансированное расписание с добавлением дат и распределением по университетам.
    """
    schedule = []
    days = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница']
    times = ['08:00-09:30', '09:40-11:10', '11:40-13:10', '13:30-15:00']
    subject_teacher_pairs = list(zip(subjects['Предмет'], teachers['Преподаватель']))
    day_to_date = {day: start_date + timedelta(days=i) for i, day in enumerate(days)}

    # Добавляем университеты в распределение
    university_cycle = universities['Вуз'].tolist()

    for group in groups['Группа']:
        used_subjects = set()
        for day in days:
            for time in times:
                for subject, teacher in subject_teacher_pairs:
                    if subject not in used_subjects:
                        room = rooms.sample(1)['Аудитория'].values[0]
                        university = university_cycle.pop(0)  # Выбираем университет
                        university_cycle.append(university)  # Возвращаем его в конец списка
                        if is_available(schedule, day, time, group, teacher, room):
                            schedule.append({
                                'Университет': university,
                                'Группа': group,
                                'Дата': day_to_date[day].strftime('%Y-%m-%d'),
                                'День': day,
                                'Время': time,
                                'Предмет': subject,
                                'Преподаватель': teacher,
                                'Аудитория': room,
                            })
                            used_subjects.add(subject)
                            break
                if len(used_subjects) == len(subjects):
                    used_subjects = set()
    return pd.DataFrame(schedule)


# Проверка корректности расписания
def validate_schedule(schedule):
    """
    Проверяет, что в расписании нет пересечений для группы, преподавателя и аудитории.
    """
    for _, group_schedule in schedule.groupby(['Группа', 'День', 'Время']):
        if len(group_schedule) > 1:
            print(f"Ошибка: пересечение в группе {group_schedule['Группа'].iloc[0]} "
                  f"в день {group_schedule['День'].iloc[0]} в {group_schedule['Время'].iloc[0]}.")
            return False

    for _, teacher_schedule in schedule.groupby(['Преподаватель', 'День', 'Время']):
        if len(teacher_schedule) > 1:
            print(f"Ошибка: пересечение у преподавателя {teacher_schedule['Преподаватель'].iloc[0]} "
                  f"в день {teacher_schedule['День'].iloc[0]} в {teacher_schedule['Время'].iloc[0]}.")
            return False

    for _, room_schedule in schedule.groupby(['Аудитория', 'День', 'Время']):
        if len(room_schedule) > 1:
            print(f"Ошибка: пересечение в аудитории {room_schedule['Аудитория'].iloc[0]} "
                  f"в день {room_schedule['День'].iloc[0]} в {room_schedule['Время'].iloc[0]}.")
            return False

    print("Расписание успешно проверено: пересечений нет.")
    return True


# Экспорт расписания
def export_to_excel(schedule, output_file='schedule.xlsx'):
    """
    Экспортирует расписание в Excel.
    """
    schedule.to_excel(output_file, index=False)
    print(f"Расписание сохранено в файл: {output_file}")


# Основная функция
def main(groups_path, rooms_path, subjects_path, teachers_path, universities_path):
    """
    Основная функция программы с использованием загруженных данных.
    """
    groups, rooms, subjects, teachers, universities = load_data_with_universities(
        groups_path, rooms_path, subjects_path, teachers_path, universities_path
    )

    # Укажите дату начала расписания
    start_date = datetime.strptime("2024-12-01", "%Y-%m-%d")

    # Генерация расписания
    schedule = generate_schedule_with_universities(groups, rooms, subjects, teachers, universities, start_date)

    # Проверка корректности
    if validate_schedule(schedule):
        # Экспорт расписания
        output_file = 'schedule_with_universities.xlsx'
        export_to_excel(schedule, output_file=output_file)
        return output_file
    else:
        print("Расписание содержит ошибки и не будет сохранено.")
        return None


# Запуск программы
if __name__ == "__main__":
    # Путь к данным
    groups_path = "groups.xlsx"
    rooms_path = "rooms.xlsx"
    subjects_path = "subjects.xlsx"
    teachers_path = "teachers.xlsx"
    universities_path = "universities.xlsx"

    # Запуск основной функции
    output_path = main(groups_path, rooms_path, subjects_path, teachers_path, universities_path)
    if output_path:
        print(f"Готовое расписание сохранено: {output_path}")