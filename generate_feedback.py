"""
generate_feedback.py
--------------------
Генерирует персональные сообщения для учеников на основе Excel-файла
со статистикой выполнения домашних заданий.
Excel-файл можно получить в лк преподавателя на вкладке с домашним заданием в профиле конкретного ученика
Если не указано имя ученика, подставляется просто слово "Ученик"

Параметры:
PROBLEM_THRESHOLD - порог в процентах для определения критичного уровня по теме
PARTIAL_THRESHOLD - порог в процентах для определения тем к улучшению

Использование:
    python generate_feedback.py <путь_к_файлу.xlsx> [имя_ученика]

Примеры:
    python generate_feedback.py results.xlsx
    python generate_feedback.py results.xlsx "Иван Иванов"

Выходной результат:
    Файл result.txt с сообщением для ученика
"""


# TODO: параметр для сообщений по каждому уровню, параметр для подводки, графический интерфейс?,
#  генерация плана дз в эксельке с чекбоксами, генерация для пачки файлов и учеников,
#  тг бот, чтобы было форматирование (?)
import sys
import pandas as pd
from pathlib import Path

PROBLEM_THRESHOLD = 50   # ниже — критично
PARTIAL_THRESHOLD = 80   # ниже — можно улучшить

STATUS_VISITED = "посещено"
STATUS_WATCHED = "просмотрено"
STATUS_MISSED  = "не посещено"


def load_data(filepath: str) -> pd.DataFrame:
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {filepath}")
    if path.suffix not in (".xlsx", ".xls"):
        raise ValueError(f"Ожидается .xlsx/.xls, получен: {path.suffix}")

    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip()

    col_map = {}
    for col in df.columns:
        low = col.lower()
        if "процент" in low or "%" in low:
            col_map[col] = "hw_percent"
        elif "название" in low:
            col_map[col] = "lesson_name"
        elif "статус" in low:
            col_map[col] = "attendance"
        elif "занятие" in low:
            col_map[col] = "lesson_number"
    df = df.rename(columns=col_map)

    required = {"lesson_number", "lesson_name", "attendance", "hw_percent"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Не найдены столбцы: {missing}\nДоступные: {list(df.columns)}")

    df["hw_percent"] = pd.to_numeric(df["hw_percent"], errors="coerce").fillna(0)

    # Занятия без статуса посещения — ещё не прошли, исключаем
    df = df[df["attendance"].notna()].copy()
    df["attendance"] = df["attendance"].str.strip().str.lower()
    return df


def classify_topic(row) -> str:
    status, pct = row["attendance"], row["hw_percent"]
    if status == STATUS_MISSED:      return "missed"
    if pct == 0:                     return "not_done"
    if pct < PROBLEM_THRESHOLD:      return "critical"
    if pct < PARTIAL_THRESHOLD:      return "partial"
    return "done"


def analyze(df: pd.DataFrame) -> dict:
    df = df.copy()
    df["category"] = df.apply(classify_topic, axis=1)

    priority = {"missed": 0, "not_done": 1, "critical": 2, "partial": 3, "done": 4}
    lessons_list = []
    for lesson_num, group in df.groupby("lesson_number", sort=False):
        topics = group.to_dict("records")
        worst  = min(topics, key=lambda r: priority[r["category"]])["category"]
        lessons_list.append({
            "lesson_number":  lesson_num,
            "topics":         topics,
            "worst_category": worst,
            "avg_pct":        group["hw_percent"].mean(),
        })

    def by_cat(cat):
        return [l for l in lessons_list if l["worst_category"] == cat]

    attended = df[df["attendance"].isin([STATUS_VISITED, STATUS_WATCHED])]
    avg_hw = attended["hw_percent"].mean() if len(attended) > 0 else 0.0

    return {
        "total_lessons": len(lessons_list),
        "total_topics":  len(df),
        "done_lessons":  len(by_cat("done")),
        "avg_hw":        avg_hw,
        "partial":       by_cat("partial"),
        "critical":      by_cat("critical"),
        "not_done":      by_cat("not_done"),
        "missed":        by_cat("missed"),
    }


def fmt_lesson(lesson: dict, show_percent: bool = False) -> str:
    """Занятие с темами."""
    topics, num = lesson["topics"], lesson["lesson_number"]
    lines = []

    if len(topics) == 1:
        t = topics[0]
        pct_str = f" — {int(t['hw_percent'])}%" if show_percent and t["hw_percent"] > 0 else ""
        lines.append(f"  {num}: {t['lesson_name']}{pct_str}")
    else:
        avg_str = f" (средний % — {int(lesson['avg_pct'])}%)" if show_percent else ""
        lines.append(f"  {num}:{avg_str}")
        for t in topics:
            pct_str = f" — {int(t['hw_percent'])}%" if show_percent and t["hw_percent"] > 0 else ""
            lines.append(f"      • {t['lesson_name']}{pct_str}")

    return "\n".join(lines)


def generate_message(student_name: str, analysis: dict) -> str:
    a, name = analysis, student_name or "Ученик"
    lines = [f"Привет, {name}! 👋", ""]

    # Общая картина
    done_pct = round(a["done_lessons"] / a["total_lessons"] * 100) if a["total_lessons"] else 0
    lines += [
        "📊 Общая картина",
        f"Пройдено занятий: {a['total_lessons']} (тем: {a['total_topics']})",
        f"Занятия полностью закрыты: {a['done_lessons']} из {a['total_lessons']} ({done_pct}%)",
        f"Средний % выполнения ДЗ: {round(a['avg_hw'])}%",
        "",
    ]

    has_issues = a["critical"] or a["not_done"] or a["missed"]

    if not has_issues:
        lines += ["✅ Отличная работа — все занятия закрыты на хорошем уровне!", ""]
    else:
        lines.append("⚠️ Проблемные места")

        if a["critical"]:
            lines.append(f"\n🔴 Сделано частично (менее {PROBLEM_THRESHOLD}%):")
            for l in a["critical"]:
                lines.append(fmt_lesson(l, show_percent=True))

        if a["partial"]:
            lines.append(f"\n🟡 Можно улучшить ({PROBLEM_THRESHOLD}–{PARTIAL_THRESHOLD}%):")
            for l in a["partial"]:
                lines.append(fmt_lesson(l, show_percent=True))

        if a["not_done"]:
            lines.append("\n⬜ Занятие посещено, но ДЗ не сдано:")
            for l in a["not_done"]:
                lines.append(fmt_lesson(l))

        if a["missed"]:
            lines.append("\n❌ Занятие пропущено (нужно наверстать самостоятельно):")
            for l in a["missed"]:
                lines.append(fmt_lesson(l))

        lines.append("")

    # План работы
    if has_issues:
        lines.append("📋 План работы")
        step = 1

        if a["missed"]:
            lines.append(f"\n{step}. Наверстать пропущенные занятия:")
            for l in a["missed"]:
                topics_str = ", ".join(t['lesson_name'] for t in l["topics"])
                lines.append(f"  • {l['lesson_number']}: {topics_str}")
                lines.append("    ↳ изучить материал самостоятельно, затем выполнить ДЗ")
            step += 1

        if a["critical"]:
            lines.append(f"\n{step}. Доработать сильно недоделанные ДЗ:")
            for l in a["critical"]:
                for t in l["topics"]:
                    if t["category"] in ("critical", "not_done"):
                        pct = int(t["hw_percent"])
                        action = f"сейчас {pct}%, разобрать ошибки и досдать" if pct > 0 else "не сдано — выполнить с нуля"
                        lines.append(f"  • {l['lesson_number']}: {t['lesson_name']} — {action}")
            step += 1

        if a["not_done"]:
            lines.append(f"\n{step}. Сдать ДЗ по посещённым занятиям:")
            for l in a["not_done"]:
                for t in l["topics"]:
                    if t["category"] == "not_done":
                        lines.append(f"  • {l['lesson_number']}: {t['lesson_name']} — выполнить с нуля")
            step += 1

        if a["partial"]:
            lines.append(f"\n{step}. По возможности улучшить частично выполненные задания:")
            for l in a["partial"]:
                for t in l["topics"]:
                    if t["category"] == "partial":
                        lines.append(
                            f"  • {l['lesson_number']}: {t['lesson_name']} — "
                            f"сейчас {int(t['hw_percent'])}%, попробуй поднять до 100%"
                        )

        lines.append("")

    lines.append("Если есть вопросы — пиши, разберём вместе! 💪")
    return "\n".join(lines)


def process_file(filepath: str, student_name: str = "") -> str:
    return generate_message(student_name, analyze(load_data(filepath)))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    try:
        text = process_file(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "")
        with open(f'result_{sys.argv[1]}_{sys.argv[2] if len(sys.argv) > 2 else "Ученик"}.txt',
                  'w', encoding='UTF-8') as file:
            file.write(text)
        print(text)
    except (FileNotFoundError, ValueError) as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
