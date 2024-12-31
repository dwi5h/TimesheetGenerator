from datetime import datetime

def isWeekend(dateStr):
    selectedDate = datetime.strptime(dateStr, "%d-%m-%Y")
    weekno = selectedDate.weekday()
    if weekno < 5:
        return False

    return True