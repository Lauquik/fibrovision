from datetime import datetime, timedelta
import openpyxl
from openpyxl import Workbook
import os

# Function to generate Fibonacci series for 30 dates
def generate_fibonacci_series():
    fib_series = [1, 2]
    for i in range(3, 14):
        fib_series.append(fib_series[-1] + fib_series[-2])
    return fib_series

# Function to write today's learning to Excel
def add_learning(topic):
    today = datetime.now().date()
    currDate = today
    fib_series = generate_fibonacci_series()

    if not os.path.exists('learning.xlsx'):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Learning'
        sheet.append(['Date', 'Topic', 'Revision Date', 'Revision Number'])
    else:
        workbook = openpyxl.load_workbook('learning.xlsx')
        sheet = workbook['Learning']
    
    for i, days in enumerate(fib_series):
        revision_date = currDate + timedelta(days=days)
        currDate=revision_date
        sheet.append([today, topic, revision_date, i + 1])

    workbook.save('learning.xlsx')
    print("Today's learning has been added.")

# Function to show today's todo
def show_todos():
    today = datetime.now().date()
    if not os.path.exists('learning.xlsx'):
        print("No learning data found.")
        return

    workbook = openpyxl.load_workbook('learning.xlsx')
    sheet = workbook['Learning']

    todos = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        revision_date = row[2].date() if isinstance(row[2], datetime) else datetime.strptime(row[2], '%Y-%m-%d').date()
        if revision_date == today:
            todos.append((row[1], row[3]))

    if todos:
        print(f"Today's todos ({today}):")
        for todo in todos:
            print(f"Topic: {todo[0]}, Revision Number: {todo[1]}")
    else:
        print("No todos for today.")

# Main function to interact with the user
def main():
    while True:
        print("\nMenu:")
        print("1. Add today's learning")
        print("2. Show today's todo")
        print("3. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            topics_input = input("Enter today's learning topics (comma-separated): ")
            topics = [topic.strip() for topic in topics_input.split(',')]
            for topic in topics:
                add_learning(topic)
        elif choice == '2':
            show_todos()
        elif choice == '3':
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
