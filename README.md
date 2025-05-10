# PSP-project
import datetime
import time
import threading
from plyer import notification
from openpyxl import Workbook, load_workbook
import pandas as pd
import os

# Excel file name
FILE_NAME = "tasks.xlsx"

# Dictionary to store tasks
tasks = {}

# Function to load tasks from Excel
def load_tasks():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.append(["Task ID", "Title", "Description", "Deadline", "Priority", "Status", "Notified"])
        wb.save(FILE_NAME)

    wb = load_workbook(FILE_NAME)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0]:  # Check if the row is not empty and Task ID is present
            try:
                tasks[row[0]] = {
                    "title": row[1] or '',
                    "description": row[2] or '',
                    "deadline": datetime.datetime.strptime(row[3], "%Y-%m-%d %H:%M:%S") if row[3] else datetime.datetime.now(),
                    "priority": row[4] or 'Low',
                    "status": row[5] or 'In Progress',
                    "notified": row[6] if row[6] is not None else False
                }
            except Exception as e:
                print(f"Error loading task {row}: {e}")

# Function to save tasks to Excel
def save_tasks():
    wb = Workbook()
    ws = wb.active
    ws.append(["Task ID", "Title", "Description", "Deadline", "Priority", "Status", "Notified"])
    for task_id, info in tasks.items():
        ws.append([
            task_id,
            info['title'],
            info['description'],
            info['deadline'].strftime("%Y-%m-%d %H:%M:%S"),
            info['priority'],
            info['status'],
            info['notified']
        ])
    wb.save(FILE_NAME)

# Function to add task
def add_task():
    task_id = input("Enter Task ID: ")
    title = input("Enter Task Title: ")
    description = input("Enter Description: ")
    deadline = input("Enter Deadline (YYYY-MM-DD HH:MM): ")
    priority = input("Enter Priority (High/Medium/Low): ")
    status = "In Progress"

    tasks[task_id] = {
        "title": title,
        "description": description,
        "deadline": datetime.datetime.strptime(deadline, "%Y-%m-%d %H:%M"),
        "priority": priority,
        "status": status,
        "notified": False
    }
    save_tasks()
    print("Task added successfully.\n")

# Function to view tasks
def view_tasks():
    if not tasks:
        print("No tasks available.")
    for tid, info in tasks.items():
        print(f"ID: {tid} | Title: {info['title']} | Status: {info['status']} | Deadline: {info['deadline']} | Priority: {info['priority']}")

# Function to update a task
def update_task():
    tid = input("Enter Task ID to update: ")
    if tid in tasks:
        print("Leave blank to keep current value.")
        title = input(f"New Title ({tasks[tid]['title']}): ") or tasks[tid]['title']
        description = input(f"New Description ({tasks[tid]['description']}): ") or tasks[tid]['description']
        deadline = input(f"New Deadline (YYYY-MM-DD HH:MM) ({tasks[tid]['deadline']}): ")
        priority = input(f"New Priority (High/Medium/Low) ({tasks[tid]['priority']}): ") or tasks[tid]['priority']
        status = input(f"New Status (In Progress/Completed) ({tasks[tid]['status']}): ") or tasks[tid]['status']
        
        if deadline:
            tasks[tid]['deadline'] = datetime.datetime.strptime(deadline, "%Y-%m-%d %H:%M")
        tasks[tid]['title'] = title
        tasks[tid]['description'] = description
        tasks[tid]['priority'] = priority
        tasks[tid]['status'] = status
        tasks[tid]['notified'] = False
        save_tasks()
        print("Task updated.\n")
    else:
        print("Task ID not found.\n")

# Function to delete task
def delete_task():
    tid = input("Enter Task ID to delete: ")
    if tid in tasks:
        del tasks[tid]
        save_tasks()
        print("Task deleted.\n")
    else:
        print("Task ID not found.\n")

# Function to show Excel data properly
def show_excel_data():
    try:
        df = pd.read_excel(FILE_NAME)
        print(df.to_string(index=False))
    except Exception as e:
        print(f"Error reading Excel file: {e}")

# Function to export Excel to CSV
def export_excel_to_csv():
    try:
        df = pd.read_excel(FILE_NAME)
        csv_file = FILE_NAME.replace('.xlsx', '.csv')
        df.to_csv(csv_file, index=False, encoding='utf-8')
        print(f"Excel file exported successfully to {csv_file}")
    except Exception as e:
        print(f"Error exporting Excel to CSV: {e}")

# Background thread function for notifications
def notify_due_tasks():
    while True:
        now = datetime.datetime.now()
        for task_id, info in list(tasks.items()):
            if not info['notified'] and info['status'] == "In Progress":
                if abs((info['deadline'] - now).total_seconds()) < 60:
                    notification.notify(
                        title=f"Task Due: {info['title']}",
                        message=f"{info['description']} (Priority: {info['priority']})",
                        timeout=10
                    )
                    info['notified'] = True
                    save_tasks()
        time.sleep(30)

# Start background thread
load_tasks()
notifier_thread = threading.Thread(target=notify_due_tasks, daemon=True)
notifier_thread.start()

# Main menu loop
while True:
    print("\n1. Add Task\n2. View Tasks\n3. Update Task\n4. Delete Task\n5. Exit\n6. Show Excel Data\n7. Export Excel to CSV")
    choice = input("Enter your choice: ")

    if choice == '1':
        add_task()
    elif choice == '2':
        view_tasks()
    elif choice == '3':
        update_task()
    elif choice == '4':
        delete_task()
    elif choice == '5':
        print("Exiting Task Manager.")
        break
    elif choice == '6':
        show_excel_data()
    elif choice == '7':
        export_excel_to_csv()
    else:
        print("Invalid choice. Try again.")


