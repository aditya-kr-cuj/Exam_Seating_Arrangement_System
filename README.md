# Exam_Seating_Arrangement_System

This project is a **Python-based automation tool** designed to generate **exam seating plans** based on course-wise student data, room capacities, and a defined schedule. It creates **individual seating files**, **summary reports**, and allows both **dense** and **sparse** allocation strategies.

---

## Features

- Reads input from an Excel file with multiple sheets:
  - Room capacities
  - Student details
  - Course registrations
  - Timetable
- Supports **dense** and **sparse** seat allocation modes
- Generates:
  - Room-wise student seating Excel files (with invigilator & TA sections)
  - Day & session-wise seating summaries
  - A consolidated file of all allocations
- Allows **buffer seat setting** in each room
- Handles **multiple courses per session**

---

## Tech Stack

- Python
- Pandas
- OpenPyXL
- Excel Automation

---

## Input Requirements

Your input Excel file (`input_data_tt.xlsx`) must contain the following sheets:

| Sheet Name             | Description                            |
|------------------------|----------------------------------------|
| `in_room_capacity`     | Room number, exam capacity, and block  |
| `in_roll_name_mapping` | Roll numbers mapped to student names   |
| `in_course_roll_mapping` | Roll numbers and corresponding courses |
| `in_timetable`         | Exam dates and morning/evening sessions with course codes |

---

## How to Run

1. Place your `input_data_tt.xlsx` file in the same folder as the script.
2. Run the script using:

   ```bash
   python your_script_name.py
