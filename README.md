# Automate
# File Manager Application

A Python-based file management tool designed to clean up old files in the "Finviz" and "TradingView" folders on your Desktop.

## Features

- User prompts for file deletion confirmation.
- Deletes all but the most recent file in the specified folders.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Limitations](#limitations)
- [How it Works](#how-it-works)
- [How To guide](#how-to)
- [Code Overview](#code-overview)

## Prerequisites

- Python 3.x
- Tkinter library (included with Python)

## Installation

1. ** Download and Run the main.exe file
2. ** Success

## Usage

1. The application will record your mouse inputs and movements as well as any keyboard movements.
2. It will also move the .csv files from Tradingview and Finviz to folders on the desktop named tradinview and finviz repectively. (optional)
3. The application will also create these folders if they are not already there.
4. There is also logic to delete all but the last downloaded csv file (the most recent) if the option is chosen.
5. This application can also schedule multiple tasks to replay.

## Limitations
- The application must be running to replay tasks.
- When replaying tasks interacting with browsers, elements the tasks interact with must be in the exact same spots as they were in the recording.
- The application only manages .csv files from 2 sites.

## How it Works

- Records inputs after the button is pushed
- Saves and loads recordings as json files called "tasks"
- Will replay loaded files
 (After downloading csv files from tradingview and/or finviz)
- Locates the user's Desktop directory.
- Searches for "Finviz" and "TradingView" folders.
- Deletes all files except the most recent one in the specified folders.

## How To Guide

1. Enter a task name. Either a prerecorded or new name.
2. Load task.
3. Start recording by selecting said button.
4. Go through the movements you want recorded.
5. Stop recording by selecting said button.
6. Save recording.
7. (optional) Check the boxes for moving csv files and/or deleting files.
8. Replay task to make sure the recording was fully saved.
9. (optional) Change the replay speed with the slider, it is automatically as fast as it can go.
10. Schedule task by entering the saved task name.
11. Select the correct Year and Month for your scheduled replay.
12. Enter the time you would like the application to replay and select the proper am/pm button.
13. Enter how many times you would like to repeat the task.
14. Enter how many second you would like the application to wait before replays.
15. Keep the application open for tasks to replay.

## Code Overview

```python
# All libraries used

import shutil
import json
import threading
import time
import os
import re
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from pynput import mouse, keyboard
from selenium import webdriver
import pyautogui
from pathlib import Path
from tkcalendar import Calendar

# Beginning of code minus the class initialization and definition

# This section accesses the folders necessary for the file management system and allows chrome to download
        self.target_directory = os.path.join(os.path.expanduser("~"), "Downloads", "Finviz")
        self.chrome_options = webdriver.ChromeOptions()
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        self.chrome_options.add_experimental_option('prefs', {
            "download.default_directory": self.target_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safe browsing.enabled": True
        })

# Initializations for all of the recording and replay actions
        self.user_actions = []
        self.recording = False
        self.task_name = ""
        self.mouse_listener = None
        self.keyboard_listener = None
        self.stop_replay_flag = False
        self.pause_replay_flag = False
        self.scheduled_tasks = []

        self.stop_event = threading.Event()
        self.pause_event = threading.Event()

        self.load_scheduled_tasks()

# Start checking the schedule
        self.root.after(1000, self.check_schedule)

# File management section begins, First we move the files by checking where the folders are/making the folders if they do not exist already
    def move_csv_files(self):
        try:
            downloads_dir = self.find_downloads_directory()
            desktop_dir = Path.home() / "Desktop"
            finviz_dir = desktop_dir / "Finviz"
            tradingview_dir = desktop_dir / "TradingView"

            if not finviz_dir.exists():
                finviz_dir.mkdir(parents=True)
                print(f"Created directory: {finviz_dir}")

            if not tradingview_dir.exists():
                tradingview_dir.mkdir(parents=True)
                print(f"Created directory: {tradingview_dir}")

            for file_name in os.listdir(downloads_dir):
                if file_name.endswith(".csv") and "finviz" in file_name.lower():
                    source_path = os.path.join(downloads_dir, file_name)
                    destination_path = os.path.join(finviz_dir, file_name)
                    try:
                        shutil.move(source_path, destination_path)
                        print(f"Moved {file_name} to {finviz_dir}")
                    except Exception as e:
                        print(f"Failed to move {file_name}: {e}")

            current_date = datetime.now().strftime("%Y-%m-%d")
            stock_screener_pattern = re.compile(rf"stock screener_\d{{4}}-\d{{2}}-\d{{2}}\.csv$", re.IGNORECASE)

            for file_name in os.listdir(downloads_dir):
                if stock_screener_pattern.match(file_name):
                    source_path = os.path.join(downloads_dir, file_name)
                    destination_path = os.path.join(tradingview_dir, file_name)
                    try:
                        shutil.move(source_path, destination_path)
                        print(f"Moved {file_name} to {tradingview_dir}")
                    except Exception as e:
                        print(f"Failed to move {file_name}: {e}")

            self.delete_old_files(finviz_dir)
            self.delete_old_files(tradingview_dir)

        except Exception as e:
            print(f"An error occurred in move_csv_files: {e}")

    def get_desktop_path(self):
        """ Return the path to the user's Desktop directory. """
        return Path(os.path.expanduser("~/Desktop"))

    def find_target_folders(self):
        """ Find 'Finviz' and 'TradingView' folders on the Desktop. """
        desktop_path = self.get_desktop_path()
        target_folders = ["Finviz", "TradingView"]
        found_folders = {}

        for folder in target_folders:
            folder_path = desktop_path / folder
            if folder_path.is_dir():
                found_folders[folder] = folder_path

        return found_folders

# File deletion section. Deletes all but the most recent file or the last downloaded
    def delete_old_files(self, directory):
        """ Delete all but the most recent files in the given directory. """
        try:
            directory = Path(directory)
            files = [f for f in directory.iterdir() if f.is_file()]
            if files:
                files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                files_to_delete = files[1:]  # All but the most recent file
                for file in files_to_delete:
                    try:
                        file.unlink()
                        print(f"Deleted old file: {file}")
                    except Exception as e:
                        print(f"Failed to delete {file}: {e}")
        except Exception as e:
            print(f"An error occurred in delete_old_files: {e}")

    def prompt_delete_old_files(self):
        """ Prompt the user to delete old files in target folders. """
        result = messagebox.askyesno("Delete Old Files",
                                     "Do you want to delete all but the most recent files in 'Finviz' and 'TradingView' folders on the Desktop?")
        if result:
            try:
                target_folders = self.find_target_folders()
                if target_folders:
                    for folder_name, folder_path in target_folders.items():
                        self.delete_old_files(folder_path)
                else:
                    messagebox.showwarning("No Folders Found",
                                           "No 'Finviz' or 'TradingView' folders found on the Desktop.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def find_downloads_directory(self):
        # Platform-independent method to find the downloads directory
        downloads_dir = str(Path.home() / "Downloads")
        return downloads_dir

    def on_task_completion(self):
        # Move CSV files from downloads directory to Finviz directory
        self.move_csv_files()

        # Other actions to perform after task completion
        print("Task completion actions executed.")
# End of file management initialization and methods

# The GUI setup section
    def setup_gui(self):
        self.root.configure(background='#34495e')  # Navy blue background
        padx, pady = 5, 5
        font = ('Arial', 10)

        # Frames for better organization
        task_frame = tk.Frame(self.root, bg='#34495e')
        task_frame.grid(row=0, column=0, padx=padx, pady=pady, sticky='n')

        schedule_frame = tk.Frame(self.root, bg='#34495e')
        schedule_frame.grid(row=0, column=1, padx=padx, pady=pady, sticky='n')

        # Task Name Entry
        tk.Label(task_frame, text="Task Name:", font=font, bg='#34495e', fg='white').grid(row=0, column=0, padx=padx,
                                                                                          pady=pady, sticky='e')
        self.task_name_entry = tk.Entry(task_frame, width=30, font=font)
        self.task_name_entry.grid(row=0, column=1, padx=padx, pady=pady, columnspan=2, sticky='w')

        # Checkbox for moving files
        self.move_files_var = tk.BooleanVar()
        self.move_files_checkbox = tk.Checkbutton(task_frame, text="Move CSV Files", variable=self.move_files_var,
                                                  font=font, bg='#34495e', fg='white', selectcolor='#2c3e50')
        self.move_files_checkbox.grid(row=1, column=1, columnspan=2, padx=padx, pady=pady, sticky='w')
        # Add the "Delete Old Files" Button
        tk.Button(task_frame, text="Delete Old Files", command=self.prompt_delete_old_files, bg="#2c3e50", fg="white",
                  font=font).grid(
            row=1, column=2, padx=padx, pady=pady)

        # Recording Controls
        tk.Button(task_frame, text="Start Recording", command=self.start_recording_task, bg="#2c3e50", fg="white",
                  font=font).grid(row=2, column=0, padx=padx, pady=pady)
        tk.Button(task_frame, text="Stop Recording", command=self.stop_recording_task, bg="#2c3e50", fg="white",
                  font=font).grid(row=2, column=1, padx=padx, pady=pady)
        tk.Button(task_frame, text="Save Recording", command=self.save_recording, bg="#34495e", fg="white",
                  font=font).grid(row=2, column=2, padx=padx, pady=pady)
        tk.Button(task_frame, text="Load Recording", command=self.load_recording, bg="#34495e", fg="white",
                  font=font).grid(row=2, column=3, padx=padx, pady=pady)
        tk.Button(task_frame, text="Delete Task", command=self.delete_task, bg="#2c3e50", fg="white", font=font).grid(
            row=2, column=4, padx=padx, pady=pady)

        # Recorded Actions Display
        tk.Label(task_frame, text="Recorded Actions:", font=font, bg='#34495e', fg='white').grid(row=3, column=0,
                                                                                                 columnspan=5,
                                                                                                 padx=padx, pady=pady)
        self.actions_text = scrolledtext.ScrolledText(task_frame, width=50, height=10, font=font)
        self.actions_text.grid(row=4, column=0, columnspan=5, padx=padx, pady=pady)

        # Replay Controls
        tk.Button(task_frame, text="Replay Actions", command=self.replay_actions, bg="#34495e", fg="white",
                  font=font).grid(row=5, column=0, padx=padx, pady=pady)
        tk.Button(task_frame, text="Pause/Resume Replay", command=self.pause_replay, bg="#2c3e50", fg="white",
                  font=font).grid(row=5, column=1, padx=padx, pady=pady)
        tk.Button(task_frame, text="Stop Replay", command=self.stop_replay, bg="#2c3e50", fg="white", font=font).grid(
            row=5, column=2, padx=padx, pady=pady)

        # Replay Speed Control
        tk.Label(task_frame, text="Replay Speed (ms delay):", font=font, bg='#34495e', fg='white').grid(row=6, column=0,
                                                                                                        padx=padx,
                                                                                                        pady=pady,
                                                                                                        sticky='e')
        self.replay_speed = tk.Scale(task_frame, from_=0, to=1000, orient='horizontal', font=font)
        self.replay_speed.set(100)
        self.replay_speed.grid(row=6, column=1, padx=padx, pady=pady, columnspan=2, sticky='w')

        # Progress Bar
        tk.Label(task_frame, text="Task Progress:", font=font, bg='#34495e', fg='white').grid(row=7, column=0,
                                                                                              padx=padx, pady=pady,
                                                                                              sticky='e')
        self.progress = ttk.Progressbar(task_frame, orient="horizontal", length=200, mode="determinate")
        self.progress.grid(row=7, column=1, columnspan=2, padx=padx, pady=pady, sticky='w')

        # Scheduled Tasks Display
        tk.Label(task_frame, text="Scheduled Tasks:", font=font, bg='#34495e', fg='white').grid(row=8, column=0,
                                                                                                columnspan=5, padx=padx,
                                                                                                pady=pady)
        self.scheduled_tasks_text = scrolledtext.ScrolledText(task_frame, width=50, height=5, font=font)
        self.scheduled_tasks_text.grid(row=9, column=0, columnspan=5, padx=padx, pady=pady)

        # Schedule Task Section
        tk.Label(schedule_frame, text="Schedule Date:", font=font, bg='#34495e', fg='white').grid(row=0, column=0,
                                                                                                  padx=padx, pady=pady,
                                                                                                  sticky='w')
        self.calendar = Calendar(schedule_frame, selectmode='day', year=2023, month=7, day=1)
        self.calendar.grid(row=1, column=0, padx=padx, pady=pady, rowspan=5, columnspan=2)

        tk.Label(schedule_frame, text="Set Time:", font=font, bg='#34495e', fg='white').grid(row=6, column=0, padx=padx,
                                                                                             pady=pady, sticky='w')
        self.hour_slider = tk.Scale(schedule_frame, from_=1, to=12, orient='horizontal', font=font)
        self.hour_slider.set(12)
        self.hour_slider.grid(row=6, column=1, padx=padx, pady=pady)

        self.minute_slider = tk.Scale(schedule_frame, from_=0, to=59, orient='horizontal', font=font)
        self.minute_slider.set(0)
        self.minute_slider.grid(row=7, column=1, padx=padx, pady=pady)

        self.am_pm_var = tk.StringVar(value='AM')
        self.am_radio = tk.Radiobutton(schedule_frame, text="AM", variable=self.am_pm_var, value='AM', font=font,
                                       bg='#E6F7FF', fg='black')
        self.am_radio.grid(row=8, column=0, padx=padx, pady=pady)
        self.pm_radio = tk.Radiobutton(schedule_frame, text="PM", variable=self.am_pm_var, value='PM', font=font,
                                       bg='#E6F7FF', fg='black')
        self.pm_radio.grid(row=8, column=1, padx=padx, pady=pady)

        tk.Label(schedule_frame, text="Repeat Times:", font=font, bg='#34495e', fg='white').grid(row=9, column=0,
                                                                                                 padx=padx, pady=pady,
                                                                                                 sticky='e')
        self.repeat_entry = tk.Entry(schedule_frame, width=10, font=font)
        self.repeat_entry.grid(row=9, column=1, padx=padx, pady=pady, sticky='w')

        tk.Label(schedule_frame, text="Delay Between Repeats (seconds):", font=font, bg='#34495e', fg='white').grid(
            row=10, column=0, padx=padx, pady=pady, sticky='e')
        self.delay_entry = tk.Entry(schedule_frame, width=10, font=font)
        self.delay_entry.grid(row=10, column=1, padx=padx, pady=pady, sticky='w')

        tk.Button(schedule_frame, text="Schedule Task",
                  command=lambda: self.schedule_task()).grid(row=11, column=0, columnspan=2, padx=padx, pady=pady)
        # Tally Display Frame
        tally_frame = tk.Frame(self.root, bg='#34495e')
        tally_frame.grid(row=1, column=0, padx=padx, pady=pady, sticky='n')

        tk.Label(tally_frame, text="Task Repeat Tally:", font=font, bg='#34495e', fg='white').grid(row=0, column=0,
                                                                                                   padx=padx, pady=pady)

        # Create a canvas for tally display
        self.tally_canvas = tk.Canvas(tally_frame, width=300, height=100, bg='#34495e', highlightthickness=0)
        self.tally_canvas.grid(row=1, column=0, padx=padx, pady=pady)

# Functions and methods. Make the buttons work

# Methods to record inputs
    def start_recording_task(self):
        self.task_name = self.task_name_entry.get()
        if not self.task_name:
            messagebox.showwarning("Warning", "Please enter a task name.")
            return

        self.recording = True
        self.user_actions = []
        self.actions_text.delete(1.0, tk.END)
        self.actions_text.insert(tk.END, "Recording started...\n")

        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()

        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.keyboard_listener.start()

    def stop_recording_task(self):
        self.recording = False
        if self.mouse_listener:
            self.mouse_listener.stop()
        if self.keyboard_listener:
            self.keyboard_listener.stop()
        self.actions_text.insert(tk.END, "Recording stopped.\n")

# Method to save the recordings as json files so that other methods can utilize them. Important to keep in mind for scheduling
    def save_recording(self):
        if not self.user_actions:
            messagebox.showwarning("Warning", "No actions recorded to save.")
            return
        if not self.task_name:
            messagebox.showwarning("Warning", "Please enter a task name.")
            return
        directory = "recordings"
        if not os.path.exists(directory):
            os.makedirs(directory)
        file_path = os.path.join(directory, f"{self.task_name}.json")
        with open(file_path, 'w') as f:
            json.dump(self.user_actions, f, indent=4)
        messagebox.showinfo("Info", f"Task saved as {file_path}")

    def load_recording(self):
        self.task_name = self.task_name_entry.get().strip()
        if not self.task_name:
            messagebox.showwarning("Warning", "Please enter a task name.")
            return
        file_path = os.path.join("recordings", f"{self.task_name}.json")
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"No recording found for task name: {self.task_name}")
            return
        with open(file_path, 'r') as f:
            self.user_actions = json.load(f)
        self.update_actions_display()
        messagebox.showinfo("Info", f"Loaded task from {file_path}")

    def delete_task(self):
        self.task_name = self.task_name_entry.get().strip()
        if not self.task_name:
            messagebox.showwarning("Warning", "Please enter a task name.")
            return
        file_path = os.path.join("recordings", f"{self.task_name}.json")
        if os.path.exists(file_path):
            os.remove(file_path)
            messagebox.showinfo("Info", f"Deleted task {self.task_name}")
        else:
            messagebox.showerror("Error", f"No recording found for task name: {self.task_name}")

    def update_actions_display(self):
        self.actions_text.delete('1.0', tk.END)
        for action in self.user_actions:
            if 'type' in action:
                if action["type"] == "mouse_click":
                    self.actions_text.insert(tk.END,
                                             f"Mouse click at ({action['x']}, {action['y']}) - {action['button']} {'pressed' if action.get('pressed', False) else 'released'}\n")
                elif action["type"] == "key_press":
                    self.actions_text.insert(tk.END, f"Key pressed: {action['key']}\n")
            else:
                self.actions_text.insert(tk.END, f"Unknown action format: {action}\n")

# Recording methods for each possible movement. Limited for speed.
    def on_click(self, x, y, button, pressed):
        if self.recording:
            button_name = button.name  # Convert mouse.Button object to string
            action_type = "mouseDown" if pressed else "mouseUp"
            action = {"action": action_type, "x": x, "y": y, "button": button_name, "time": time.time()}
            self.user_actions.append(action)
            self.actions_text.insert(tk.END, f"{action}\n")

    def on_press(self, key):
        if self.recording:
            try:
                key_name = key.char
            except AttributeError:
                key_name = str(key)
            action = {"action": "keyPress", "key": key_name, "time": time.time()}
            self.user_actions.append(action)
            self.actions_text.insert(tk.END, f"{action}\n")

# replay_actions: Checks if there are actions to replay, initializes replay state, and starts a new thread to perform the replay.
    def replay_actions(self):
        if not self.user_actions:
            messagebox.showwarning("Warning", "No recorded actions to replay.")
            return
        self.stop_replay_flag = False
        self.pause_replay_flag = False
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.user_actions)
        threading.Thread(target=self._replay_actions).start()

        # Move CSV files after replaying actions
        if self.move_files_var.get():
            self.move_csv_files()

# _replay_actions: Iterates through recorded actions and performs each action with a delay.
    def _replay_actions(self):
        for index, action in enumerate(self.user_actions):
            if self.stop_replay_flag:
                break
            while self.pause_replay_flag:
                time.sleep(0.1)
            action_type = action["action"]
            delay = self.replay_speed.get() / 1000
            if action_type == "mouseDown" or action_type == "mouseUp":
                x, y = action["x"], action["y"]
                button = action["button"]
                self.perform_mouse_click(action_type, x, y, button)
            elif action_type == "keyPress":
                key = action["key"]
                pyautogui.press(key)
            time.sleep(delay)
            self.progress["value"] = index + 1

    def perform_mouse_click(self, action_type, x, y, button):
        button_str = button  # Button is already in string format
        pyautogui.moveTo(x, y)
        if action_type == "mouseDown":
            pyautogui.mouseDown(button=button_str)
        elif action_type == "mouseUp":
            pyautogui.mouseUp(button=button_str)

    def pause_replay(self):
        self.pause_replay_flag = not self.pause_replay_flag
        if self.pause_replay_flag:
            self.actions_text.insert(tk.END, "Replay paused.\n")
        else:
            self.actions_text.insert(tk.END, "Replay resumed.\n")

    def stop_replay(self):
        self.stop_replay_flag = True
        self.actions_text.insert(tk.END, "Replay stopped.\n")

        # Move CSV files if checkbox is checked
        if self.move_files_var.get():
            self.move_csv_files()

# Methods for scheduling
    def schedule_task(self):
        try:
            # Retrieve the selected date and time from the UI
            selected_date = self.calendar.get_date()
            hour = self.hour_slider.get()
            minute = self.minute_slider.get()
            am_pm = self.am_pm_var.get()

            # Convert 12-hour time format to 24-hour format
            if am_pm == 'PM' and hour != 12:
                hour += 12
            elif am_pm == 'AM' and hour == 12:
                hour = 0

            # Combine date and time
            scheduled_time_str = f"{selected_date} {hour:02}:{minute:02}:00"

            # Parse the combined string into a datetime object
            scheduled_time = datetime.strptime(scheduled_time_str, '%m/%d/%y %H:%M:%S')

            # Get repeat times and delay
            repeat_times = int(self.repeat_entry.get())
            delay_between_repeats = int(self.delay_entry.get())

            # Retrieve the task name and actions
            task_name = self.task_name_entry.get()
            actions = self.user_actions  # Assuming you want to use the currently recorded actions

            # Add the scheduled task
            self.scheduled_tasks.append({
                'task_name': task_name,
                'actions': actions,
                'time': scheduled_time,
                'repeat': repeat_times,
                'delay': delay_between_repeats
            })

            # Update the scheduled tasks display
            self.update_scheduled_tasks_text()

        except Exception as e:
            print(f"Failed to schedule task: {e}")

    def save_scheduled_tasks(self):
        tasks_to_save = []
        for task in self.scheduled_tasks:
            task_copy = task.copy()
            task_copy['time'] = task_copy['time'].isoformat()
            tasks_to_save.append(task_copy)
        with open("scheduled_tasks.json", 'w') as file:
            json.dump(tasks_to_save, file, indent=4)

# Keep this section in mind if you save tasks in a format other than json
    def load_scheduled_tasks(self):
        try:
            with open("scheduled_tasks.json", 'r') as file:
                content = file.read()
                if not content:
                    self.scheduled_tasks = []
                    return

                tasks_from_file = json.loads(content)
                self.scheduled_tasks = []
                for task in tasks_from_file:
                    task_copy = task.copy()
                    task_copy['time'] = datetime.fromisoformat(task_copy['time'])
                    self.scheduled_tasks.append(task_copy)
                self.update_scheduled_tasks_text()
        except FileNotFoundError:
            self.scheduled_tasks = []
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON: {e}")
            self.scheduled_tasks = []

# Visual 1 for task replays
    def update_scheduled_tasks_text(self):
        self.scheduled_tasks_text.delete(1.0, tk.END)
        for task in self.scheduled_tasks:
            # Format the datetime object directly
            task_time = task["time"].strftime("%Y-%m-%d %H:%M:%S")
            self.scheduled_tasks_text.insert(tk.END,
                                             f"Task: {task['task_name']}, Scheduled Time: {task_time}, Repeat: {task['repeat']} times, Delay: {task['delay']} seconds\n")

    def check_schedule(self):
        now = datetime.now()  # Get current time as a datetime object
        for task in self.scheduled_tasks:
            if now >= task["time"]:  # Compare datetime objects
                threading.Thread(target=self.execute_scheduled_task, args=(task,)).start()
                self.scheduled_tasks.remove(task)
        self.save_scheduled_tasks()
        self.update_scheduled_tasks_text()
        self.root.after(1000, self.check_schedule)

        # Move CSV files if checkbox is checked
        if self.move_files_var.get():
            self.move_csv_files()

    def execute_scheduled_task(self, task):
        print(f"Executing task: {task['task_name']}")  # Debug print
        if 'actions' in task:
            self.user_actions = task["actions"]
        else:
            print(f"Task does not contain 'actions': {task}")
            return
        repeat_times = task["repeat"]
        delay_between_repeats = task["delay"]
        for i in range(repeat_times):
            try:
                self.replay_actions()
                self.update_tally(i, 'success')  # Mark this repeat as successful
            except Exception as e:
                print(f"Error executing task: {e}")
                self.update_tally(i, 'error')  # Mark this repeat as failed
            time.sleep(delay_between_repeats)

# Visualization 2 for replayed scheduled task so the user can tell if all repeats were successful
    def update_tally(self, repeat_index, status):
        # Define colors
        color = 'green' if status == 'success' else 'red'

        # Calculate position
        x = 10 + repeat_index * 30
        y = 50

        # Draw or update the bubble
        radius = 20
        self.tally_canvas.create_oval(x - radius, y - radius, x + radius, y + radius, fill=color, outline='black')
# End of code minus the main execution logic and run block
