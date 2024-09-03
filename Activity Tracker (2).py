import tkinter as tk
from tkinter import ttk
import datetime
import time
import threading
import win32gui
import pandas as pd
import win32process
import psutil
import os

class TaskManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Active Window Tracker")
        self.root.geometry("600x400")
        self.root.configure(bg="black")

        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Treeview", 
                        background="black",  
                        foreground="cyan",  
                        fieldbackground="black",
                        font=("Helvetica", 10))

        style.configure("Treeview.Heading", 
                        background="black",  
                        foreground="white",  
                        font=("Helvetica", 10, "bold"))

        self.tree = ttk.Treeview(root, columns=("Window", "Time Spent"), show="headings", height=15)
        self.tree.heading("Window", text="Window Name")
        self.tree.heading("Time Spent", text="Time Spent")
        self.tree.column("Window", width=400)
        self.tree.column("Time Spent", width=150)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.data = {}
        self.start_times = {}
        self.window_sessions = {}

        self.history_button = tk.Button(root, text="App History", bg="black", fg="cyan", font=("Helvetica", 10, "bold"), command=self.show_history)
        self.history_button.pack(pady=10)

        self.current_window = None
        self.update_treeview()
        self.refresh()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.periodic_save_interval = 30  
        self.last_save_time = time.time()
        self.root.after(1000, self.check_periodic_save)  

    def log_activity(self, window_name, start_time, end_time):
        duration = end_time - start_time

        if window_name in self.data:
            self.data[window_name] += duration
        else:
            self.data[window_name] = duration

        if window_name in self.window_sessions:
            self.window_sessions[window_name]["Time Spent"] += duration
        else:
            self.window_sessions[window_name] = {
                "Start Time": datetime.datetime.fromtimestamp(start_time).strftime('%Y-%m-%d %H:%M:%S'),
                "Time Spent": duration
            }

    def update_treeview(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        for window_name, total_time in self.data.items():
            total_time_str = str(datetime.timedelta(seconds=total_time))
            self.tree.insert("", tk.END, values=(window_name, total_time_str))

    def refresh(self):
        current_window = get_active_window()

        if current_window != self.current_window:
            if self.current_window:
                self.log_activity(self.current_window, self.start_times[self.current_window], time.time())

            self.current_window = current_window
            self.start_times[self.current_window] = time.time()

        if self.current_window:
            elapsed_time = time.time() - self.start_times[self.current_window]
            self.data[self.current_window] = self.data.get(self.current_window, 0) + elapsed_time
            self.start_times[self.current_window] = time.time()

        self.update_treeview()
        self.root.after(1000, self.refresh)

    def show_history(self):
        self.root.withdraw()

        history_window = tk.Toplevel(self.root)
        history_window.title("App History")
        history_window.geometry("600x400")
        history_window.configure(bg="black")

        history_tree = ttk.Treeview(history_window, columns=("Window", "Start Time", "Time Spent"), show="headings", height=15)
        history_tree.heading("Window", text="Window Name")
        history_tree.heading("Start Time", text="Start Time")
        history_tree.heading("Time Spent", text="Time Spent")
        history_tree.column("Window", width=300)
        history_tree.column("Start Time", width=150)
        history_tree.column("Time Spent", width=150)
        history_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Treeview", 
                        background="black",  
                        foreground="cyan",  
                        fieldbackground="black",
                        font=("Helvetica", 10))

        style.configure("Treeview.Heading", 
                        background="black",  
                        foreground="white",  
                        font=("Helvetica", 10, "bold"))

        for window_name, session_data in self.window_sessions.items():
            total_time_str = str(datetime.timedelta(seconds=session_data["Time Spent"]))
            history_tree.insert("", tk.END, values=(window_name, session_data["Start Time"], total_time_str))

        back_button = tk.Button(history_window, text="Back", bg="black", fg="cyan", font=("Helvetica", 10, "bold"), command=lambda: self.go_back(history_window))
        back_button.pack(pady=10)

    def go_back(self, history_window):
        history_window.destroy()
        self.root.deiconify()

    def on_closing(self):
        self.save_to_excel()
        self.root.destroy()

    def save_to_excel(self):
        try:
            # r"C:\Users\lykha\Downloads\Activity_Tracker (2).py"
            folder_path = r"data"
            os.makedirs(folder_path, exist_ok=True)

            file_path = os.path.join(folder_path, "window_tracking_data.xlsx")
            print(f"Saving to: {file_path}")

            df = pd.DataFrame([
                {"Window": window_name, "Time Spent": str(datetime.timedelta(seconds=time_spent))}
                for window_name, time_spent in self.data.items()
            ])
          
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Window Tracking Data')
                worksheet = writer.sheets['Window Tracking Data']
                
                for i, col in enumerate(df.columns):
                    column_len = df[col].astype(str).str.len().max()
                    column_len = max(column_len, len(col))
                    worksheet.column_dimensions[chr(65 + i)].width = column_len

            history_file_path = os.path.join(folder_path, "app_history.xlsx")
            print(f"Saving history to: {history_file_path}")

            history_df = pd.DataFrame([
                {"Window": window_name, "Start Time": session_data["Start Time"], "Time Spent": str(datetime.timedelta(seconds=session_data["Time Spent"]))}
                for window_name, session_data in self.window_sessions.items()
            ])
            
            with pd.ExcelWriter(history_file_path, engine='openpyxl') as writer:
                history_df.to_excel(writer, index=False, sheet_name='App History')
                worksheet = writer.sheets['App History']
                
                for i, col in enumerate(history_df.columns):
                    column_len = history_df[col].astype(str).str.len().max()
                    column_len = max(column_len, len(col))
                    worksheet.column_dimensions[chr(65 + i)].width = column_len

            print("Data saved successfully.")

        except Exception as e:
            print(f"Error saving data to Excel: {e}")

    def check_periodic_save(self):
        current_time = time.time()
        if current_time - self.last_save_time >= self.periodic_save_interval:
            print("Periodic save triggered...")
            self.save_to_excel()
            self.last_save_time = current_time
        self.root.after(1000, self.check_periodic_save) 

def track_time(app, stop_event):
    previous_window = None
    previous_start_time = None
    
    while not stop_event.is_set():
        try:
            time.sleep(1)
            current_window = get_active_window()
            current_time = time.time()

            if current_window:
                if previous_window and previous_window != current_window:
                    if previous_start_time:
                        app.log_activity(previous_window, previous_start_time, current_time)

                if current_window not in app.start_times:
                    app.start_times[current_window] = current_time
                    if previous_window:
                        if previous_start_time:
                            app.log_activity(previous_window, previous_start_time, current_time)

                previous_window = current_window
                previous_start_time = app.start_times[current_window] if current_window in app.start_times else None

        except Exception as e:
            print(f"Error in tracking thread: {e}")

def get_active_window():
    try:
        window = win32gui.GetForegroundWindow()
        window_name = win32gui.GetWindowText(window)

        if window_name in ["Active Window Tracker", "App History"]:
            return None

        thread_id, process_id = win32process.GetWindowThreadProcessId(window)
        if process_id <= 0:
            return "Unknown Window"

        process_name = psutil.Process(process_id).name()

        if process_name in ["chrome.exe", "firefox.exe", "msedge.exe"]:
            return window_name

        return window_name
    except Exception as e:
        print(f"Error getting active window: {e}")
        return "Unknown Window"

def main():
    root = tk.Tk()
    app = TaskManagerApp(root)

    stop_event = threading.Event()

    tracking_thread = threading.Thread(target=track_time, args=(app, stop_event))
    tracking_thread.daemon = True
    tracking_thread.start()

    app.refresh()
    root.mainloop()

    stop_event.set()
    tracking_thread.join()

if __name__ == "__main__":
    try:
        print("Tracking started... Close the window to stop.")
        main()
    except KeyboardInterrupt:
        print("\nTracking stopped.")
