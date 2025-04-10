import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import threading
import time
import json
import random
import math
from pynput import mouse, keyboard
import win32gui
from pynput.mouse import Controller as MouseController

class MacroRecorderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Macro Recorder")
        self.macro = []
        self.recording = False
        self.abort_playback = False
        self.is_playing = False
        self.area_selection_mode = False
        self.area_selection_start = None
        self.area_selection_target_index = None
        self.playback_start_time = None
        self.current_loop = 0
        self.total_loops = 0
        self.ctrl_pressed = False

        self.loop_count = tk.IntVar(value=1)
        
        # Remove loop delay distribution variables
        self.tree = ttk.Treeview(root, columns=("name", "action", "delay_text", "dist", "min", "max", "click_dist", "area"), show="headings")
        self.tree.heading("name", text="Name")
        self.tree.heading("action", text="Action")
        self.tree.heading("delay_text", text="After")
        self.tree.heading("dist", text="Time Dist")
        self.tree.heading("min", text="Min")
        self.tree.heading("max", text="Max")
        self.tree.heading("click_dist", text="Click Dist")
        self.tree.heading("area", text="Area")
        
        self.tree.column("name", width=80)
        self.tree.column("action", width=130)
        self.tree.column("delay_text", width=60)
        self.tree.column("dist", width=80)
        self.tree.column("min", width=50)
        self.tree.column("max", width=50)
        self.tree.column("click_dist", width=80)
        self.tree.column("area", width=150)
        
        self.tree.pack(pady=10, fill=tk.BOTH, expand=True)

        self.context_menu = tk.Menu(self.tree, tearoff=0)
        self.context_menu.add_command(label="Delete Command", command=self.delete_command)
        self.context_menu.add_command(label="Modify Command", command=self.modify_command)
        self.context_menu.add_command(label="Edit Condition", command=self.edit_condition)
        self.context_menu.add_command(label="Edit Name", command=self.edit_action_name)
        self.context_menu.add_command(label="Modify Time Distribution", command=self.edit_time_distribution)
        self.context_menu.add_command(label="Modify Click Area", command=self.edit_click_area)
        self.context_menu.add_separator()
        
        self.tree.bind("<Double-1>", self.edit_action)
        self.tree.bind("<Button-3>", self.show_context_menu)

        control_frame = tk.Frame(root)
        control_frame.pack()

        self.record_button = tk.Button(control_frame, text="Record Macro", command=self.start_recording)
        self.record_button.grid(row=0, column=0)
        self.stop_button = tk.Button(control_frame, text="Stop Recording", command=self.stop_recording, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1)
        self.play_button = tk.Button(control_frame, text="Play Macro", command=self.play_macro)
        self.play_button.grid(row=0, column=2)
        tk.Button(control_frame, text="Save Macro", command=self.save_macro).grid(row=0, column=3)
        tk.Button(control_frame, text="Load Macro", command=self.load_macro).grid(row=0, column=4)
        self.add_button = tk.Button(control_frame, text="Add Command", command=self.add_command)
        self.add_button.grid(row=0, column=5)
        tk.Button(control_frame, text="Add Condition", command=self.add_condition).grid(row=0, column=6)
        tk.Button(control_frame, text="Move Mouse", command=self.add_move_command).grid(row=0, column=7)
        tk.Button(control_frame, text="Reset", command=self.reset_macro).grid(row=0, column=8)
        reorder_frame = tk.Frame(root)
        reorder_frame.pack()
        tk.Button(reorder_frame, text="Move Up", command=self.move_up).pack(side=tk.LEFT)
        tk.Button(reorder_frame, text="Move Down", command=self.move_down).pack(side=tk.LEFT)
        tk.Button(reorder_frame, text="Insert Timer", command=self.insert_timer).pack(side=tk.LEFT)

        loop_frame = tk.Frame(root)
        loop_frame.pack(pady=5)
        tk.Label(loop_frame, text="Loop count:").grid(row=0, column=0)
        tk.Entry(loop_frame, textvariable=self.loop_count, width=5).grid(row=0, column=1)

        status_frame = tk.Frame(root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=2)
        
        self.loop_status = tk.Label(status_frame, text="", width=20, anchor="w")
        self.loop_status.pack(side=tk.LEFT, padx=5)
        
        self.status_label = tk.Label(status_frame, text="", fg="blue")
        self.status_label.pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        self.time_status = tk.Label(status_frame, text="", width=20, anchor="e")
        self.time_status.pack(side=tk.RIGHT, padx=5)

        self.listener = keyboard.Listener(on_press=self.on_global_press, on_release=self.on_global_release)
        self.listener.start()
        self.click_counter = {"left": 0, "right": 0, "middle": 0}

    def add_move_command(self):
        self.add_button.config(state=tk.DISABLED)
        self.status_label.config(text="Click anywhere to set mouse position (ESC to cancel)")
        self.capture_command(None, is_move_command=True)  # Pass the flag here
    
    def add_condition(self):
        self.status_label.config(text="Click the pixel you want to monitor...")
        self.area_selection_mode = True
        self.condition_capture_mode = True
        self.prepare_for_pixel_selection()

    def prepare_for_pixel_selection(self):
        self.root.withdraw()
        self.overlay = tk.Toplevel(self.root)
        self.overlay.attributes("-fullscreen", True)
        self.overlay.attributes("-alpha", 0.3)
        self.overlay.configure(bg="green")
        self.overlay.attributes("-topmost", True)

        self.overlay_canvas = tk.Canvas(self.overlay, highlightthickness=0)
        self.overlay_canvas.pack(fill=tk.BOTH, expand=True)

        # Bring focus to overlay
        self.overlay.grab_set()
        self.overlay.focus_force()

        # Mouse listeners
        self.area_mouse_listener = mouse.Listener(on_click=self.on_pixel_selection_click)
        self.area_mouse_listener.start()

        # Keyboard listener
        self.area_keyboard_listener = keyboard.Listener(on_press=self.on_area_selection_key)
        self.area_keyboard_listener.start()

    def on_pixel_selection_click(self, x, y, button, pressed):
        if pressed:
            try:
                # Cleanup overlay first
                self.cancel_area_selection()

                # Get pixel color
                dc = win32gui.GetDC(0)
                color = win32gui.GetPixel(dc, x, y)
                win32gui.ReleaseDC(0, dc)

                # Convert color
                rgb = (color & 0xff, (color >> 8) & 0xff, (color >> 16) & 0xff)
                hex_color = '#%02x%02x%02x' % rgb

                # Create dialog as child of root window
                self.root.deiconify()
                condition_type = simpledialog.askstring(
                    "Condition Type",
                    "Should color match? (yes/no):",
                    parent=self.root,
                    initialvalue="yes"
                )

                if condition_type and condition_type.lower() in ["yes", "no"]:
                    event = {
                        "type": "condition",
                        "pos": (x, y),
                        "color": hex_color,
                        "match": condition_type.lower() == "yes",
                        "delay": 0.5,
                        "timeout": 60,
                        "name": f"Pixel Condition @ {x},{y}"
                    }

                    idx = self.tree.selection()
                    insert_idx = int(idx[0]) if idx else len(self.macro)
                    self.macro.insert(insert_idx, event)
                    self.refresh_tree()

            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                self.root.deiconify()

    def cancel_area_selection(self):
        if hasattr(self, 'overlay') and self.overlay:
            try:
                self.overlay.grab_release()
                self.overlay.destroy()
            except tk.TclError:
                pass
        if hasattr(self, 'area_mouse_listener'):
            self.area_mouse_listener.stop()
        if hasattr(self, 'area_keyboard_listener'):
            self.area_keyboard_listener.stop()

        self.area_selection_mode = False
        self.area_selection_start = None
        self.area_selection_target_index = None
        self.root.deiconify()

    def edit_condition(self):
        idx = self.tree.selection()
        if not idx:
            return
        index = int(idx[0])
        condition = self.macro[index]
        
        # Create edit dialog
        edit_dialog = tk.Toplevel(self.root)
        edit_dialog.title("Edit Condition Parameters")
        
        # Condition name
        tk.Label(edit_dialog, text="Condition Name:").grid(row=0, column=0)
        name_var = tk.StringVar(value=condition["name"])
        tk.Entry(edit_dialog, textvariable=name_var).grid(row=0, column=1, columnspan=2, sticky="ew")
        
        # Timeout setting
        tk.Label(edit_dialog, text="Timeout (0 for infinite):").grid(row=1, column=0)
        timeout_var = tk.DoubleVar(value=condition["timeout"])
        tk.Entry(edit_dialog, textvariable=timeout_var).grid(row=1, column=1)
        
        # Check interval
        tk.Label(edit_dialog, text="Check interval (seconds):").grid(row=2, column=0)
        interval_var = tk.DoubleVar(value=condition["delay"])
        tk.Entry(edit_dialog, textvariable=interval_var).grid(row=2, column=1)
        
        # Color match
        tk.Label(edit_dialog, text="Color match:").grid(row=3, column=0)
        match_var = tk.BooleanVar(value=condition["match"])
        tk.Radiobutton(edit_dialog, text="Yes", variable=match_var, value=True).grid(row=3, column=1)
        tk.Radiobutton(edit_dialog, text="No", variable=match_var, value=False).grid(row=3, column=2)
        
        # Position display
        x, y = condition["pos"]
        tk.Label(edit_dialog, text=f"Position: ({x}, {y})").grid(row=4, column=0, columnspan=3)
        
        # Save button
        def save_changes():
            condition["timeout"] = timeout_var.get()
            condition["delay"] = interval_var.get()
            condition["match"] = match_var.get()
            condition["name"] = name_var.get()  # Update name from entry
            self.refresh_tree()
            edit_dialog.destroy()
            
        tk.Button(edit_dialog, text="Save", command=save_changes).grid(row=5, column=0, columnspan=3)
    
        # Make dialog resizable
        edit_dialog.grid_columnconfigure(1, weight=1)

    def reset_macro(self):
        self.macro = []
        self.refresh_tree()
        self.status_label.config(text="Macro reset")

    def add_command(self):
        self.add_button.config(state=tk.DISABLED)
        self.status_label.config(text="Press any key or click anywhere to add command (ESC to cancel)")
        self.capture_command(None)

    def modify_command(self):
        idx = self.tree.selection()
        if not idx:
            return
        self.add_button.config(state=tk.DISABLED)
        self.selected_command_index = int(idx[0])
        self.status_label.config(text="Press new key/click to modify command (ESC to cancel)")
        self.capture_command(self.selected_command_index)

    def capture_command(self, modify_index, is_move_command=False):
        self.command_capture_running = True
        self.captured_event = None
        self.modify_index = modify_index
        self.is_move_command = is_move_command  # Store the flag

        self.capture_mouse_listener = mouse.Listener(on_click=self.on_capture_click)
        self.capture_keyboard_listener = keyboard.Listener(on_press=self.on_capture_press)
        self.capture_mouse_listener.start()
        self.capture_keyboard_listener.start()

        threading.Thread(target=self.monitor_command_capture).start()

    def monitor_command_capture(self):
        while self.command_capture_running:
            time.sleep(0.1)
            if self.captured_event is not None:
                self.handle_captured_command()
                break
            elif not self.capture_mouse_listener.running or not self.capture_keyboard_listener.running:
                break
        self.add_button.config(state=tk.NORMAL)

    def on_capture_click(self, x, y, button, pressed):
        if pressed and self.command_capture_running:
            if self.is_move_command:  # Check the flag here
                self.captured_event = {
                    "type": "move",
                    "pos": (x, y),
                    "delay": 0.0,
                    "dist": "fixed",
                    "min": 0.0,
                    "max": 0.0,
                    "name": f"Move to {x},{y}"
                }
            else:
                self.captured_event = {
                    "type": "click",
                    "button": str(button),
                    "pos": (x, y),
                    "delay": 0.0,
                    "dist": "fixed",
                    "min": 0.0,
                    "max": 0.0,
                    "click_dist": "fixed",
                    "area": None,
                    "name": f"Click {str(button).split('.')[-1]}"
                }
            self.command_capture_running = False

    def on_capture_press(self, key):
        """Handle key presses during command capture"""
        try:
            key_name = f"Key.{key.name}"
        except AttributeError:
            try:
                key_name = f"'{key.char}'"
            except AttributeError:
                key_name = str(key)

        if key == keyboard.Key.esc:
            self.command_capture_running = False
            self.status_label.config(text="Command capture cancelled")
            return

        self.captured_event = {
            "type": "key",
            "key": key_name,
            "delay": 0.0,
            "dist": "fixed",
            "min": 0.0,
            "max": 0.0,
            "click_dist": "fixed",
            "area": None,
            "name": f"Key {key_name.replace('Key.', '')}"
        }
        self.command_capture_running = False

    def handle_captured_command(self):
        """Process the captured command"""
        if self.captured_event:
            if self.modify_index is not None:
                # Modify existing command
                self.macro[self.modify_index] = self.captured_event
            else:
                # Add new command
                idx = self.tree.selection()
                insert_idx = int(idx[0]) if idx else len(self.macro)
                self.macro.insert(insert_idx, self.captured_event)
            
            self.refresh_tree()
            self.status_label.config(text="Command added/modified successfully")
        else:
            self.status_label.config(text="Command capture failed")

        # Cleanup listeners
        self.capture_mouse_listener.stop()
        self.capture_keyboard_listener.stop()
        self.modify_index = None
        self.captured_event = None

    def delete_command(self):
        """Delete selected command"""
        idx = self.tree.selection()
        if not idx:
            return
        index = int(idx[0])
        del self.macro[index]
        self.refresh_tree()

    def edit_action_name(self):
        """Edit the name of the selected action"""
        idx = self.tree.selection()
        if not idx:
            return
        index = int(idx[0])
        new_name = simpledialog.askstring("Edit Name", "Enter action name:", 
                                         initialvalue=self.macro[index].get("name", ""))
        if new_name is not None:  # Check if user didn't cancel
            self.macro[index]["name"] = new_name
            self.refresh_tree()

    def show_context_menu(self, event):
        """Show context menu on right-click"""
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            idx = int(iid)
            
            # Create fresh menu
            self.context_menu.delete(0, tk.END)
            
            # Add common items
            self.context_menu.add_command(label="Delete Command", command=self.delete_command)
            self.context_menu.add_command(label="Modify Command", command=self.modify_command)
            
            # Add type-specific items
            if self.macro[idx]["type"] == "condition":
                self.context_menu.add_command(label="Edit Condition", command=self.edit_condition)
            else:
                if self.macro[idx]["type"] == "click":
                    self.context_menu.add_command(label="Modify Click Area", command=self.edit_click_area)
                self.context_menu.add_command(label="Edit Name", command=self.edit_action_name)
                self.context_menu.add_command(label="Modify Time Distribution", command=self.edit_time_distribution)
            
            self.context_menu.post(event.x_root, event.y_root)
    
    def edit_time_distribution(self):
        """Edit the time distribution for the selected action"""
        idx = self.tree.selection()
        if not idx:
            return
        self.edit_action(None)  # Reuse existing edit functionality
        
    def edit_click_area(self):
        """Edit the click area for the selected action"""
        idx = self.tree.selection()
        if not idx:
            return
        index = int(idx[0])
        if self.macro[index]["type"] != "click":
            return
            
        # Ask for distribution type
        dist = simpledialog.askstring("Click Distribution", 
                                      "Enter click distribution (fixed/uniform/poisson):", 
                                      initialvalue=self.macro[index].get("click_dist", "fixed"))
        if not dist:
            return
            
        self.macro[index]["click_dist"] = dist  # Update click distribution type immediately
            
        if dist == "fixed":
            # Reset to fixed click at original position
            self.macro[index]["area"] = None
            self.refresh_tree()
        else:
            # Ask user to select area by clicking top-left and bottom-right corners
            self.area_selection_mode = True
            self.area_selection_target_index = index
            self.area_selection_start = None
            self.status_label.config(text="Please click the TOP-LEFT corner of the click area...")
            
            # Create an overlay window for area selection
            self.prepare_for_area_selection()

    def prepare_for_area_selection(self):
        """Prepare for area selection mode"""
        # Save current state
        self.root.withdraw()  # Hide the main window
        time.sleep(0.5)  # Give time for window to hide
        
        # Create a transparent overlay for selection
        self.overlay = tk.Toplevel()
        self.overlay.attributes("-fullscreen", True)
        self.overlay.attributes("-alpha", 0.3)
        self.overlay.configure(bg="blue")
        
        # Create canvas for drawing selection
        self.overlay_canvas = tk.Canvas(self.overlay, highlightthickness=0)
        self.overlay_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Mouse listeners for area selection
        self.area_mouse_listener = mouse.Listener(on_click=self.on_area_selection_click)
        self.area_mouse_listener.start()
        
        # Abort selection with Escape key
        self.area_keyboard_listener = keyboard.Listener(on_press=self.on_area_selection_key)
        self.area_keyboard_listener.start()
        self.root.update_idletasks()

    def on_area_selection_key(self, key):
        """Handle keyboard events during area selection"""
        if key == keyboard.Key.esc:
            self.root.after(100, self.cancel_area_selection)

    def cancel_area_selection(self):
        """Cancel the area selection process"""
        if hasattr(self, 'area_mouse_listener') and self.area_mouse_listener:
            self.area_mouse_listener.stop()
        if hasattr(self, 'area_keyboard_listener') and self.area_keyboard_listener:
            self.area_keyboard_listener.stop()
        if hasattr(self, 'overlay') and self.overlay:
            self.overlay.destroy()
        
        self.area_selection_mode = False
        self.area_selection_start = None
        self.area_selection_target_index = None
        self.status_label.config(text="Area selection cancelled")
        self.root.deiconify()  # Show main window again

    def on_area_selection_click(self, x, y, button, pressed):
        """Handle mouse clicks during area selection"""
        if not pressed or button != mouse.Button.left:
            return
            
        if not self.area_selection_start:
            # First click: top-left corner
            self.area_selection_start = (x, y)
            self.status_label.config(text="Now click the BOTTOM-RIGHT corner of the click area...")
            
            # Draw a marker at the first position
            self.overlay_canvas.create_oval(x-5, y-5, x+5, y+5, fill="red")
        else:
            # Second click: bottom-right corner
            top_left = self.area_selection_start
            bottom_right = (x, y)
            
            # Calculate center
            center_x = (top_left[0] + bottom_right[0]) / 2
            center_y = (top_left[1] + bottom_right[1]) / 2
            
            # Update the macro with area information
            index = self.area_selection_target_index
            click_dist = self.macro[index]["click_dist"]  # Get the already updated click_dist
            
            # Store area as dictionary with proper coordinates
            self.macro[index]["area"] = {
                "top_left": top_left,
                "bottom_right": bottom_right,
                "center": (center_x, center_y)
            }
            
            # Clean up and exit area selection mode
            self.area_mouse_listener.stop()
            self.area_keyboard_listener.stop()
            self.overlay.destroy()
            
            self.area_selection_mode = False
            self.status_label.config(text=f"Area selection completed: {top_left} to {bottom_right}")
            self.refresh_tree()
            self.root.deiconify()  # Show main window again
            
            # Stop listening
            return False

    def start_recording(self):
        self.macro = []
        self.tree.delete(*self.tree.get_children())
        self.recording = True
        self.start_time = time.time()
        self.click_counter = {"left": 0, "right": 0, "middle": 0}  # Reset click counters

        self.record_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.mouse_listener.start()
        self.keyboard_listener.start()

    def stop_recording(self):
        self.recording = False
        self.mouse_listener.stop()
        self.keyboard_listener.stop()

        if self.macro:
            self.macro.pop()
        self.refresh_tree()

        self.record_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def on_click(self, x, y, button, pressed):
        if self.recording and pressed:
            delay = time.time() - self.start_time
            self.start_time = time.time()
            
            # Determine button type and increment counter
            button_str = str(button).split(".")[-1].lower()
            if "left" in button_str:
                button_type = "left"
            elif "right" in button_str:
                button_type = "right"
            else:
                button_type = "middle"
            
            self.click_counter[button_type] += 1
            
            # Create default name for the click
            default_name = f"Click {button_type} {self.click_counter[button_type]}"
            
            event = {
                "type": "click", 
                "button": str(button), 
                "pos": (x, y), 
                "delay": delay, 
                "dist": "fixed", 
                "min": delay, 
                "max": delay,
                "click_dist": "fixed",
                "area": None,
                "name": default_name
            }
            self.macro.append(event)
            self.refresh_tree()

    def on_press(self, key):
        if self.recording:
            delay = time.time() - self.start_time
            self.start_time = time.time()
            try:
                # Store special keys differently
                if hasattr(key, 'name'):
                    key_name = f"Key.{key.name}"
                else:
                    key_name = f"'{key.char}'"  # Wrap regular chars in quotes
                
                event = {
                    "type": "key", 
                    "key": key_name, 
                    "delay": delay, 
                    "dist": "fixed", 
                    "min": delay, 
                    "max": delay,
                    "click_dist": "fixed",
                    "area": None,
                    "name": f"Key {key_name.replace('Key.', '')}"
                }
            except AttributeError:
                key_name = str(key)
                event = {
                    "type": "key", 
                    "key": key_name, 
                    "delay": delay, 
                    "dist": "fixed", 
                    "min": delay, 
                    "max": delay,
                    "click_dist": "fixed",
                    "area": None,
                    "name": f"Key {key_name}"
                }
            self.macro.append(event)
            self.refresh_tree()

    def on_global_press(self, key):
        """Handle global key presses for safe stop"""
        try:
            # Track CTRL key state
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.ctrl_pressed = True
            # Check for CTRL+ESC combination
            elif key == keyboard.Key.esc and self.ctrl_pressed:
                self.abort_playback = True
        except AttributeError:
            pass

    def on_global_release(self, key):
        """Handle global key releases for CTRL tracking"""
        try:
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.ctrl_pressed = False
        except AttributeError:
            pass

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for idx, event in enumerate(self.macro):
            if event["type"] == "move":
                action = f"Move to {event['pos']}"
                delay_text = f"{event['delay']:.2f}s"
                area_str = "N/A"
            elif event["type"] == "condition":
                action = f"Wait until pixel {event['pos']} "
                action += f"is {event['color']}" if event["match"] else f"is not {event['color']}"
                delay_text = f"{event['delay']:.2f}s"
                area_str = f"Timeout: {event['timeout']}s | Check every: {event['delay']}s"
            elif event["type"] == "click":
                action = f"Click {event['button']} at {event['pos']}"
                delay_text = f"{event['delay']:.2f}s"
                if event.get("click_dist") != "fixed" and event.get("area"):
                    area_str = f"({event['area']['top_left'][0]:.0f},{event['area']['top_left'][1]:.0f} to {event['area']['bottom_right'][0]:.0f},{event['area']['bottom_right'][1]:.0f})"
                else:
                    area_str = "N/A"
            elif event["type"] == "key":
                action = f"Press {event['key']}"
                delay_text = f"{event['delay']:.2f}s"
                area_str = "N/A"
            elif event["type"] == "delay":
                action = "Delay"
                delay_text = f"{event['delay']:.2f}s"
                area_str = "N/A"
            else:
                action = "Unknown"
                delay_text = ""
                area_str = "N/A"
                
            self.tree.insert("", "end", iid=idx, values=(
                event.get("name", ""),
                action,
                delay_text,
                event.get("dist", "fixed"),
                f"{event.get('min', 0.0):.2f}",
                f"{event.get('max', 0.0):.2f}",
                event.get("click_dist", "fixed"),
                area_str
            ))

    def edit_action(self, event):
        """Edit time distribution for an action (double-click or via context menu)"""
        idx = self.tree.selection()
        if not idx:
            return
        index = int(idx[0])
        dist = simpledialog.askstring("Time Distribution", 
                                     "Enter delay distribution (fixed/uniform/poisson):", 
                                     initialvalue=self.macro[index].get("dist", "fixed"))
        if not dist:
            return
            
        if dist == "fixed":
            val = simpledialog.askfloat("Delay", "Enter fixed delay (sec):", 
                                       initialvalue=self.macro[index].get("min", 0.5))
            if val is not None:  # Check if user didn't cancel
                self.macro[index].update({"dist": dist, "delay": val, "min": val, "max": val})
        else:
            min_val = simpledialog.askfloat("Min", "Enter minimum delay (sec):", 
                                          initialvalue=self.macro[index].get("min", 0.1))
            if min_val is not None:  # Check if user didn't cancel
                max_val = simpledialog.askfloat("Max", "Enter maximum delay (sec):", 
                                              initialvalue=self.macro[index].get("max", 1.0))
                if max_val is not None:  # Check if user didn't cancel
                    self.macro[index].update({"dist": dist, "min": min_val, "max": max_val})
        self.refresh_tree()

    def get_delay(self, event):
        dist = event.get("dist", "fixed")
        if dist == "uniform":
            return random.uniform(event.get("min", 0.1), event.get("max", 1.0))
        elif dist == "poisson":
            lam = (event.get("min", 0.1) + event.get("max", 1.0)) / 2
            return min(max(random.expovariate(1 / lam), event.get("min", 0.1)), event.get("max", 1.0))
        else:
            return event.get("delay", 0.5)
            
    def get_click_position(self, event):
        """Get click position based on distribution settings"""
        click_dist = event.get("click_dist", "fixed")
        area = event.get("area")
        
        # Return original position if fixed distribution or no area defined
        if click_dist == "fixed" or not area:
            return event["pos"]
        
        top_left = area["top_left"]
        bottom_right = area["bottom_right"]
        
        if click_dist == "uniform":
            # Uniform distribution within the area
            x = random.uniform(top_left[0], bottom_right[0])
            y = random.uniform(top_left[1], bottom_right[1])
            return (int(x), int(y))  # Convert to int for pixel coordinates
        elif click_dist == "poisson":
            # Poisson distribution around the center
            center = area["center"]
            width = bottom_right[0] - top_left[0]
            height = bottom_right[1] - top_left[1]
            
            # Use exponential distribution from center, clamped to area boundaries
            angle = random.uniform(0, 2 * math.pi)  # Random angle
            distance = random.expovariate(4 / (width + height))  # Exponential distribution
            
            # Convert polar to cartesian coords
            x = center[0] + distance * math.cos(angle)
            y = center[1] + distance * math.sin(angle)
            
            # Clamp to area boundaries
            x = max(top_left[0], min(bottom_right[0], x))
            y = max(top_left[1], min(bottom_right[1], y))
            
            return (int(x), int(y))  # Convert to int for pixel coordinates
        
        # Fallback to original position
        return event["pos"]

    def update_execution_time(self):
        """Update the execution time display"""
        if self.playback_start_time and self.is_playing:
            elapsed = time.time() - self.playback_start_time
            hours, remainder = divmod(elapsed, 3600)
            minutes, seconds = divmod(remainder, 60)
            
            if hours > 0:
                time_str = f"Time: {int(hours)}h {int(minutes)}m {seconds:.1f}s"
            elif minutes > 0:
                time_str = f"Time: {int(minutes)}m {seconds:.1f}s"
            else:
                time_str = f"Time: {seconds:.1f}s"
                
            self.time_status.config(text=time_str)
            
            # Schedule next update
            self.root.after(100, self.update_execution_time)

    def perform_human_click(self, mc, target_x, target_y, button):
        """Simulate human-like click with movement jitter and variable press duration"""
        # Add small random offset for initial movement
        offset_x = random.randint(-4, 4)
        offset_y = random.randint(-4, 4)
        intermediate_x = target_x + offset_x
        intermediate_y = target_y + offset_y

        # Move to intermediate position
        mc.position = (intermediate_x, intermediate_y)
        time.sleep(random.uniform(0.05, 0.15))  # Random delay before main movement

        # Move to actual target with sub-pixel precision
        mc.position = (target_x, target_y)
        time.sleep(random.uniform(0.03, 0.1))  # Short delay before clicking

        # Press with random delay before release
        mc.press(button)
        
        # Generate press duration using Poisson distribution (exponential)
        lam = 0.12  # Average press duration in seconds (120ms)
        press_duration = random.expovariate(1/lam)
        
        # Clamp duration to human-like range
        press_duration = max(0.05, min(press_duration, 0.3))
        time.sleep(press_duration)
        
        mc.release(button)

    def play_macro(self):
        def play():
            self.abort_playback = False
            self.is_playing = True
            self.play_button.config(state=tk.DISABLED)
            self.playback_start_time = time.time()
            self.total_loops = self.loop_count.get()
            self.current_loop = 1
            
            self.status_label.config(text="Playing macro... (Press CTRL+ESC to stop)")
            self.loop_status.config(text=f"Loop: {self.current_loop}/{self.total_loops}")
            self.update_execution_time()
            
            try:
                for loop_num in range(1, self.total_loops + 1):
                    if self.abort_playback:
                        break
                    
                    self.current_loop = loop_num
                    self.loop_status.config(text=f"Loop: {self.current_loop}/{self.total_loops}")
                    
                    for event_idx, event in enumerate(self.macro):
                        if self.abort_playback:
                            break

                        if event["type"] in ["click", "key", "delay", "move"]:    
                            event_name = event.get("name", f"Action {event_idx+1}")
                            delay = self.get_delay(event)

                            # Update status with delay information
                            self.status_label.config(text=f"Playing: {event_name} (waiting {delay:.2f}s)")
                            time.sleep(delay)

                            if event["type"] == "move":
                                x, y = event["pos"]
                                mc = MouseController()
                                mc.position = (x, y)
                                self.status_label.config(text=f"Moved to {x},{y}")
                            elif event["type"] == "click":
                                x, y = self.get_click_position(event)
                                button_str = event["button"].split(".")[-1].split(":")[0].strip(">").lower()
                                button = getattr(mouse.Button, button_str, mouse.Button.left)
                                mc = mouse.Controller()
                                self.perform_human_click(mc, x, y, button)
                                self.status_label.config(text=f"Executed: {event_name}")
                            elif event["type"] == "key":
                                key_str = event["key"]
                                if key_str.startswith("Key."):
                                    key_name = key_str.split(".", 1)[1]
                                    try:
                                        key = getattr(keyboard.Key, key_name)
                                    except AttributeError:
                                        continue
                                else:
                                    key = key_str.strip("'\"")

                                kb = keyboard.Controller()
                                try:
                                    kb.press(key)
                                    kb.release(key)
                                    self.status_label.config(text=f"Executed: {event_name}")
                                except Exception as e:
                                    self.status_label.config(text=f"Error: {str(e)}")
                        elif event["type"] == "condition":
                            self.status_label.config(
                                text=f"Checking condition: {event['name']} " +
                                (f"(Timeout in {event['timeout']}s)" if event["timeout"] > 0 else "(Infinite wait)")
                            )
                            
                            start_time = time.time()
                            condition_met = False
                            timeout = event["timeout"]
                            
                            while not condition_met and not self.abort_playback:
                                # Get current pixel color
                                x, y = event["pos"]
                                dc = win32gui.GetDC(0)
                                try:
                                    color = win32gui.GetPixel(dc, x, y)
                                finally:
                                    win32gui.ReleaseDC(0, dc)
                                
                                current_rgb = (color & 0xff, (color >> 8) & 0xff, (color >> 16) & 0xff)
                                hex_current = '#%02x%02x%02x' % current_rgb
                                
                                # Check condition
                                if event["match"]:
                                    condition_met = (hex_current.lower() == event["color"].lower())
                                else:
                                    condition_met = (hex_current.lower() != event["color"].lower())
                                
                                if condition_met:
                                    break
                                
                                # Check timeout only if specified
                                if timeout > 0 and (time.time() - start_time > timeout):
                                    self.status_label.config(
                                        text=f"Condition timeout: {event['name']}"
                                    )
                                    self.abort_playback = True
                                    break
                                
                                time.sleep(event["delay"])
                            
                            if self.abort_playback:
                                break

                    if loop_num < self.total_loops and not self.abort_playback:
                        self.status_label.config(text=f"Loop {loop_num} complete. Starting next loop...")
                        time.sleep(0.1)  # Small delay between loops
            
            except Exception as e:
                self.status_label.config(text=f"Error: {str(e)}")
                
            self.is_playing = False
            self.play_button.config(state=tk.NORMAL)
            
            if not self.abort_playback:
                elapsed = time.time() - self.playback_start_time
                self.status_label.config(text=f"Macro playback completed in {elapsed:.2f} seconds")
            else:
                self.status_label.config(text="Macro playback stopped by user")
                
            if self.playback_start_time:
                elapsed = time.time() - self.playback_start_time
                hours, remainder = divmod(elapsed, 3600)
                minutes, seconds = divmod(remainder, 60)
                time_str = (f"{int(hours)}h {int(minutes)}m {seconds:.1f}s" if hours > 0 else
                            f"{int(minutes)}m {seconds:.1f}s" if minutes > 0 else
                            f"{seconds:.1f}s")
                self.time_status.config(text=f"Total: {time_str}")
                
        threading.Thread(target=play).start()

    def get_loop_delay(self):
        dist = self.loop_delay_dist.get()
        if dist == "uniform":
            return random.uniform(self.loop_delay_min.get(), self.loop_delay_max.get())
        elif dist == "poisson":
            lam = (self.loop_delay_min.get() + self.loop_delay_max.get()) / 2
            return min(max(random.expovariate(1 / lam), self.loop_delay_min.get()), self.loop_delay_max.get())
        else:
            return self.loop_delay_min.get()

    # Update save/load methods
    def save_macro(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json")
        if file_path:
            # Convert color to lowercase for consistency
            for event in self.macro:
                if event["type"] == "condition":
                    event["color"] = event["color"].lower()
            with open(file_path, 'w') as f:
                json.dump(self.macro, f)
            messagebox.showinfo("Saved", f"Macro saved to {file_path}")

    def load_macro(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'r') as f:
                self.macro = json.load(f)

            # Convert legacy color formats
            for event in self.macro:
               if event["type"] == "condition" and "variance" not in event:
                  event["variance"] = 0  # Add default variance
               if event["type"] == "condition" and isinstance(event["color"], list):
                  event["color"] = '#%02x%02x%02x' % tuple(event["color"][:3])
               elif event["type"] == "condition":
                  event["color"] = event["color"].lower()

            # Refresh the UI to show loaded macro
            self.refresh_tree()  # Add this line

    def move_up(self):
        idx = self.tree.selection()
        if idx:
            i = int(idx[0])
            if i > 0:
                self.macro[i-1], self.macro[i] = self.macro[i], self.macro[i-1]
                self.refresh_tree()
                self.tree.selection_set(i-1)

    def move_down(self):
        idx = self.tree.selection()
        if idx:
            i = int(idx[0])
            if i < len(self.macro)-1:
                self.macro[i+1], self.macro[i] = self.macro[i], self.macro[i+1]
                self.refresh_tree()
                self.tree.selection_set(i+1)

    def insert_timer(self):
        dist = simpledialog.askstring("Distribution", "Enter delay distribution (fixed/uniform/poisson):", initialvalue="fixed")
        if not dist:
            return
            
        # Get timer name
        timer_name = simpledialog.askstring("Timer Name", "Enter timer name:", initialvalue=f"Timer {len(self.macro)+1}")
        if timer_name is None:  # User cancelled
            return
            
        if dist == "fixed":
            val = simpledialog.askfloat("Delay", "Enter delay (sec):", initialvalue=0.5)
            if val is None:  # User cancelled
                return
            timer_event = {
                "type": "delay", 
                "delay": val, 
                "dist": dist, 
                "min": val, 
                "max": val, 
                "click_dist": "fixed", 
                "area": None,
                "name": timer_name
            }
        else:
            min_val = simpledialog.askfloat("Min", "Enter minimum delay (sec):", initialvalue=0.1)
            if min_val is None:  # User cancelled
                return
            max_val = simpledialog.askfloat("Max", "Enter maximum delay (sec):", initialvalue=1.0)
            if max_val is None:  # User cancelled
                return
            timer_event = {
                "type": "delay", 
                "delay": (min_val + max_val)/2, 
                "dist": dist, 
                "min": min_val, 
                "max": max_val,
                "click_dist": "fixed", 
                "area": None,
                "name": timer_name
            }

        idx = self.tree.selection()
        insert_idx = int(idx[0]) if idx else len(self.macro)
        self.macro.insert(insert_idx, timer_event)
        self.refresh_tree()

if __name__ == "__main__":
    root = tk.Tk()
    app = MacroRecorderApp(root)
    root.mainloop()
