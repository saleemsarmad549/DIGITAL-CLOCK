import customtkinter as ctk
from datetime import datetime
import pytz

ctk.set_appearance_mode("light") 
ctk.set_default_color_theme("blue") 

# Create main window
root = ctk.CTk()
root.title("DIGITAL CLOCK")
root.geometry("500x400")
root.resizable(False, False)

# ---------------- Countries ----------------
country_timezones = {
    "germany": "Europe/Berlin",
    "usa": "US/Eastern",
    "india": "Asia/Kolkata",
    "uk": "Europe/London",
    "japan": "Asia/Tokyo",
    "australia": "Australia/Sydney",
    "pakistan": "Asia/Karachi"
}

current_country = "germany"

# ---------------- Functions ----------------
def change_country():
    global current_country
    country = country_entry.get().strip().lower()
    if country in country_timezones:
        current_country = country
        country_label.configure(text=current_country.capitalize(), text_color="green")
    else:
        country_label.configure(text="Invalid Country", text_color="red")
    update_clock()  

def update_clock():
    tz_name = country_timezones.get(current_country, "Europe/Berlin")
    tz = pytz.timezone(tz_name)
    now = datetime.now(tz)
    time_string = now.strftime("%I:%M:%S %p") 
    date_string = now.strftime("%A, %d %B %Y")
    time_label.configure(text=time_string)
    date_label.configure(text=date_string)
    root.after(1000, update_clock)  


country_entry = ctk.CTkEntry(root, width=200, placeholder_text="Enter country")
country_entry.pack(pady=20)

set_button = ctk.CTkButton(root, text="Set Country", command=change_country)
set_button.pack(pady=10)

time_label = ctk.CTkLabel(root, text="", font=("Arial", 30))
time_label.pack(pady=20)

date_label = ctk.CTkLabel(root, text="", font=("Arial", 18))
date_label.pack()

country_label = ctk.CTkLabel(root, text=current_country.capitalize(), font=("Arial", 16), text_color="green")
country_label.pack(pady=10)


update_clock()
root.mainloop()
