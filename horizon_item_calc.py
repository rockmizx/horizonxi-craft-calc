import os
import json
import re
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
import time
from openpyxl import Workbook
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import importlib.util  # Substitute for 'imp'

# Paths to geckodriver and JSON
GECKO_DRIVER_PATH = os.path.join(os.path.dirname(__file__), "geckodriver.exe")
JSON_FILE = os.path.join(os.path.dirname(__file__), "items.json")
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

# Initialize Firefox with the existing profile
options = Options()

# Initialize the items_materials variable
items_materials = {}

def load_config():
    global FIREFOX_BINARY_PATH, FIREFOX_PROFILE_PATH
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config_data = json.load(f)
            FIREFOX_BINARY_PATH = config_data.get("firefox_binary_path", "")
            FIREFOX_PROFILE_PATH = config_data.get("firefox_profile_path", "")
            options.binary_location = FIREFOX_BINARY_PATH
            options.set_preference("profile", FIREFOX_PROFILE_PATH)
    else:
        messagebox.showerror("Error", f"The file {CONFIG_FILE} was not found.")

def start_browser():
    global driver
    driver = webdriver.Firefox(service=Service(GECKO_DRIVER_PATH), options=options)
    driver.get('https://horizonxi.com/account')

def load_items_from_json():
    global items_materials
    if os.path.exists(JSON_FILE):
        with open(JSON_FILE, 'r') as f:
            data = json.load(f)
            for entry in data:
                item = entry.get('item')
                materials = entry.get('materials', [])
                produced = entry.get('produced', 1)
                stack_size = entry.get('stack_size', 12)
                if item and materials:
                    items_materials[item] = {
                        "materials": materials,
                        "produced": produced,
                        "stack_size": stack_size
                    }
    else:
        messagebox.showerror("Error", f"The file {JSON_FILE} was not found.")

def populate_dropdown():
    item_var.set('')
    
    # Clear the previous menu
    item_menu['menu'].delete(0, 'end')
    
    # Add items to the menu
    for item in items_materials.keys():
        item_menu['menu'].add_command(label=item, command=lambda value=item: item_var.set(value))
    
    if items_materials:
        item_var.set(list(items_materials.keys())[0])  # Set the first item as default

def extract_numeric_value(text):
    # Use regex to extract only the numeric part of the string
    match = re.search(r"(\d+(\.\d+)?)", text.replace(',', ''))
    return float(match.group(0)) if match else None

def get_item_prices(item_name):
    url = f'https://horizonxi.com/items/{item_name.replace(" ", "_").lower()}'
    driver.get(url)
    time.sleep(3)  # Wait for the page to load

    try:
        # Get the price per unit
        price_tag = driver.find_element(By.CLASS_NAME, 'item.gil.svelte-10vtzsb')
        price_text_unit = price_tag.text.strip()
        price_unit = extract_numeric_value(price_text_unit)
        
        # Check if the "View Stacks" button is present and click it
        try:
            stack_button = driver.find_element(By.CLASS_NAME, 'button.default.default.svelte-ve8mq7')
            ActionChains(driver).move_to_element(stack_button).click(stack_button).perform()
            time.sleep(2)  # Wait for the page to update
            # Get the price per stack
            price_tag = driver.find_element(By.CLASS_NAME, 'item.gil.svelte-10vtzsb')
            price_text_stack = price_tag.text.strip()
            price_stack = extract_numeric_value(price_text_stack)
        except NoSuchElementException:
            price_stack = None

        return price_unit, price_stack
    except Exception as e:
        print(f"Error processing item {item_name}: {e}")
        return None, None

def format_price(price):
    return f"{price:,.0f}g"

def calculate(item):
    data = items_materials[item]
    materials = data["materials"]
    produced = data["produced"]
    stack_size = data["stack_size"]
    num_crafts = 12  # Considering we will do 12 crafts
    total_produced = produced * num_crafts  # Total produced after 12 crafts
    stacks_produced = total_produced // stack_size  # Divide the total produced by the stack size
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['Item', 'Materials', 'Price per Unit', 'Price per Stack', 'Total Cost'])

    item_price_unit, item_price_stack = get_item_prices(item)
    total_material_cost = 0  # Initialize the variable here
    if item_price_unit is not None:
        material_prices_str = []
        for material in materials:
            material_name = material['name']
            material_quantity = material['quantity'] * num_crafts
            material_stack_size = material.get("stack_size", 12)
            material_price_unit, material_price_stack = get_item_prices(material_name)
            if material_price_unit is not None:
                if material_price_stack is not None:
                    # Calculate the total cost of the stack
                    total_cost = material_price_stack * (material_quantity // material_stack_size)
                    if material_quantity > material_stack_size:
                        material_prices_str.append(f"{material_name}: {format_price(material_price_unit)} / {format_price(material_price_stack)} ({format_price(total_cost)})")
                    else:
                        material_prices_str.append(f"{material_name}: {format_price(material_price_unit)} / {format_price(material_price_stack)}")
                else:
                    # Material without stack, calculate stack by repeating the purchase according to the required quantity
                    total_cost = material_price_unit * material_quantity
                    if material_quantity > material_stack_size:
                        material_prices_str.append(f"{material_name}: {format_price(material_price_unit)} / {format_price(material_price_unit * material_stack_size)} ({format_price(total_cost)})")
                    else:
                        material_prices_str.append(f"{material_name}: {format_price(material_price_unit)} / {format_price(material_price_unit * material_stack_size)}")
                
                total_material_cost += total_cost
            else:
                messagebox.showwarning("Warning", f"Price for material {material_name} not found.")
                return

        total_item_price_stack = item_price_stack if item_price_stack else item_price_unit * stack_size

        materials_str = ', '.join([f"{m['name']} x{m['quantity']}" for m in materials])
        material_prices_str = '\n'.join(material_prices_str)

        # Calculate the total price per stacks
        total_price_stacks = stacks_produced * item_price_unit * stack_size
        profit = total_price_stacks - total_material_cost  # Calculate the profit
        
        sheet.append([item, materials_str, format_price(item_price_unit), format_price(total_item_price_stack), format_price(total_material_cost)])
        
        workbook.save('ff11_crafts.xlsx')
        
        # Show success message with simplified information
        messagebox.showinfo("Information", f"Calculation completed!\n\nItem: {item}\nPrice per Unit: {format_price(item_price_unit)} / {format_price(item_price_unit * stack_size)}\n\nMaterials:\n{materials_str}\n\nMaterial Prices:\n{material_prices_str}\n\nTotal Material Cost: {format_price(total_material_cost)}\n\nStacks Produced: {stacks_produced}\nTotal Price of Stacks: {format_price(total_price_stacks)}\n\nProfit: {format_price(profit)}")

def start_calculation():
    selected_item = item_var.get()
    if selected_item:
        calculate(selected_item)
    else:
        messagebox.showwarning("Warning", "Please select an item to calculate.")

# Load the configuration before starting the GUI
load_config()

# Create the GUI with Tkinter
root = tk.Tk()
root.title("Horizon XI Item Price Calculator by Rockmizx")

# Load items from JSON and populate the dropdown menu
load_items_from_json()
item_var = tk.StringVar(root)

# Dropdown menu to select the item
tk.Label(root, text="Item:").grid(row=1, column=0, padx=10, pady=10)
item_menu = ttk.OptionMenu(root, item_var, "")
item_menu.grid(row=1, column=1, padx=10, pady=10)
populate_dropdown()

# Button to start the browser
start_button = tk.Button(root, text="Start Browser", command=start_browser)
start_button.grid(row=0, column=0, padx=10, pady=10)

# Button to start the calculation
calculate_button = tk.Button(root, text="Calculate", command=start_calculation)
calculate_button.grid(row=0, column=1, padx=10, pady=10)

root.mainloop()