import configparser
import datetime
import json
import logging
import os
import shutil
import sys
import threading
import time
import traceback
import enum
from functools import wraps
from tkinter import PhotoImage

from collections import Counter
from logging.handlers import RotatingFileHandler
from tkinter import Tk, Button, Frame, Label, Entry, StringVar, IntVar, Toplevel, Canvas
from tkinter import ttk as tkttk

import autoit
import cv2
import difflib
import easyocr
import glob
import keyboard
import numpy as np
import pygetwindow as gw
import pyperclip
import requests
from PIL import ImageGrab, Image
from ttkbootstrap import Style

CONFIG_FILE = 'config.ini'
SSR_TEMPLATE_PATH = os.path.join('images_scans', 'ssr.png')

class ConfigSection(enum.Enum):
    WEBHOOK = 'WEBHOOK'
    REROLL_SETTINGS = 'REROLL_SETTINGS'
    SSR_REROLL = 'SSR_REROLL'
    INGAME_MENU = 'INGAME_MENU'
    SAFE_CHECK = 'SAFE_CHECK'
    CARDS = 'CARDS'
    REGISTER_ACCOUNT = 'REGISTER_ACCOUNT'
    MACRO_REROLL = 'MACRO_REROLL'
    LINK_ACCOUNT = 'LINK_ACCOUNT'

class WebhookStatus(enum.Enum):
    START = 'start'
    STOP = 'stop'
    SUMMARY = 'summary'
    LINKED = 'linked'
    REGISTER = 'register'
    CARATS = 'carats'
    SCOUTMENU = 'scoutmenu'
    RUNSTART = 'runstart'

# Decorator for macro step methods
def macro_active_only(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if not getattr(self, 'is_macro_started', True):
            return
        return func(self, *args, **kwargs)
    return wrapper

def create_default_config():
    config = configparser.ConfigParser()
    config.optionxform = str
    config['WEBHOOK'] = {
        'url': 'YOUR_DISCORD_WEBHOOK_URL_HERE',
        'ping_user_id': ''
    }
    config['CARDS'] = {
        'card1': '318,112,31,31',
        'card2': '497,115,34,29',
        'card3': '679,116,34,27',
        'card4': '407,317,34,28',
        'card5': '590,314,32,31',
        'card6': '316,522,34,29',
        'card7': '498,523,36,28',
        'card8': '678,517,36,36',
        'card9': '409,723,31,30',
        'card10': '591,726,32,28',
    }
    config['SSR_REROLL'] = {
        'password': '',
    }
    config['MACRO_REROLL'] = {
        'menu_list': '1835,1008',
        'delete_row': '963,745',
        'confirm_delete': '1098,708',
    }
    config['REGISTER_ACCOUNT'] = {
        'terms_view': '1165,480',
        'privacy_view': '1167,595',
        'i_agree': '1100,772',
        'country_change_btn': '1151,560',
        'country_ok_btn': '1104,770',
        'countrylist_ok_btn': '1083,709',
        'age_input_box': '932,563',
        'age_ok_btn': '1089,700',
        'trainer_name_box': '983,488',
        'register_btn': '962,705',
    }
    config['INGAME_MENU'] = {
        'forward_icon': '901,1023',
        'mini_gift_icon': '798,771',
        'collect_all_btn': '673,1000',
        'scout_menu_btn': '795,1027',
        'support_card_banner_btn': '846,661',
        'x10_scout_btn': '744,859',
        'confirm_scout_btn': '715,711',
        'scout_again_btn': '706,1011',
        'title_screen_btn': '1506,907',
        'banner_right_arrow_btn': '923,655',
        'support_card_banner_pos': '2',
    }
    config['SAFE_CHECK'] = {
        'conn_error_region': '855,333,217,33',
        'title_screen_btn': '860,703',
        'scout_result_region': '470,41,160,33',
        'found_ssr_card_name': '474,128,152,25',
        'found_ssr_card_epithet': '477,100,206,29',
        'ssr_misletter': 'SS, SS1, SSH, S5R',
    }
    config['REROLL_SETTINGS'] = {
        'carats': '19450',
        'general_delay': '860',
        'theme': 'simplex',
    }
    config['LINK_ACCOUNT'] = {
        'profile_btn': '1486,167',
        'copy_trainer_id_btn': '801,335',
        'data_link_btn': '1513,736',
        'data_link_confirm_btn': '683,684',
        'set_link_password_btn': '759,616',
        'password_input_box': '526,452',
        'password_confirm_input_box': '552,567',
        'privacy_policy_tick_box': '364,698',
        'next_btn': '561,1013',
        'ok_btn': '684,773',
    }
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)
    return config

def capture_and_resize_screenshot(filename_prefix):
    screenshot_dir = os.path.join(os.getcwd(), 'umapyoi_screenshots')
    if not os.path.exists(screenshot_dir):
        os.makedirs(screenshot_dir)
    screenshot = ImageGrab.grab()
    width = 640
    w, h = screenshot.size
    if w > width:
        new_h = int(h * (width / w))
        screenshot = screenshot.resize((width, new_h), Image.LANCZOS)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    screenshot_path = os.path.join(screenshot_dir, f'{filename_prefix}_{timestamp}.png')
    screenshot.save(screenshot_path, 'PNG')
    return screenshot_path

class SnippingWidget:
    def __init__(self, root, callback=None, slot_key=None):
        self.root = root
        self.callback = callback
        self.slot_key = slot_key
        self.snipping_window = None
        self.begin_x = None
        self.begin_y = None
        self.end_x = None
        self.end_y = None

    def start(self):
        self.snipping_window = Toplevel(self.root)
        self.snipping_window.attributes('-fullscreen', True)
        self.snipping_window.attributes('-alpha', 0.5)
        self.snipping_window.configure(bg="lightblue")
        self.snipping_window.lift()
        self.snipping_window.focus_force()
        self.snipping_window.bind("<Button-1>", self.on_mouse_press)
        self.snipping_window.bind("<B1-Motion>", self.on_mouse_drag)
        self.snipping_window.bind("<ButtonRelease-1>", self.on_mouse_release)
        self.canvas = Canvas(self.snipping_window, bg="lightblue", highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        self.rect = None
        self.rect_fill = None

    def on_mouse_press(self, event):
        self.begin_x = event.x
        self.begin_y = event.y
        self.canvas.delete("selection_rect")
        self.canvas.delete("selection_fill")

    def on_mouse_drag(self, event):
        self.end_x, self.end_y = event.x, event.y
        self.canvas.delete("selection_rect")
        self.canvas.delete("selection_fill")
        # Draw solid light red fill (no alpha)
        self.rect_fill = self.canvas.create_rectangle(
            self.begin_x, self.begin_y, self.end_x, self.end_y,
            fill="#ffcccc", outline="", tag="selection_fill"
        )
        # Draw red outline
        self.rect = self.canvas.create_rectangle(
            self.begin_x, self.begin_y, self.end_x, self.end_y,
            outline="red", width=2, tag="selection_rect"
        )

    def on_mouse_release(self, event):
        self.end_x = event.x
        self.end_y = event.y
        x1, y1 = min(self.begin_x, self.end_x), min(self.begin_y, self.end_y)
        x2, y2 = max(self.begin_x, self.end_x), max(self.begin_y, self.end_y)
        w, h = x2 - x1, y2 - y1
        if self.callback:
            self.callback(self.slot_key, x1, y1, w, h)
        self.snipping_window.destroy()

def load_config():
    if not os.path.exists(CONFIG_FILE):
        print('config.ini not found. Creating default config.ini...')
        create_default_config()
    config = configparser.ConfigParser()
    config.optionxform = str  # preserve case and underscores
    config.read(CONFIG_FILE)
    return config

def save_config(webhook_url, ping_user_id, card_coords=None):
    config = load_config()  # Load existing config to preserve all sections
    config['WEBHOOK'] = {
        'url': webhook_url,
        'ping_user_id': ping_user_id
    }
    if card_coords:
        config['CARDS'] = {k: f"{v[0]},{v[1]},{v[2]},{v[3]}" for k, v in card_coords.items()}
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

def load_templates():
    templates = {'SR': [], 'SSR': []}
    template_files = {'SR': [], 'SSR': []}
    for rarity in ['SR', 'SSR']:
        pattern = os.path.join('images_scans', f'{rarity.lower()}*.png')
        for path in glob.glob(pattern):
            img = cv2.imread(path, cv2.IMREAD_UNCHANGED)
            if img is not None:
                if img.shape[-1] == 4: img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                templates[rarity].append(img)
                template_files[rarity].append(os.path.basename(path))
    return templates

def preprocess_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    return gray

def detect_rarity_in_row(region_img, templates, threshold=0.7):
    results = []
    region_gray = preprocess_image(region_img)
    for rarity, tlist in templates.items():
        best_val = 0
        for template in tlist:
            template_gray = preprocess_image(template)
            if region_gray.shape[0] < template_gray.shape[0] or region_gray.shape[1] < template_gray.shape[1]:
                continue
            res = cv2.matchTemplate(region_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            if max_val > best_val:
                best_val = max_val
        if best_val >= threshold:
            results.append((rarity, best_val))
    return results

class UMAPanel:
    def __init__(self, root):
        self.root = root
        self.root.title('Uma Musume Reroll Macro (v1.0-beta)')
        self.root.geometry('530x300')
        style = Style(theme='cosmo')

        try:
            icon_path = os.path.join('misc', 'uma_icon.png')
            if os.path.exists(icon_path):
                self.root.iconphoto(False, PhotoImage(file=icon_path))
        except Exception:
            pass

        # --- Webhook and Ping User ID StringVars (fix AttributeError) ---
        config = load_config()
        webhook_url = ''
        ping_user_id = ''
        if config.has_section('WEBHOOK'):
            webhook_url = config.get('WEBHOOK', 'url', fallback='')
            ping_user_id = config.get('WEBHOOK', 'ping_user_id', fallback='')
        self.webhook_url = StringVar(value=webhook_url)
        self.ping_user_id = StringVar(value=ping_user_id)
        self.reroll_run_count = 1
        # --- Macro logging system ---
        self.logger = logging.getLogger('macro')
        self.logger.setLevel(logging.INFO)
        handler = RotatingFileHandler('macro_log.txt', maxBytes=5*1024*1024, backupCount=2, encoding='utf-8')
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        handler.setFormatter(formatter)
        if not self.logger.hasHandlers():
            self.logger.addHandler(handler)
        # --- Safe check dict (always loaded in __init__) ---
        self.safe_check = {}
        config = load_config()
        safe_keys = ['conn_error_region', 'title_screen_btn', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet', 'ssr_misletter']
        if config.has_section('SAFE_CHECK'):
            for k in safe_keys:
                v = config.get('SAFE_CHECK', k, fallback=None)
                if v is not None:
                    try:
                        if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet']:
                            x, y, w, h = map(int, v.split(','))
                            self.safe_check[k] = (x, y, w, h)
                        elif k == 'ssr_misletter':
                            self.safe_check[k] = v
                        else:
                            x, y = map(int, v.split(','))
                            self.safe_check[k] = (x, y)
                    except Exception:
                        self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
                else:
                    self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
        else:
            for k in safe_keys:
                self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
        # --- Card coordinates (x, y, w, h for 10 cards) ---
        self.card_coords = {}
        if config.has_section('CARDS'):
            for k, v in config.items('CARDS'):
                try:
                    x, y, w, h = map(int, v.split(','))
                    self.card_coords[k] = (x, y, w, h)
                except Exception:
                    self.card_coords[k] = (0, 0, 0, 0)
        else:
            for i in range(10):
                slot_key = f'card{i+1}'
                self.card_coords[slot_key] = (0, 0, 0, 0)
        # --- Card vars (must be after card_coords) ---
        self.card_vars = {}
        for i in range(10):
            slot_key = f'card{i+1}'
            x_var = IntVar(value=self.card_coords.get(slot_key, (0, 0, 0, 0))[0])
            y_var = IntVar(value=self.card_coords.get(slot_key, (0, 0, 0, 0))[1])
            w_var = IntVar(value=self.card_coords.get(slot_key, (0, 0, 0, 0))[2])
            h_var = IntVar(value=self.card_coords.get(slot_key, (0, 0, 0, 0))[3])
            self.card_vars[slot_key] = (x_var, y_var, w_var, h_var)
        # Macro reroll click coordinates
        self.reroll_clicks = {}
        config = load_config()
        if config.has_section('MACRO_REROLL'):
            for k, v in config.items('MACRO_REROLL'):
                try:
                    x, y = map(int, v.split(','))
                    self.reroll_clicks[k] = (x, y)
                except Exception:
                    self.reroll_clicks[k] = (0, 0)
        else:
            for k in ['menu_list', 'delete_row', 'confirm_delete']:
                self.reroll_clicks[k] = (0, 0)
        self.reroll_click_vars = {}
        for k in ['menu_list', 'delete_row', 'confirm_delete']:
            x_var = IntVar(value=self.reroll_clicks.get(k, (0, 0))[0])
            y_var = IntVar(value=self.reroll_clicks.get(k, (0, 0))[1])
            self.reroll_click_vars[k] = (x_var, y_var)

        # Register user account click coordinates
        self.register_clicks = {}
        config = load_config()
        country_keys = ['terms_view', 'privacy_view', 'i_agree', 'country_change_btn', 'country_ok_btn', 'countrylist_ok_btn', 'age_input_box', 'age_ok_btn', 'trainer_name_box', 'register_btn']
        if config.has_section('REGISTER_ACCOUNT'):
            for k, v in config.items('REGISTER_ACCOUNT'):
                try:
                    x, y = map(int, v.split(','))
                    self.register_clicks[k] = (x, y)
                except Exception:
                    self.register_clicks[k] = (0, 0)
        else:
            for k in country_keys:
                self.register_clicks[k] = (0, 0)
        self.register_click_vars = {}
        for k in country_keys:
            x_var = IntVar(value=self.register_clicks.get(k, (0, 0))[0])
            y_var = IntVar(value=self.register_clicks.get(k, (0, 0))[1])
            self.register_click_vars[k] = (x_var, y_var)

        # In-game menu click coordinates
        self.ingame_menu_clicks = {}
        config = load_config()
        ingame_keys = ['forward_icon', 'mini_gift_icon', 'collect_all_btn', 'scout_menu_btn', 'support_card_banner_btn', 'x10_scout_btn', 'confirm_scout_btn', 'scout_again_btn', 'title_screen_btn', 'banner_right_arrow_btn']
        if config.has_section('INGAME_MENU'):
            for k, v in config.items('INGAME_MENU'):
                try:
                    x, y = map(int, v.split(','))
                    self.ingame_menu_clicks[k] = (x, y)
                except Exception:
                    self.ingame_menu_clicks[k] = (0, 0)
        else:
            for k in ingame_keys:
                self.ingame_menu_clicks[k] = (0, 0)
        self.ingame_menu_click_vars = {}
        for k in ingame_keys:
            x_var = IntVar(value=self.ingame_menu_clicks.get(k, (0, 0))[0])
            y_var = IntVar(value=self.ingame_menu_clicks.get(k, (0, 0))[1])
            self.ingame_menu_click_vars[k] = (x_var, y_var)
        def save_ingame_menu_clicks():
            for k, (x_var, y_var) in self.ingame_menu_click_vars.items():
                self.ingame_menu_clicks[k] = (x_var.get(), y_var.get())
            config = load_config()
            if not config.has_section('INGAME_MENU'):
                config.add_section('INGAME_MENU')
            for k, (x, y) in self.ingame_menu_clicks.items():
                config.set('INGAME_MENU', k, f'{x},{y}')
            # Save support card banner position
            config.set('INGAME_MENU', 'support_card_banner_pos', str(self.support_card_banner_pos_var.get()))
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
        self.save_ingame_menu_clicks = save_ingame_menu_clicks
        def start_ingame_menu_click_snip(key):
            snip = SnippingWidget(self.root, callback=lambda k, x, y, w, h: self.set_ingame_menu_click_coord(key, x, y), slot_key=key)
            snip.start()
        self.start_ingame_menu_click_snip = start_ingame_menu_click_snip
        def set_ingame_menu_click_coord(key, x, y, w=None, h=None):
            self.ingame_menu_click_vars[key][0].set(x)
            self.ingame_menu_click_vars[key][1].set(y)
            self.save_ingame_menu_clicks()
        self.set_ingame_menu_click_coord = set_ingame_menu_click_coord

        notebook = tkttk.Notebook(root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        webhook_frame = Frame(notebook)
        reroll_frame = Frame(notebook)

        # Webhook Tab
        Label(webhook_frame, text='Discord Webhook URL:').pack(pady=(20, 5))
        url_entry = Entry(webhook_frame, textvariable=self.webhook_url, width=50, show='*')
        url_entry.pack(pady=5)
        url_entry.bind('<FocusOut>', lambda e: save_config(self.webhook_url.get(), self.ping_user_id.get(), self.card_coords))
        Label(webhook_frame, text='Ping User ID (optional):').pack(pady=(10, 5))
        ping_entry = Entry(webhook_frame, textvariable=self.ping_user_id, width=30)
        ping_entry.pack(pady=5)
        ping_entry.bind('<FocusOut>', lambda e: save_config(self.webhook_url.get(), self.ping_user_id.get(), self.card_coords))
        Button(webhook_frame, text='Send Test Notification', command=self.send_test_notification).pack(pady=10)

        # Reroll Tab - Add carats and general delay input
        self.rerolled_account_carats = IntVar(value=0)
        self.general_delay = IntVar(value=800)
        config = load_config()
        carats_val = 0
        delay_val = 800
        if config.has_section('REROLL_SETTINGS'):
            carats_val = config.getint('REROLL_SETTINGS', 'carats', fallback=0)
            delay_val = config.getint('REROLL_SETTINGS', 'general_delay', fallback=800)
        self.rerolled_account_carats.set(carats_val)
        self.general_delay.set(delay_val)
        def save_carats_and_delay(*args):
            config = load_config()
            if not config.has_section('REROLL_SETTINGS'):
                config.add_section('REROLL_SETTINGS')
            config.set('REROLL_SETTINGS', 'carats', str(self.rerolled_account_carats.get()))
            config.set('REROLL_SETTINGS', 'general_delay', str(self.general_delay.get()))
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
        carats_frame = Frame(reroll_frame)
        carats_frame.pack(pady=(10, 2))
        Label(carats_frame, text='Rerolled Account Carats (to finalize how many pulls with remains carats):').pack(side='left')
        carats_entry = Entry(carats_frame, textvariable=self.rerolled_account_carats, width=8)
        carats_entry.pack(side='left', padx=5)
        carats_entry.bind('<FocusOut>', save_carats_and_delay)
        delay_frame = Frame(reroll_frame)
        delay_frame.pack(pady=(2, 2))
        Label(delay_frame, text='General delay for all macro actions (milliseconds):').pack(side='left')
        delay_entry = Entry(delay_frame, textvariable=self.general_delay, width=8)
        delay_entry.pack(side='left', padx=5)
        delay_entry.bind('<FocusOut>', save_carats_and_delay)
        # Reroll Tab - Assign Card Regions in a separate window
        # --- 2x2 button grid layout ---
        button_grid = Frame(reroll_frame)
        button_grid.pack(pady=5)
        # Left column
        btn_assign_card = Button(button_grid, text='Assign Card Regions...', command=self.open_assign_window, width=22)
        btn_assign_card.grid(row=0, column=0, padx=5, pady=5)
        btn_macro_reroll = Button(button_grid, text='Macro Reroll Clicks', command=self.open_reroll_clicks_window, width=22)
        btn_macro_reroll.grid(row=1, column=0, padx=5, pady=5)
        # Right column
        btn_ssr_options = Button(button_grid, text='SSR Reroll Options', command=self.open_ssr_reroll_options, width=22)
        btn_ssr_options.grid(row=0, column=1, padx=5, pady=5)
        btn_safe_check = Button(button_grid, text='Macro Safe Check', command=self.open_safe_check_window, width=22)
        btn_safe_check.grid(row=1, column=1, padx=5, pady=5)

        notebook.add(webhook_frame, text='Webhook')
        notebook.add(reroll_frame, text='Reroll')

        control_frame = Frame(root)
        control_frame.pack(pady=10)
        Button(control_frame, text='Start Macro (F1)', command=self.on_start).pack(side='left', padx=10)
        Button(control_frame, text='Stop Macro (F3)', command=self.on_stop).pack(side='left', padx=10)
        # Theme dropdown with label
        theme_label = Label(control_frame, text='Theme:')
        theme_label.pack(side='left', padx=(40, 2))
        self.style = style
        config_theme = config.get('REROLL_SETTINGS', 'theme', fallback=style.theme.name)
        theme_var = StringVar(value=config_theme)
        theme_names = style.theme_names()
        theme_dropdown = tkttk.Combobox(control_frame, values=theme_names, textvariable=theme_var, width=16, state='readonly')
        theme_dropdown.pack(side='left', padx=5)
        try:
            style.theme_use(config_theme)
        except Exception:
            pass
        def on_theme_change(event=None):
            selected = theme_var.get()
            style.theme_use(selected)
            config = load_config()
            if not config.has_section('REROLL_SETTINGS'):
                config.add_section('REROLL_SETTINGS')
            config.set('REROLL_SETTINGS', 'theme', selected)
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
        theme_dropdown.bind('<<ComboboxSelected>>', on_theme_change)

        print('Loading OCR engine (EasyOCR)... This may take a few seconds.')
        self.ocr_reader = easyocr.Reader(['en'], gpu=True)
        print('OCR engine loaded.')
        self.is_macro_started = False
        self.macro_thread = None
        self.setup_global_hotkeys()

    def start_snip(self, slot_key):
        snip = SnippingWidget(self.root, callback=self.set_card_coord, slot_key=slot_key)
        snip.start()

    def set_card_coord(self, slot_key, x, y, w, h):
        self.card_vars[slot_key][0].set(x)
        self.card_vars[slot_key][1].set(y)
        self.card_vars[slot_key][2].set(w)
        self.card_vars[slot_key][3].set(h)
        self.save_card_coords()

    def save_card_coords(self):
        for k, (x_var, y_var, w_var, h_var) in self.card_vars.items():
            self.card_coords[k] = (x_var.get(), y_var.get(), w_var.get(), h_var.get())
        save_config(self.webhook_url.get(), self.ping_user_id.get(), self.card_coords)

    def open_assign_window(self):
        # Modal window for card region assignment
        assign_win = Toplevel(self.root)
        assign_win.title('Assign Card Regions')
        assign_win.geometry('420x480')
        assign_win.transient(self.root)
        Label(assign_win, text='Assign region for each card:').pack(pady=(10, 5))
        self.assign_card_vars = {}
        for i in range(10):
            slot_key = f'card{i+1}'
            x_var, y_var, w_var, h_var = self.card_vars[slot_key]
            self.assign_card_vars[slot_key] = (x_var, y_var, w_var, h_var)
            row_frame = Frame(assign_win)
            row_frame.pack(side='top', anchor='w', padx=5, pady=2)
            Label(row_frame, text=f"Card {i+1} X:").pack(side='left')
            Entry(row_frame, textvariable=x_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='Y:').pack(side='left')
            Entry(row_frame, textvariable=y_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='W:').pack(side='left')
            Entry(row_frame, textvariable=w_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='H:').pack(side='left')
            Entry(row_frame, textvariable=h_var, width=5).pack(side='left', padx=5)
            Button(row_frame, text='Assign Click', command=lambda k=slot_key: self.start_snip(k)).pack(side='left', padx=2)
        Button(assign_win, text='Save Card Regions', command=self.save_card_coords).pack(pady=10)
        Button(assign_win, text='Close', command=assign_win.destroy).pack(pady=5)

    def open_ssr_reroll_options(self):
        # SSR card data by type (specialty, card name, unique id)
        ssr_types = [
            ("Speed", [
                ("Kitasan Black", "Speed_Kitasan Black"),
                ("Kawakami Princess", "Speed_Kawakami Princess"),
                ("Special Week", "Speed_Special Week"),
                ("Twin Turbo", "Speed_Twin Turbo"),
                ("Biko Pegasus", "Speed_Biko Pegasus"),
                ("Nishino Flower", "Speed_Nishino Flower"),
                ("Sakura Bakushin O", "Speed_Sakura Bakushin O"),
                ("Gold City", "Speed_Gold City"),
                ("Tokai Teio", "Speed_Tokai Teio"),
                ("Silence Suzuka", "Speed_Silence Suzuka"),
            ]),
            ("Stamina", [
                ("Super Creek", "Stamina_Super Creek"),
                ("Satono Diamond", "Stamina_Satono Diamond"),
                ("Mejiro McQueen", "Stamina_Mejiro McQueen"),
                ("Tamamo Cross", "Stamina_Tamamo Cross"),
                ("Seiun Sky", "Stamina_Seiun Sky"),
                ("Gold Ship", "Stamina_Gold Ship"),
                ("Sakura Chiyono O", "Stamina_Sakura Chiyono O"),
                ("Rice Shower", "Stamina_Rice Shower"),
            ]),
            ("Power", [
                ("Oguri Cap", "Power_Oguri Cap"),
                ("El Condor Pasa", "Power_El Condor Pasa"),
                ("Smart Falcon", "Power_Smart Falcon"),
                ("Vodka", "Power_Vodka"),
                ("Winning Ticket", "Power_Winning Ticket"),
                ("Yaeno Muteki", "Power_Yaeno Muteki"),
            ]),
            ("Guts", [
                ("Hishi Akebono", "Guts_Hishi Akebono"),
                ("Matikane Tannhauser", "Guts_Matikane Tannhauser"),
                ("Mejiro Palmer", "Guts_Mejiro Palmer"),
                ("Haru Urara", "Guts_Haru Urara"),
                ("Winning Ticket (BNWinner!)", "Guts_Winning Ticket (BNWinner!)"),
                ("Ines Fujin", "Guts_Ines Fujin"),
                ("Grass Wonder", "Guts_Grass Wonder"),
                ("Special Week (The Brightest Star in Japan)", "Guts_Special Week (The Brightest Star in Japan)"),
            ]),
            ("Wit", [
                ("Yukino Bijin", "Wit_Yukino Bijin"),
                ("Air Shakur", "Wit_Air Shakur"),
                ("Fine Motion", "Wit_Fine Motion"),
            ]),
            ("Pal", [
                ("Tazuna Hayakawa", "Pal_Tazuna Hayakawa"),
            ]),
        ]
        # Load from config
        config = load_config()
        ssr_settings = {}
        if config.has_section('SSR_REROLL'):
            for k, v in config.items('SSR_REROLL'):
                if k == 'password':
                    continue
                try:
                    checked, min_copies = v.split(',')
                    ssr_settings[k] = (checked == '1', int(min_copies))
                except Exception:
                    ssr_settings[k] = (False, 1)
        password = config.get('SSR_REROLL', 'password', fallback='')
        add_password = bool(password)
        # UI state
        ssr_type_idx = [0]  # mutable for inner function
        assign_win = Toplevel(self.root)
        assign_win.title('SSR Reroll Options')
        assign_win.geometry('420x420')
        assign_win.transient(self.root)
        assign_win.grab_set()
        Label(assign_win, text='Ping user through discord when reroll enough of cards!').pack(pady=(10, 5))
        # Navigation bar (fixed position)
        nav_frame = Frame(assign_win)
        nav_frame.pack(pady=(0, 5))
        left_btn = Button(nav_frame, text='<', width=3)
        left_btn.pack(side='left', padx=5)
        type_label = Label(nav_frame, text='', width=12)
        type_label.pack(side='left', padx=5)
        right_btn = Button(nav_frame, text='>', width=3)
        right_btn.pack(side='left', padx=5)
        # Card frame (dynamic)
        card_frame = Frame(assign_win)
        card_frame.pack(pady=0)
        # Persistent card_vars for all specialties
        card_vars = {}
        for tname, cards in ssr_types:
            for cname, cid in cards:
                if cid not in card_vars:
                    var_checked = IntVar(value=1 if ssr_settings.get(cid, (False, 1))[0] else 0)
                    var_min = IntVar(value=ssr_settings.get(cid, (False, 1))[1])
                    card_vars[cid] = (var_checked, var_min)
        def save_settings():
            # Save SSR settings
            config = load_config()
            if not config.has_section('SSR_REROLL'):
                config.add_section('SSR_REROLL')
            for tname, cards in ssr_types:
                for cname, cid in cards:
                    checked, min_copies = card_vars[cid][0].get(), card_vars[cid][1].get()
                    config.set('SSR_REROLL', cid, f'{1 if checked else 0},{min_copies}')
            pw = pw_var.get() if pw_check_var.get() else ''
            config.set('SSR_REROLL', 'password', pw)
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
        def render_cards():
            for widget in card_frame.winfo_children():
                widget.destroy()
            tname, cards = ssr_types[ssr_type_idx[0]]
            type_label.config(text=f'< {tname} >')
            for cname, cid in cards:
                row = Frame(card_frame)
                row.pack(anchor='w', pady=2)
                var_checked, var_min = card_vars[cid]
                def on_check(event=None, cid=cid):
                    save_settings()
                def on_entry(event=None, cid=cid):
                    save_settings()
                cb = tkttk.Checkbutton(row, variable=var_checked)
                cb.pack(side='left')
                cb.bind('<ButtonRelease-1>', on_check)
                Label(row, text=f'{cname}:').pack(side='left', padx=2)
                entry = Entry(row, textvariable=var_min, width=3)
                entry.pack(side='left', padx=2)
                entry.bind('<FocusOut>', on_entry)
                Label(row, text='(Card copies minimum)').pack(side='left')
        def go_left():
            ssr_type_idx[0] = (ssr_type_idx[0] - 1) % len(ssr_types)
            render_cards()
        def go_right():
            ssr_type_idx[0] = (ssr_type_idx[0] + 1) % len(ssr_types)
            render_cards()
        left_btn.config(command=go_left)
        right_btn.config(command=go_right)
        render_cards()
        # Password option (fixed position at bottom)
        pw_frame = Frame(assign_win)
        pw_frame.pack(pady=10, side='bottom', anchor='s')
        pw_var = StringVar(value=password)
        pw_check_var = IntVar(value=1 if add_password else 0)
        def on_pw_check():
            if pw_check_var.get():
                pw_entry.config(state='normal')
            else:
                pw_entry.config(state='disabled')
        pw_check = tkttk.Checkbutton(pw_frame, text='Add password to met required SSR reroll:', variable=pw_check_var, command=on_pw_check)
        pw_check.pack(side='left')
        pw_entry = Entry(pw_frame, textvariable=pw_var, width=16)
        pw_entry.pack(side='left', padx=5)
        if not add_password:
            pw_entry.config(state='disabled')
        # Password validation label
        pw_valid_label = Label(assign_win, text='', fg='red')
        pw_valid_label.pack()
        def validate_password(pw):
            if not pw:
                return True, ''
            if not (8 <= len(pw) <= 16):
                return False, 'Password must be 8-16 characters.'
            if not any(c.isdigit() for c in pw):
                return False, 'Password must include at least one number.'
            if not any(c.islower() for c in pw):
                return False, 'Password must include at least one lowercase letter.'
            if not any(c.isupper() for c in pw):
                return False, 'Password must include at least one uppercase letter.'
            return True, ''
        def on_pw_entry(event=None):
            pw = pw_var.get()
            valid, msg = validate_password(pw)
            if not valid:
                pw_valid_label.config(text=msg)
            else:
                pw_valid_label.config(text='')
        pw_entry.bind('<FocusOut>', on_pw_entry)
        # Save button (fixed position under password)
        save_btn = Button(assign_win, text='Save', command=save_settings)
        save_btn.pack(pady=(0, 10))

    def open_reroll_clicks_window(self):
        assign_win = Toplevel(self.root)
        assign_win.title('Assign Macro Reroll Clicks')
        assign_win.geometry('420x520')
        assign_win.transient(self.root)
        # Page navigation state
        pages = [
            'Reroll',
            'Register',
            'InGameMenu',
            'LinkAccount',
        ]
        page_idx = [0]
        page_frames = {}
        def show_page(idx):
            for pf in page_frames.values():
                pf.pack_forget()
            page_frames[pages[idx]].pack(fill='both', expand=True)
            nav_label.config(text=f'< {pages[idx]} >')
        # --- Navigation bar ---
        nav_frame = Frame(assign_win)
        nav_frame.pack(pady=(10, 5))
        left_btn = Button(nav_frame, text='<', width=3)
        left_btn.pack(side='left', padx=5)
        nav_label = Label(nav_frame, text='', width=16)
        nav_label.pack(side='left', padx=5)
        right_btn = Button(nav_frame, text='>', width=3)
        right_btn.pack(side='left', padx=5)
        def go_left():
            page_idx[0] = (page_idx[0] - 1) % len(pages)
            show_page(page_idx[0])
        def go_right():
            page_idx[0] = (page_idx[0] + 1) % len(pages)
            show_page(page_idx[0])
        left_btn.config(command=go_left)
        right_btn.config(command=go_right)
        # --- Page 1: Reroll Clicks ---
        reroll_frame = Frame(assign_win)
        page_frames['Reroll'] = reroll_frame
        Label(reroll_frame, text='Assign click for each macro reroll step:').pack(pady=(10, 5))
        click_names = [
            ('Menu List Button', 'menu_list'),
            ('Delete User Data Row', 'delete_row'),
            ('Confirm Delete Data Button', 'confirm_delete'),
        ]
        for label, key in click_names:
            row_frame = Frame(reroll_frame)
            row_frame.pack(side='top', anchor='w', padx=5, pady=2)
            x_var, y_var = self.reroll_click_vars[key]
            Label(row_frame, text=f"{label} X:").pack(side='left')
            Entry(row_frame, textvariable=x_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='Y:').pack(side='left')
            Entry(row_frame, textvariable=y_var, width=5).pack(side='left', padx=5)
            Button(row_frame, text='Assign Click', command=lambda k=key: self.start_reroll_click_snip(k)).pack(side='left', padx=2)
        # --- Page 2: Register User Account ---
        register_frame = Frame(assign_win)
        page_frames['Register'] = register_frame
        Label(register_frame, text='Assign click for Register User Account:').pack(pady=(10, 5))
        reg_click_names = [
            ('Terms of Use View Button', 'terms_view'),
            ('Privacy Policy View Button', 'privacy_view'),
            ('I Agree Button', 'i_agree'),
            ('Country Change Button', 'country_change_btn'),
            ('Country OK Button', 'country_ok_btn'),
            ('Country List OK Button', 'countrylist_ok_btn'),
            ('Age Input Box', 'age_input_box'),
            ('Age Confirm OK Button', 'age_ok_btn'),
            ('Trainer Name Input Box', 'trainer_name_box'),
            ('Register Button', 'register_btn'),
        ]
        for label, key in reg_click_names:
            row_frame = Frame(register_frame)
            row_frame.pack(side='top', anchor='w', padx=5, pady=2)
            x_var, y_var = self.register_click_vars[key]
            Label(row_frame, text=f"{label} X:").pack(side='left')
            Entry(row_frame, textvariable=x_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='Y:').pack(side='left')
            Entry(row_frame, textvariable=y_var, width=5).pack(side='left', padx=5)
            Button(row_frame, text='Assign Click', command=lambda k=key: self.start_register_click_snip(k)).pack(side='left', padx=2)
        # --- Page 3: In-Game Menu Clicks ---
        ingame_frame = Frame(assign_win)
        page_frames['InGameMenu'] = ingame_frame
        Label(ingame_frame, text='Assign click for In-Game Menu:').pack(pady=(10, 5))
        # Load ingame menu clicks
        self.ingame_menu_clicks = {}
        config = load_config()
        ingame_keys = ['forward_icon', 'mini_gift_icon', 'collect_all_btn', 'scout_menu_btn', 'support_card_banner_btn', 'x10_scout_btn', 'confirm_scout_btn', 'scout_again_btn', 'title_screen_btn', 'banner_right_arrow_btn']
        if config.has_section('INGAME_MENU'):
            for k, v in config.items('INGAME_MENU'):
                try:
                    x, y = map(int, v.split(','))
                    self.ingame_menu_clicks[k] = (x, y)
                except Exception:
                    self.ingame_menu_clicks[k] = (0, 0)
        else:
            for k in ingame_keys:
                self.ingame_menu_clicks[k] = (0, 0)
        self.ingame_menu_click_vars = {}
        for k in ingame_keys:
            x_var = IntVar(value=self.ingame_menu_clicks.get(k, (0, 0))[0])
            y_var = IntVar(value=self.ingame_menu_clicks.get(k, (0, 0))[1])
            self.ingame_menu_click_vars[k] = (x_var, y_var)
        ingame_click_names = [
            ('Forward Icon', 'forward_icon'),
            ('Mini Gift Icon', 'mini_gift_icon'),
            ('Collect All Button', 'collect_all_btn'),
            ('Scout Menu Button', 'scout_menu_btn'),
            ('Support Card Banner Button', 'support_card_banner_btn'),
            ('10x Scout Button', 'x10_scout_btn'),
            ('Confirm Scout Button', 'confirm_scout_btn'),
            ('Scout Again Button', 'scout_again_btn'),
            ('Title Screen Button', 'title_screen_btn'),
            ('Banner Right Arrow Button', 'banner_right_arrow_btn'),
        ]
        for label, key in ingame_click_names:
            row_frame = Frame(ingame_frame)
            row_frame.pack(side='top', anchor='w', padx=5, pady=2)
            x_var, y_var = self.ingame_menu_click_vars[key]
            Label(row_frame, text=f"{label} X:").pack(side='left')
            Entry(row_frame, textvariable=x_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='Y:').pack(side='left')
            Entry(row_frame, textvariable=y_var, width=5).pack(side='left', padx=5)
            Button(row_frame, text='Assign Click', command=lambda k=key: self.start_ingame_menu_click_snip(k)).pack(side='left', padx=2)
        # Add support card banner position calibrate
        Label(ingame_frame, text='Support Card Banner Position (Clicks to reach):').pack(side='top', anchor='w', padx=5, pady=2)
        support_card_banner_pos_var = IntVar(value=config.getint('INGAME_MENU', 'support_card_banner_pos', fallback=0))
        Entry(ingame_frame, textvariable=support_card_banner_pos_var, width=5).pack(side='top', anchor='w', padx=5, pady=2)
        self.support_card_banner_pos_var = support_card_banner_pos_var
        # --- Page 4: Link Account Clicks ---
        link_frame = Frame(assign_win)
        page_frames['LinkAccount'] = link_frame
        Label(link_frame, text='Assign click for Link Account / Data Link:').pack(pady=(10, 5))
        link_keys = [
            ('Profile Button', 'profile_btn'),
            ('Next Button', 'next_btn'),
            ('Copy Trainer ID Button', 'copy_trainer_id_btn'),
            ('Data Link Button', 'data_link_btn'),
            ('Data Link Confirm Button', 'data_link_confirm_btn'),
            ('Set a Link Password Button', 'set_link_password_btn'),
            ('Password Input Box', 'password_input_box'),
            ('Password Confirmation Input Box', 'password_confirm_input_box'),
            ('Privacy Policy Tick Box', 'privacy_policy_tick_box'),
            ('OK Button', 'ok_btn'),
        ]
        # Load from config
        self.link_account_clicks = {}
        config = load_config()
        if config.has_section('LINK_ACCOUNT'):
            for k, v in config.items('LINK_ACCOUNT'):
                try:
                    x, y = map(int, v.split(','))
                    self.link_account_clicks[k] = (x, y)
                except Exception:
                    self.link_account_clicks[k] = (0, 0)
        else:
            for _, key in link_keys:
                self.link_account_clicks[key] = (0, 0)
        self.link_account_click_vars = {}
        for label, key in link_keys:
            x_var = IntVar(value=self.link_account_clicks.get(key, (0, 0))[0])
            y_var = IntVar(value=self.link_account_clicks.get(key, (0, 0))[1])
            self.link_account_click_vars[key] = (x_var, y_var)
            row_frame = Frame(link_frame)
            row_frame.pack(side='top', anchor='w', padx=5, pady=2)
            Label(row_frame, text=f"{label} X:").pack(side='left')
            Entry(row_frame, textvariable=x_var, width=5).pack(side='left', padx=5)
            Label(row_frame, text='Y:').pack(side='left')
            Entry(row_frame, textvariable=y_var, width=5).pack(side='left', padx=5)
            Button(row_frame, text='Assign Click', command=lambda k=key: self.start_link_account_click_snip(k)).pack(side='left', padx=2)
        # Save logic
        def save_link_account_clicks():
            for k, (x_var, y_var) in self.link_account_click_vars.items():
                self.link_account_clicks[k] = (x_var.get(), y_var.get())
            config = load_config()
            if not config.has_section('LINK_ACCOUNT'):
                config.add_section('LINK_ACCOUNT')
            for k, (x, y) in self.link_account_clicks.items():
                config.set('LINK_ACCOUNT', k, f'{x},{y}')
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
        self.save_link_account_clicks = save_link_account_clicks
        def start_link_account_click_snip(key):
            snip = SnippingWidget(self.root, callback=lambda k, x, y, w, h: self.set_link_account_click_coord(key, x, y), slot_key=key)
            snip.start()
        self.start_link_account_click_snip = start_link_account_click_snip
        def set_link_account_click_coord(key, x, y, w=None, h=None):
            self.link_account_click_vars[key][0].set(x)
            self.link_account_click_vars[key][1].set(y)
            self.save_link_account_clicks()
        self.set_link_account_click_coord = set_link_account_click_coord
        # --- Fixed bottom buttons ---
        btn_frame = Frame(assign_win)
        btn_frame.pack(side='bottom', pady=10, anchor='s')
        Button(btn_frame, text='Save Reroll Clicks', command=self.save_reroll_clicks, width=20).grid(row=0, column=0, padx=5, pady=3)
        Button(btn_frame, text='Save Register Account Clicks', command=self.save_register_clicks, width=20).grid(row=0, column=1, padx=5, pady=3)
        Button(btn_frame, text='Save In-Game Menu Clicks', command=self.save_ingame_menu_clicks, width=20).grid(row=1, column=0, padx=5, pady=3)
        Button(btn_frame, text='Save Link Account Clicks', command=self.save_link_account_clicks, width=20).grid(row=1, column=1, padx=5, pady=3)
        Button(btn_frame, text='Close', command=assign_win.destroy, width=42).grid(row=2, column=0, columnspan=2, padx=5, pady=3)
        # Show first page
        show_page(page_idx[0])

    def start_reroll_click_snip(self, key):
        snip = SnippingWidget(self.root, callback=lambda k, x, y, w, h: self.set_reroll_click_coord(key, x, y), slot_key=key)
        snip.start()

    def set_reroll_click_coord(self, key, x, y, w=None, h=None):
        self.reroll_click_vars[key][0].set(x)
        self.reroll_click_vars[key][1].set(y)
        self.save_reroll_clicks()

    def save_reroll_clicks(self):
        for k, (x_var, y_var) in self.reroll_click_vars.items():
            self.reroll_clicks[k] = (x_var.get(), y_var.get())
        config = load_config()
        if not config.has_section('MACRO_REROLL'):
            config.add_section('MACRO_REROLL')
        for k, (x, y) in self.reroll_clicks.items():
            config.set('MACRO_REROLL', k, f'{x},{y}')
        with open(CONFIG_FILE, 'w') as f:
            config.write(f)

    def start_register_click_snip(self, key):
        snip = SnippingWidget(self.root, callback=lambda k, x, y, w, h: self.set_register_click_coord(key, x, y), slot_key=key)
        snip.start()

    def set_register_click_coord(self, key, x, y, w=None, h=None):
        self.register_click_vars[key][0].set(x)
        self.register_click_vars[key][1].set(y)
        self.save_register_clicks()

    def save_register_clicks(self):
        for k, (x_var, y_var) in self.register_click_vars.items():
            self.register_clicks[k] = (x_var.get(), y_var.get())
        config = load_config()
        if not config.has_section('REGISTER_ACCOUNT'):
            config.add_section('REGISTER_ACCOUNT')
        for k, (x, y) in self.register_clicks.items():
            config.set('REGISTER_ACCOUNT', k, f'{x},{y}')
        with open(CONFIG_FILE, 'w') as f:
            config.write(f)

    def open_safe_check_window(self):
        assign_win = Toplevel(self.root)
        assign_win.title('Assign Macro Safe Check')
        assign_win.geometry('550x350')
        assign_win.transient(self.root)
        Label(assign_win, text='Assign Macro Safe Check:').pack(pady=(10, 5))
        # Load safe check clicks/regions
        self.safe_check = {}
        config = load_config()
        safe_keys = ['conn_error_region', 'title_screen_btn', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet', 'ssr_misletter']
        if config.has_section('SAFE_CHECK'):
            for k in safe_keys:
                v = config.get('SAFE_CHECK', k, fallback=None)
                if v is not None:
                    try:
                        if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet']:
                            x, y, w, h = map(int, v.split(','))
                            self.safe_check[k] = (x, y, w, h)
                        elif k == 'ssr_misletter':
                            self.safe_check[k] = v
                        else:
                            x, y = map(int, v.split(','))
                            self.safe_check[k] = (x, y)
                    except Exception:
                        self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
                else:
                    self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
        else:
            for k in safe_keys:
                self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
        # Vars for UI
        self.safe_check_vars = {}
        # Connection error region (snip)
        conn_x = IntVar(value=self.safe_check['conn_error_region'][0])
        conn_y = IntVar(value=self.safe_check['conn_error_region'][1])
        conn_w = IntVar(value=self.safe_check['conn_error_region'][2])
        conn_h = IntVar(value=self.safe_check['conn_error_region'][3])
        self.safe_check_vars['conn_error_region'] = (conn_x, conn_y, conn_w, conn_h)
        row_frame = Frame(assign_win)
        row_frame.pack(side='top', anchor='w', padx=5, pady=2)
        Label(row_frame, text='Connection Error Label Region:').pack(side='left')
        Entry(row_frame, textvariable=conn_x, width=5).pack(side='left', padx=2)
        Entry(row_frame, textvariable=conn_y, width=5).pack(side='left', padx=2)
        Entry(row_frame, textvariable=conn_w, width=5).pack(side='left', padx=2)
        Entry(row_frame, textvariable=conn_h, width=5).pack(side='left', padx=2)
        Button(row_frame, text='Assign Snip', command=lambda: self.start_safe_check_snip('conn_error_region')).pack(side='left', padx=2)
        # Title Screen button
        ts_x = IntVar(value=self.safe_check['title_screen_btn'][0])
        ts_y = IntVar(value=self.safe_check['title_screen_btn'][1])
        self.safe_check_vars['title_screen_btn'] = (ts_x, ts_y)
        row_frame2 = Frame(assign_win)
        row_frame2.pack(side='top', anchor='w', padx=5, pady=2)
        Label(row_frame2, text='Title Screen Button X:').pack(side='left')
        Entry(row_frame2, textvariable=ts_x, width=5).pack(side='left', padx=2)
        Label(row_frame2, text='Y:').pack(side='left')
        Entry(row_frame2, textvariable=ts_y, width=5).pack(side='left', padx=2)
        Button(row_frame2, text='Assign Click', command=lambda: self.start_safe_check_click('title_screen_btn')).pack(side='left', padx=2)
        # Scout Result region (snip)
        scout_x = IntVar(value=self.safe_check['scout_result_region'][0])
        scout_y = IntVar(value=self.safe_check['scout_result_region'][1])
        scout_w = IntVar(value=self.safe_check['scout_result_region'][2])
        scout_h = IntVar(value=self.safe_check['scout_result_region'][3])
        self.safe_check_vars['scout_result_region'] = (scout_x, scout_y, scout_w, scout_h)
        row_frame3 = Frame(assign_win)
        row_frame3.pack(side='top', anchor='w', padx=5, pady=2)
        Label(row_frame3, text='Scout Result Region:').pack(side='left')
        Entry(row_frame3, textvariable=scout_x, width=5).pack(side='left', padx=2)
        Entry(row_frame3, textvariable=scout_y, width=5).pack(side='left', padx=2)
        Entry(row_frame3, textvariable=scout_w, width=5).pack(side='left', padx=2)
        Entry(row_frame3, textvariable=scout_h, width=5).pack(side='left', padx=2)
        Button(row_frame3, text='Assign Snip', command=lambda: self.start_safe_check_snip('scout_result_region')).pack(side='left', padx=2)
        # Found SSR card name region (snip)
        ssrname_x = IntVar(value=self.safe_check['found_ssr_card_name'][0])
        ssrname_y = IntVar(value=self.safe_check['found_ssr_card_name'][1])
        ssrname_w = IntVar(value=self.safe_check['found_ssr_card_name'][2])
        ssrname_h = IntVar(value=self.safe_check['found_ssr_card_name'][3])
        self.safe_check_vars['found_ssr_card_name'] = (ssrname_x, ssrname_y, ssrname_w, ssrname_h)
        row_frame4 = Frame(assign_win)
        row_frame4.pack(side='top', anchor='w', padx=5, pady=2)
        Label(row_frame4, text='SSR Card Name Region:').pack(side='left')
        Entry(row_frame4, textvariable=ssrname_x, width=5).pack(side='left', padx=2)
        Entry(row_frame4, textvariable=ssrname_y, width=5).pack(side='left', padx=2)
        Entry(row_frame4, textvariable=ssrname_w, width=5).pack(side='left', padx=2)
        Entry(row_frame4, textvariable=ssrname_h, width=5).pack(side='left', padx=2)
        Button(row_frame4, text='Assign Snip', command=lambda: self.start_safe_check_snip('found_ssr_card_name')).pack(side='left', padx=2)
        # Found SSR card epithet region (snip)
        ssrepi_x = IntVar(value=self.safe_check['found_ssr_card_epithet'][0])
        ssrepi_y = IntVar(value=self.safe_check['found_ssr_card_epithet'][1])
        ssrepi_w = IntVar(value=self.safe_check['found_ssr_card_epithet'][2])
        ssrepi_h = IntVar(value=self.safe_check['found_ssr_card_epithet'][3])
        self.safe_check_vars['found_ssr_card_epithet'] = (ssrepi_x, ssrepi_y, ssrepi_w, ssrepi_h)
        row_frame5 = Frame(assign_win)
        row_frame5.pack(side='top', anchor='w', padx=5, pady=2)
        Label(row_frame5, text='SSR Card Epithet Region:').pack(side='left')
        Entry(row_frame5, textvariable=ssrepi_x, width=5).pack(side='left', padx=2)
        Entry(row_frame5, textvariable=ssrepi_y, width=5).pack(side='left', padx=2)
        Entry(row_frame5, textvariable=ssrepi_w, width=5).pack(side='left', padx=2)
        Entry(row_frame5, textvariable=ssrepi_h, width=5).pack(side='left', padx=2)
        Button(row_frame5, text='Assign Snip', command=lambda: self.start_safe_check_snip('found_ssr_card_epithet')).pack(side='left', padx=2)
        # SSR misletter input
        ssr_misletter_var = StringVar(value=self.safe_check.get('ssr_misletter', ''))
        self.safe_check_vars['ssr_misletter'] = ssr_misletter_var
        row_frame6 = Frame(assign_win)
        row_frame6.pack(side='top', anchor='w', padx=5, pady=2)
        Label(row_frame6, text='SSR Card Type Misletter (optional):').pack(side='left')
        Entry(row_frame6, textvariable=ssr_misletter_var, width=24).pack(side='left', padx=2)
        Label(row_frame6, text='(comma separated)').pack(side='left')
        Button(assign_win, text='Save Safe Check', command=self.save_safe_check).pack(pady=10)
        Button(assign_win, text='Close', command=assign_win.destroy).pack(pady=5)

    def start_safe_check_snip(self, key):
        snip = SnippingWidget(self.root, callback=lambda k, x, y, w, h: self.set_safe_check_region(key, x, y, w, h), slot_key=key)
        snip.start()

    def set_safe_check_region(self, key, x, y, w, h):
        self.safe_check_vars[key][0].set(x)
        self.safe_check_vars[key][1].set(y)
        self.safe_check_vars[key][2].set(w)
        self.safe_check_vars[key][3].set(h)
        self.save_safe_check()

    def start_safe_check_click(self, key):
        snip = SnippingWidget(self.root, callback=lambda k, x, y, w, h: self.set_safe_check_click_coord(key, x, y), slot_key=key)
        snip.start()

    def set_safe_check_click_coord(self, key, x, y, w=None, h=None):
        self.safe_check_vars[key][0].set(x)
        self.safe_check_vars[key][1].set(y)
        self.save_safe_check()

    def save_safe_check(self):
        config = load_config()
        if not config.has_section('SAFE_CHECK'):
            config.add_section('SAFE_CHECK')
        # Save region
        x, y, w, h = [v.get() for v in self.safe_check_vars['conn_error_region']]
        config.set('SAFE_CHECK', 'conn_error_region', f'{x},{y},{w},{h}')
        # Save title screen btn
        x, y = [v.get() for v in self.safe_check_vars['title_screen_btn']]
        config.set('SAFE_CHECK', 'title_screen_btn', f'{x},{y}')
        # Save scout result region
        x, y, w, h = [v.get() for v in self.safe_check_vars['scout_result_region']]
        config.set('SAFE_CHECK', 'scout_result_region', f'{x},{y},{w},{h}')
        # Save found ssr card name region
        x, y, w, h = [v.get() for v in self.safe_check_vars['found_ssr_card_name']]
        config.set('SAFE_CHECK', 'found_ssr_card_name', f'{x},{y},{w},{h}')
        # Save found ssr card epithet region
        x, y, w, h = [v.get() for v in self.safe_check_vars['found_ssr_card_epithet']]
        config.set('SAFE_CHECK', 'found_ssr_card_epithet', f'{x},{y},{w},{h}')
        # Save SSR misletter
        ssr_misletter = self.safe_check_vars['ssr_misletter'].get()
        config.set('SAFE_CHECK', 'ssr_misletter', ssr_misletter)
        with open(CONFIG_FILE, 'w') as f:
            config.write(f)

    def run_rarity_detection(self, return_results=False, run_idx=1):
        templates = {'R': [], 'SR': [], 'SSR': []}
        template_files = {'R': [], 'SR': [], 'SSR': []}
        for rarity in ['R', 'SR', 'SSR']:
            pattern = os.path.join('images_scans', f'{rarity.lower()}*.png')
            for path in glob.glob(pattern):
                img = cv2.imread(path, cv2.IMREAD_UNCHANGED)
                if img is not None:
                    if img.shape[-1] == 4:
                        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                    templates[rarity].append(img)
                    template_files[rarity].append(os.path.basename(path))
        if not any(templates.values()):
            print('No rarity templates found!')
            return [] if return_results else None
        icon_w, icon_h = 32, 32
        reader = self.ocr_reader
        config = load_config()
        ssr_misletters = ['ss', 'ss1']
        if config.has_section('SAFE_CHECK'):
            misletter_str = config.get('SAFE_CHECK', 'ssr_misletter', fallback='')
            if misletter_str:
                ssr_misletters += [m.strip().lower() for m in misletter_str.split(',') if m.strip()]
        num_runs = 5
        card_ocr_results = [[] for _ in range(10)]
        card_template_results = [[] for _ in range(10)]
        card_ocr_raw = [[] for _ in range(10)]
        card_ocr_valid = [[] for _ in range(10)]
        for run in range(num_runs):
            if not self.is_macro_started: return [ [{'rarity': 'Unknown'} for _ in range(10)] ] if return_results else None
            if run > 0: time.sleep(0.55)
            screenshot = ImageGrab.grab()
            screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            regions = []
            valid_indices = []
            for i in range(10):
                slot_key = f'card{i+1}'
                coords = self.card_coords.get(slot_key)
                if not coords:
                    card_ocr_results[i].append('Unknown')
                    card_template_results[i].append('Unknown')
                    card_ocr_raw[i].append('')
                    card_ocr_valid[i].append('')
                    regions.append(None)
                    continue
                x, y, w, h = coords
                region = screenshot[y:y+h, x:x+w]
                regions.append(region)
                valid_indices.append(i)
                region_gray = preprocess_image(region)
                best_rarity = 'Unknown'
                best_score = 0
                for rarity, tlist in templates.items():
                    for idx, template in enumerate(tlist):
                        template_gray = preprocess_image(template)
                        if region_gray.shape[0] < template_gray.shape[0] or region_gray.shape[1] < template_gray.shape[1]:
                            continue
                        res = cv2.matchTemplate(region_gray, template_gray, cv2.TM_CCOEFF_NORMED)
                        _, max_val, _, _ = cv2.minMaxLoc(res)
                        if max_val > best_score:
                            best_score = max_val
                            best_rarity = rarity
                card_template_results[i].append(best_rarity)
            # Batch OCR for all valid regions
            batch_regions = [r for r in regions if r is not None]
            batch_results = reader.readtext(batch_regions) if batch_regions else []
            # EasyOCR returns a list of lists (one per image)
            batch_idx = 0
            for i in range(10):
                if regions[i] is None:
                    continue
                ocr_valid = None
                result = batch_results[batch_idx] if batch_idx < len(batch_results) else []
                batch_idx += 1
                text = ' '.join([r[1] for r in result]).upper().strip()
                card_ocr_raw[i].append(text)
                if any(label in text for label in ['SSR']):
                    ocr_valid = 'SSR'
                elif ssr_misletters and any(mis and mis in text.lower() for mis in ssr_misletters):
                    ocr_valid = 'SSR'
                elif 'SR' in text:
                    ocr_valid = 'SR'
                elif 'R' in text:
                    ocr_valid = 'R'
                card_ocr_results[i].append(ocr_valid if ocr_valid else 'Unknown')
                card_ocr_valid[i].append(ocr_valid if ocr_valid else '')
        final_results = []
        for i in range(10):
            # If all ocr_raw and ocr_valid are blank for this card, assign R immediately
            if all(x == '' for x in card_ocr_raw[i]) and all(x == '' for x in card_ocr_valid[i]):
                final = 'R'
            else:
                ocr_counts = {'SSR': 0, 'SR': 0, 'R': 0}
                for val in card_ocr_results[i]:
                    if val in ocr_counts:
                        ocr_counts[val] += 1
                # --- Improved SSR/SR voting logic ---
                if ocr_counts['SSR'] >= 1 and ocr_counts['SSR'] >= ocr_counts['SR']:
                    final = 'SSR'
                elif ocr_counts['SR'] > ocr_counts['SSR'] and ocr_counts['SR'] >= 2:
                    final = 'SR'
                elif ocr_counts['R'] >= 2:
                    final = 'R'
                else:
                    filtered = [r for r in card_template_results[i] if r != 'Unknown']
                    if filtered:
                        most_common = max(set(filtered), key=filtered.count)
                        final = most_common
                    else:
                        ocr_non_unknown = [val for val in card_ocr_results[i] if val != 'Unknown']
                        if ocr_non_unknown:
                            final = max(set(ocr_non_unknown), key=ocr_non_unknown.count)
                        else:
                            final = 'Unknown'
            print(f"Card {i+1}: template={card_template_results[i]}, ocr_raw={card_ocr_raw[i]}, ocr_valid={card_ocr_valid[i]}, final={final}")
            final_results.append({'rarity': final})
        if return_results:
            return [final_results]

    def send_test_notification(self):
        url = self.webhook_url.get()
        ping = self.ping_user_id.get()
        send_discord_notification('Test notification from Uma Musume Macro!', url, ping)

    def Global_MouseClick(self, x, y, click=1):
        time.sleep(0.1)
        autoit.mouse_click("left", x, y, click, speed=3)

    @macro_active_only
    def macro_main_loop(self):
        try:
            while self.is_macro_started:
                win = None
                for w in gw.getAllWindows():
                    if w.title == 'Umamusume':
                        win = w
                        break
                if win:
                    try:
                        win.restore()
                    except Exception:
                        pass
                    win.activate()
                    print('Activated Umamusume window.')
                    time.sleep(3)
                    if not self.is_macro_started:
                        print('Macro stopped before reroll delete.')
                        break
                    self.macro_reroll_delete_loop()
                    for i in range(5):
                        if not self.is_macro_started:
                            print(f'Macro stopped during pre-register click loop (iteration {i+1}).')
                            break
                        time.sleep(1.5)
                        self.Global_MouseClick(600,600)
                    if not self.is_macro_started: break
                    time.sleep(1.5)
                    if not self.is_macro_started: break

                    self.register_user_account_loop()
                    screenshot_path = capture_and_resize_screenshot('trainer_registered')
                    send_discord_notification(
                        message=f'Trainer registered successfully (Run {self.reroll_run_count})',
                        webhook_url=self.webhook_url.get(),
                        ping_user_id=self.ping_user_id.get(),
                        status='register',
                        description=None,
                        ping_on_embed=False,
                        ssr_details=None,
                        is_summary=False,
                        screenshot_path=screenshot_path
                    )
                    if not self.is_macro_started: break
                    time.sleep(1.8)
                   
                    self.receive_carats_before_reroll()
                    
                    if not self.is_macro_started: break
                    time.sleep(1.2)
                    # Open up scout menu webhook
                    screenshot_path = capture_and_resize_screenshot('opening_scout_menu')
                    send_discord_notification(
                        message=f'Opening scout menu (Run {self.reroll_run_count})',
                        webhook_url=self.webhook_url.get(),
                        ping_user_id=self.ping_user_id.get(),
                        status='scoutmenu',
                        description=None,
                        ping_on_embed=False,
                        ssr_details=None,
                        is_summary=False,
                        screenshot_path=screenshot_path
                    )
                    self.scout_reroll_loop()
                    time.sleep(15)
                    send_discord_notification(
                        message=f'Starting Run {self.reroll_run_count} (Going back to title)',
                        webhook_url=self.webhook_url.get(),
                        ping_user_id=self.ping_user_id.get(),
                        status='runstart',
                        description=None,
                        ping_on_embed=False,
                        ssr_details=None,
                        is_summary=False,
                        screenshot_path=capture_and_resize_screenshot('starting_run')
                    )
                else:
                    print('Umamusume window not found!')
                    break
        except Exception as e:
            self.logger.error(f'Error automating Umamusume: {e}')
        finally: pass

    def setup_global_hotkeys(self):
        keyboard.add_hotkey('F1', lambda: self.on_start())
        keyboard.add_hotkey('F3', lambda: self.on_stop())
        print('Hotkeys registered: F1=start, F3=stop')

    def on_start(self):
        if self.is_macro_started:
            print('Macro already started!')
            return
        print('Starting Macro...')
        self.is_macro_started = True
        send_discord_notification('', self.webhook_url.get(), self.ping_user_id.get(), status='start')
        self.start_connection_error_failsafe()  # Start failsafe thread
        self.macro_thread = threading.Thread(target=self.macro_main_loop, daemon=True)
        self.macro_thread.start()

    def on_stop(self):
        if not self.is_macro_started:
            print('Macro is not running!')
            return
        print('Stopping Macro...')
        self.is_macro_started = False
        send_discord_notification('', self.webhook_url.get(), self.ping_user_id.get(), status='stop')

    @macro_active_only
    def macro_reroll_delete_loop(self):
        config = load_config()
        section = ConfigSection.MACRO_REROLL.value
        try:
            menu_list = config.get(section, 'menu_list', fallback=None)
            delete_row = config.get(section, 'delete_row', fallback=None)
            confirm_delete = config.get(section, 'confirm_delete', fallback=None)
            if not (menu_list and delete_row and confirm_delete):
                print('One or more Macro Reroll Clicks are not set!')
                return
            x1, y1 = map(int, menu_list.split(','))
            x2, y2 = map(int, delete_row.split(','))
            x3, y3 = map(int, confirm_delete.split(','))
            print(f'Clicking Menu List at ({x1}, {y1})')
            self.Global_MouseClick(x1, y1)
            self.macro_sleep(1.5)
            print(f'Clicking Delete User Data Row at ({x2}, {y2})')
            self.Global_MouseClick(x2, y2)
            self.macro_sleep(1.5)
            print(f'Clicking Confirm Delete Data Button at ({x3}, {y3}) 5 times')
            for i in range(5):
                if not self.is_macro_started:
                    print(f'Macro stopped during Confirm Delete Data Button loop (iteration {i+1}).')
                    return
                self.Global_MouseClick(x3, y3)
                self.macro_sleep(0.85)
        except Exception as e:
            print(f'Error in macro_reroll_delete_loop: {e}')
    
    def register_user_account_loop(self):
        config = load_config()
        section = 'REGISTER_ACCOUNT'
        try:
            terms_view = config.get(section, 'terms_view', fallback=None)
            privacy_view = config.get(section, 'privacy_view', fallback=None)
            i_agree = config.get(section, 'i_agree', fallback=None)
            country_change_btn = config.get(section, 'country_change_btn', fallback=None)
            country_ok_btn = config.get(section, 'country_ok_btn', fallback=None)
            countrylist_ok_btn = config.get(section, 'countrylist_ok_btn', fallback=None)
            age_input_box = config.get(section, 'age_input_box', fallback=None)
            age_ok_btn = config.get(section, 'age_ok_btn', fallback=None)
            trainer_name_box = config.get(section, 'trainer_name_box', fallback=None)
            register_btn = config.get(section, 'register_btn', fallback=None)
            if not (terms_view and privacy_view and i_agree):
                print('One or more Register Account Clicks are not set!')
                return
            x1, y1 = map(int, terms_view.split(','))
            x2, y2 = map(int, privacy_view.split(','))
            x3, y3 = map(int, i_agree.split(','))
            time.sleep(2.5)
            print(f'Clicking Terms of Use View at ({x1}, {y1})')
            self.Global_MouseClick(x1, y1)
            time.sleep(5.5)
            if not self.is_macro_started:
                print('Macro stopped after Terms of Use View click.')
                return
            print('Sending Ctrl+W to close browser tab')
            autoit.send('^w')
            time.sleep(2.2)
            if not self.is_macro_started:
                print('Macro stopped after first Ctrl+W.')
                return
            # Focus game window again
            win = None
            for w in gw.getAllWindows():
                if w.title == 'Umamusume':
                    win = w
                    break
            if win:
                win.activate()
                time.sleep(3.5)
            if not self.is_macro_started:
                print('Macro stopped after refocusing game (after Terms).')
                return
            print(f'Clicking Privacy Policy View at ({x2}, {y2})')
            self.Global_MouseClick(x2, y2)
            time.sleep(5.2)
            if not self.is_macro_started:
                print('Macro stopped after Privacy Policy View click.')
                return
            print('Sending Ctrl+W to close browser tab')
            autoit.send('^w')
            time.sleep(3.7)
            if not self.is_macro_started:
                print('Macro stopped after second Ctrl+W.')
                return
            if win:
                win.activate()
                time.sleep(2.5)
            if not self.is_macro_started:
                print('Macro stopped after refocusing game (after Privacy).')
                return
            print(f'Clicking I Agree at ({x3}, {y3})')
            self.Global_MouseClick(x3, y3)
            time.sleep(2.5)
            if not self.is_macro_started:
                print('Macro stopped after I Agree click.')
                return
            # After country selection, handle age input
            if country_change_btn and country_ok_btn and countrylist_ok_btn:
                x1, y1 = map(int, country_change_btn.split(','))
                x2, y2 = map(int, country_ok_btn.split(','))
                x3, y3 = map(int, countrylist_ok_btn.split(','))
                print(f'Clicking Country Change Button at ({x1}, {y1})')
                self.Global_MouseClick(x1, y1)
                time.sleep(1.8)
                print(f'Clicking Country List OK Button at ({x3}, {y3})')
                self.Global_MouseClick(x2, y2)
                time.sleep(1.8)
                print(f'Clicking Country OK Button at ({x2}, {y2})')
                self.Global_MouseClick(x3, y3)
                time.sleep(1.85)
            # Age input
            if age_input_box and age_ok_btn:
                x4, y4 = map(int, age_input_box.split(','))
                x5, y5 = map(int, age_ok_btn.split(','))
                print(f'Clicking Age Input Box at ({x4}, {y4})')
                self.Global_MouseClick(x4, y4)
                time.sleep(0.8)
                print('Typing 199001 for age confirmation')
                autoit.send('199001')
                time.sleep(0.8)
                print(f'Clicking Age Confirm OK Button at ({x5}, {y5})')
                self.Global_MouseClick(x5, y5)
                time.sleep(4.5)
                # Skip Tutorial (reuse age_ok_btn coordinates)
                print(f'Clicking Skip Tutorial Button at ({x5}, {y5})')
                for i in range(7):
                    if not self.is_macro_started: return
                    self.Global_MouseClick(x5, y5)
                    time.sleep(1.65)
                
                time.sleep(1.5)
            # Trainer Registration
            if trainer_name_box and register_btn:
                x6, y6 = map(int, trainer_name_box.split(','))
                x7, y7 = map(int, register_btn.split(','))
                print(f'Clicking Trainer Name Input Box at ({x6}, {y6})')
                self.Global_MouseClick(x6, y6)
                time.sleep(0.8)
                print("Typing 'RerolledAccount' as trainer name")
                autoit.send('RerolledAccount')
                time.sleep(1.2)
                print(f'Clicking Register Button at ({x7}, {y7})')
                self.Global_MouseClick(x7, y7)
                time.sleep(1.5)
                # Confirm registration OK (reuse age_ok_btn)
                if age_ok_btn:
                    x5, y5 = map(int, age_ok_btn.split(','))
                    print(f'Clicking Registration OK Button at ({x5}, {y5})')
                    self.Global_MouseClick(x5, y5)
                    time.sleep(1.5)
        except Exception as e:
            print(f'Error in register_user_account_loop: {e}') 

    def receive_carats_before_reroll(self):
        print('Starting receive_carats_before_reroll...')
        config = load_config()
        section = 'INGAME_MENU'
        try:
            forward_icon = config.get(section, 'forward_icon', fallback=None)
            mini_gift_icon = config.get(section, 'mini_gift_icon', fallback=None)
            collect_all_btn = config.get(section, 'collect_all_btn', fallback=None)
            if not forward_icon or not mini_gift_icon or not collect_all_btn:
                print('One or more in-game menu clicks are not set!')
                return
            x_fwd, y_fwd = map(int, forward_icon.split(','))
            x_gift, y_gift = map(int, mini_gift_icon.split(','))
            x_collect, y_collect = map(int, collect_all_btn.split(','))
            
            screenshot_path = capture_and_resize_screenshot('collecting_carats')
            send_discord_notification(
                message=f'Collecting carats (Run {self.reroll_run_count})',
                webhook_url=self.webhook_url.get(),
                ping_user_id=self.ping_user_id.get(),
                status='carats',
                description=None,
                ping_on_embed=False,
                ssr_details=None,
                is_summary=False,
                screenshot_path=screenshot_path
            )
            
            for i in range(30):
                if not self.is_macro_started:
                    print(f'Macro stopped during forward_icon intro loop (iteration {i+1}).')
                    return
                print(f'Clicking Forward Icon ({i+1}/20) at ({x_fwd}, {y_fwd})')
                self.Global_MouseClick(x_fwd, y_fwd)
                time.sleep(0.85)
            if not self.is_macro_started: return
            time.sleep(1.85)
            print(f'Clicking Mini Gift Icon at ({x_gift}, {y_gift})')
            self.Global_MouseClick(x_gift, y_gift)
            time.sleep(3)
            if not self.is_macro_started: return
            print(f'Clicking Collect All Button at ({x_collect}, {y_collect})')
         
            for i in range(4):
                if not self.is_macro_started: return
                self.Global_MouseClick(x_collect, y_collect)
                time.sleep(1.5)
                
            time.sleep(1.5)
            
            if not self.is_macro_started: return
            for i in range(6):
                if not self.is_macro_started:
                    print(f'Macro stopped during forward_icon after collect loop (iteration {i+1}).')
                    return
                print(f'Clicking Forward Icon ({i+1}/6) at ({x_fwd}, {y_fwd})')
                self.Global_MouseClick(x_fwd, y_fwd)
                time.sleep(0.85)
            print('Finished receive_carats_before_reroll.')
        except Exception as e:
            print(f'Error in receive_carats_before_reroll: {e}')
    
    @macro_active_only
    def scout_reroll_loop(self):
        self.logger.info('Starting scout_reroll_loop...')
        if not hasattr(self, 'reroll_run_count'):
            self.reroll_run_count = 1
        if not self.is_macro_started: return
        config = load_config()
        section = 'INGAME_MENU'
        safe_keys = ['conn_error_region', 'title_screen_btn', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet', 'ssr_misletter']
        self.safe_check = {}
        if config.has_section('SAFE_CHECK'):
            for k in safe_keys:
                v = config.get('SAFE_CHECK', k, fallback=None)
                if v is not None:
                    try:
                        if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet']:
                            x, y, w, h = map(int, v.split(','))
                            self.safe_check[k] = (x, y, w, h)
                        elif k == 'ssr_misletter':
                            self.safe_check[k] = v
                        else:
                            x, y = map(int, v.split(','))
                            self.safe_check[k] = (x, y)
                    except Exception:
                        self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
                else:
                    self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
        else:
            for k in safe_keys:
                self.safe_check[k] = (0, 0, 0, 0) if k in ['conn_error_region', 'scout_result_region', 'found_ssr_card_name', 'found_ssr_card_epithet'] else (0, 0)
        try:
            if not self.is_macro_started: return
            carats = 0
            delay_ms = 800
            if config.has_section('REROLL_SETTINGS'):
                try:
                    carats = config.getint('REROLL_SETTINGS', 'carats', fallback=0)
                except Exception:
                    carats = 0
                try:
                    delay_ms = config.getint('REROLL_SETTINGS', 'general_delay', fallback=800)
                except Exception:
                    delay_ms = 800
            pulls = max(1, carats // 1500)
            delay = max(0.01, delay_ms / 1000.0)
            print(f'Carats: {carats}, will perform {pulls} x10 pulls. General delay: {delay}s')
            
            scout_menu_btn = config.get(section, 'scout_menu_btn', fallback=None)
            support_card_banner_btn = config.get(section, 'support_card_banner_btn', fallback=None)
            x10_scout_btn = config.get(section, 'x10_scout_btn', fallback=None)
            confirm_scout_btn = config.get(section, 'confirm_scout_btn', fallback=None)
            forward_icon = config.get(section, 'forward_icon', fallback=None)
            scout_again_btn = config.get(section, 'scout_again_btn', fallback=None)
            if not (scout_menu_btn and support_card_banner_btn and x10_scout_btn and confirm_scout_btn and forward_icon and scout_again_btn):
                print('One or more scout reroll clicks are not set!')
                return
            x_scout, y_scout = map(int, scout_menu_btn.split(','))
            x_x10, y_x10 = map(int, x10_scout_btn.split(','))
            x_confirm, y_confirm = map(int, confirm_scout_btn.split(','))
            x_fwd, y_fwd = map(int, forward_icon.split(','))
            x_again, y_again = map(int, scout_again_btn.split(','))
            ssr_reroll_names = []
            ssr_reroll_types = []
            ssr_reroll_keys = []
            if config.has_section('SSR_REROLL'):
                for k in config['SSR_REROLL'].keys():
                    if k == 'password':
                        continue
                    if '_' in k:
                        ssr_type, ssr_name = k.split('_', 1)
                    else:
                        ssr_type, ssr_name = '', k
                    ssr_reroll_keys.append(k.lower())
                    ssr_reroll_names.append(ssr_name)
                    ssr_reroll_types.append(ssr_type)
            total_ssr_counter = {}
            for pull_num in range(pulls):
                if not self.is_macro_started: return
                print(f'\n===== Rerolling SSR\'s Card (Run {self.reroll_run_count}) (#{pull_num+1} Scout) =====')
                if pull_num == 0:
                    time.sleep(delay)
                    if not self.is_macro_started: return
                    print(f'Clicking Scout Menu Button at ({x_scout}, {y_scout})')
                    self.Global_MouseClick(x_scout, y_scout)
                    time.sleep(1.2 + delay)
                    self.Global_MouseClick(x_scout, y_scout)
                    time.sleep(5.2 + delay)
                    if not self.is_macro_started: return
                    support_card_banner_pos = config.getint('INGAME_MENU', 'support_card_banner_pos', fallback=0)
                    banner_right_arrow_btn = config.get('INGAME_MENU', 'banner_right_arrow_btn', fallback=None)
                    if banner_right_arrow_btn:
                        x_arrow, y_arrow = map(int, banner_right_arrow_btn.split(','))
                        for _ in range(support_card_banner_pos):
                            if not self.is_macro_started: return
                            self.Global_MouseClick(x_arrow, y_arrow)
                            time.sleep(1.25 + delay)
                    time.sleep(3.2 + delay)
                    if not self.is_macro_started: return
                else:
                    if not self.is_macro_started: return
                    print('Clicking Scout Again Button...')
                    self.Global_MouseClick(x_again, y_again)
                    time.sleep(1.5 + delay)
                    print('Clicking Confirm Scout Button...')
                    self.Global_MouseClick(x_confirm, y_confirm)
                    time.sleep(1.5 + delay)
                if pull_num == 0:
                    print(f'Clicking 10x Scout Button at ({x_x10}, {y_x10})')
                if pull_num == 0:
                    self.Global_MouseClick(x_x10, y_x10)
                    time.sleep(1.2 + delay)
                if pull_num == 0:
                    print(f'Clicking Confirm Scout Button at ({x_confirm}, {y_confirm})')
                    self.Global_MouseClick(x_confirm, y_confirm)
                    time.sleep(1.2 + delay)
                for i in range(12):
                    if not self.is_macro_started: return
                    self.Global_MouseClick(x_fwd, y_fwd)
                    time.sleep(0.3 + delay)
                scout_region = config.get('SAFE_CHECK', 'scout_result_region', fallback=None)
                if not scout_region:
                    print('Scout result region not set!')
                    return
                x, y, w, h = map(int, scout_region.split(','))
                found = False
                for attempt in range(4):
                    if not self.is_macro_started: return
                    time.sleep(0.5 + delay)
                    screenshot = ImageGrab.grab()
                    region = screenshot.crop((x, y, x+w, y+h))
                    img = cv2.cvtColor(np.array(region), cv2.COLOR_RGB2BGR)
                    result = self.ocr_reader.readtext(img)
                    text = ' '.join([r[1] for r in result]).strip().lower()
                    print(f'Scout Results OCR attempt {attempt+1}: "{text}"')
                    if ('scout results' in text) or ('scout' in text) or ('results' in text):
                        found = True
                        break
                if not found:
                    print('Scout Results not detected! Macro will not proceed.')
                    return
                print('Scout Results detected. Proceeding to card rarity detection...')
                time.sleep(0.5 + delay)
                if not self.is_macro_started: return
                all_card_details = self.run_rarity_detection(return_results=True, run_idx=pull_num+1)
                if not all_card_details:
                    print('No rarity results found!')
                    return
                first_run = all_card_details[0] if isinstance(all_card_details[0], list) else all_card_details
                ssr_count = sum(1 for c in first_run if c['rarity'] == 'SSR')
                sr_count = sum(1 for c in first_run if c['rarity'] == 'SR')
                r_count = sum(1 for c in first_run if c['rarity'] == 'R')
                card_coords = config['CARDS'] if config.has_section('CARDS') else {}
                ssr_indices = [i for i, c in enumerate(first_run) if c['rarity'] == 'SSR']
                ssr_this_pull = {}
                if not ssr_indices:
                    print('No SSR cards found in scout results.')
                else:
                    print(f'SSR found at card indices: {[i+1 for i in ssr_indices]}')
                    ssr_name_region = config.get('SAFE_CHECK', 'found_ssr_card_name', fallback=None)
                    if not ssr_name_region:
                        print('SSR card name region not set!')
                        return
                    x_name, y_name, w_name, h_name = map(int, ssr_name_region.split(','))
                    for idx in ssr_indices:
                        if not self.is_macro_started: return
                        slot_key = f'card{idx+1}'
                        if slot_key not in card_coords:
                            print(f'No coordinates for {slot_key} in config!')
                            continue
                        x_card, y_card, *_ = map(int, card_coords[slot_key].split(','))
                        print(f'Clicking SSR card {slot_key} at ({x_card}, {y_card})')
                        self.Global_MouseClick(x_card, y_card)
                        time.sleep(0.8 + delay)
                        ocr_names = []
                        match_results = []
                        found_ssr = None
                        epithet_ocr = None
                        best_match = None
                        best_score = 0
                        for ocr_attempt in range(4):
                            if not self.is_macro_started: return
                            screenshot = ImageGrab.grab()
                            region = screenshot.crop((x_name, y_name, x_name+w_name, y_name+h_name))
                            img = cv2.cvtColor(np.array(region), cv2.COLOR_RGB2BGR)
                            result = self.ocr_reader.readtext(img)
                            card_name = ' '.join([r[1] for r in result]).strip()
                            ocr_names.append(card_name)
                            for ssr_type, ssr_name in zip(ssr_reroll_types, ssr_reroll_names):
                                ocr_clean = card_name.replace(' ', '').lower()
                                cfg_clean = ssr_name.replace(' ', '').lower()
                                if '(' in ssr_name and ')' in ssr_name:
                                    epithet_region = self.safe_check.get('found_ssr_card_epithet', None)
                                    if epithet_region:
                                        x_epi, y_epi, w_epi, h_epi = epithet_region
                                        region_epi = screenshot.crop((x_epi, y_epi, x_epi+w_epi, y_epi+h_epi))
                                        img_epi = cv2.cvtColor(np.array(region_epi), cv2.COLOR_RGB2BGR)
                                        result_epi = self.ocr_reader.readtext(img_epi)
                                        epithet_ocr = ' '.join([r[1] for r in result_epi]).strip()
                                        start = ssr_name.find('(') + 1
                                        end = ssr_name.find(')')
                                        config_epithet = ssr_name[start:end].replace(' ', '').lower()
                                        epithet_ocr_clean = epithet_ocr.replace(' ', '').lower()
                                        score_name = self.fuzzy_ratio(ocr_clean, cfg_clean.split('(')[0])
                                        score_epithet = self.fuzzy_ratio(config_epithet, epithet_ocr_clean)
                                        print(f"[DEBUG] Name OCR: '{ocr_clean}', Config: '{cfg_clean.split('(')[0]}', Score: {score_name:.2f}")
                                        print(f"[DEBUG] Epithet OCR: '{epithet_ocr_clean}', Config: '{config_epithet}', Score: {score_epithet:.2f}")
                                        combined_score = (score_name + score_epithet) / 2
                                        if not epithet_ocr or score_epithet < 0.5:
                                            print(f"[Warning] Epithet OCR for {ssr_name} is empty or low confidence: '{epithet_ocr}' (score={score_epithet:.2f})")
                                        if combined_score > best_score and score_epithet > 0.7 and score_name > 0.7:
                                            best_score = combined_score
                                            best_match = (ssr_type, ssr_name)
                                else:
                                    score = self.fuzzy_ratio(ocr_clean, cfg_clean)
                                    if score > best_score and score > 0.7:
                                        best_score = score
                                        best_match = (ssr_type, ssr_name)
                        if best_match:
                            match_results.append(best_match)
                        found_types = [m[0] for m in match_results if m]
                        found_names = [m[1] for m in match_results if m]
                        if found_types and found_names:
                            found_ssr = (found_names[0], found_types[0])
                            print(f'Best match for {slot_key}: {found_types[0]}_{found_names[0]} (OCRs: {ocr_names})')
                        else:
                            print(f'No SSR config match for {slot_key} (OCRs: {ocr_names})')
                        if found_ssr:
                            key = f'{found_ssr[0]} ({found_ssr[1]})'
                            ssr_this_pull[key] = ssr_this_pull.get(key, 0) + 1
                            total_ssr_counter[key] = total_ssr_counter.get(key, 0) + 1
                        for _ in range(5):
                            if not self.is_macro_started: return
                            self.Global_MouseClick(x_fwd, y_fwd)
                            time.sleep(0.25 + delay)
                if ssr_this_pull:
                    ssr_str = ', '.join([f'x{v} {k}' for k, v in ssr_this_pull.items()])
                    print(f'#{pull_num+1} Specific rolled SSR card: {ssr_str}')
                else:
                    print(f'#{pull_num+1} Specific rolled SSR card: None')
                screenshot_path = None
                try:
                    screenshot = ImageGrab.grab()
                    width = 640
                    w, h = screenshot.size
                    if w > width:
                        new_h = int(h * (width / w))
                        screenshot = screenshot.resize((width, new_h), Image.LANCZOS)
                    screenshot_dir = os.path.join(os.getcwd(), 'umapyoi_screenshots')
                    if not os.path.exists(screenshot_dir):
                        os.makedirs(screenshot_dir)
                    screenshot_path = os.path.join(screenshot_dir, f'scout_result_{self.reroll_run_count}_{pull_num+1}.png')
                    screenshot.save(screenshot_path, 'PNG')
                except Exception as e:
                    print(f'Failed to capture screenshot for webhook: {e}')
                msg = f"#{pull_num+1} Scout Results: SSR: {ssr_count}, SR: {sr_count}, R: {r_count} (Run {self.reroll_run_count})"
                description = None
                ssr_details = None
                if ssr_this_pull:
                    ssr_details = '\n'.join([f'x{v} {k}' for k, v in ssr_this_pull.items()])
                send_discord_notification(msg, self.webhook_url.get(), self.ping_user_id.get(), description=description, ping_on_embed=False, ssr_details=ssr_details, screenshot_path=screenshot_path)
                if screenshot_path and os.path.exists(screenshot_path):
                    try:
                        os.remove(screenshot_path)
                    except Exception:
                        pass
            if total_ssr_counter:
                total_str = ', '.join([f'x{v} {k}' for k, v in total_ssr_counter.items()])
                print(f'Total rolled card for #{self.reroll_run_count} scout: {total_str}')
            else:
                print(f'Total rolled card for #{self.reroll_run_count} scout: None')
            ssr_met = False
            config = load_config()
            if config.has_section('SSR_REROLL'):
                for k, v in config.items('SSR_REROLL'):
                    if k == 'password':
                        continue
                    try:
                        enabled, min_count = v.split(',')
                        enabled = enabled.strip() == '1'
                        min_count = int(min_count.strip())
                        if '_' in k:
                            _, ssr_name = k.split('_', 1)
                        else:
                            ssr_name = k
                        for key in total_ssr_counter:
                            if ssr_name.replace(' ', '').lower() in key.replace(' ', '').lower():
                                if enabled and total_ssr_counter[key] >= min_count:
                                    ssr_met = True
                                    break
                        if ssr_met:
                            break
                    except Exception:
                        continue
            summary_title = f'#{self.reroll_run_count}'
            summary_desc = total_str if total_ssr_counter else 'None'
            summary_ssr_details = '\n'.join([f'x{v} {k}' for k, v in total_ssr_counter.items()]) if total_ssr_counter else None
            send_discord_notification(summary_title, self.webhook_url.get(), self.ping_user_id.get(), status='summary', description=summary_desc, ping_on_embed=ssr_met, ssr_details=summary_ssr_details, is_summary=True)
            if ssr_met:
                if not self.is_macro_started: return
                print('SSR reroll requirement met. Proceeding to link account process.')
                self.link_met_required_account(total_ssr_counter)
            else:
                print('SSR reroll requirement NOT met. Returning to title screen.')
                config = load_config()
                title_btn = config.get('INGAME_MENU', 'title_screen_btn', fallback=None)
                if title_btn:
                    x, y = map(int, title_btn.split(','))
                    print(f'Clicking Title Screen Button at ({x}, {y})')
                    self.Global_MouseClick(x, y)
                else:
                    print('Title Screen Button not assigned in INGAME_MENU!')
            self.reroll_run_count += 1
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            line_number = exc_tb.tb_lineno if exc_tb else 'unknown'
            tb_str = traceback.format_exc()
            self.logger.error(f"Error in scout_reroll_loop: {e} (Line {line_number})\n{tb_str}")
            print(f"Error in scout_reroll_loop: {e} (Line {line_number})\n{tb_str}")

    def fuzzy_ratio(self, a, b):
        a = a.replace(' ', '').lower()
        b = b.replace(' ', '').lower()
        return difflib.SequenceMatcher(None, a, b).ratio()

    def link_met_required_account(self, total_ssr_counter):
        try:
            if not self.is_macro_started: return
            print('link_met_required_account called with:', total_ssr_counter)
            config = load_config()
            delay_ms = config.getint('REROLL_SETTINGS', 'general_delay', fallback=800)
            delay = max(0.01, delay_ms / 1000.0)
            link_keys = [
                'profile_btn', 'next_btn', 'copy_trainer_id_btn', 'data_link_btn', 'data_link_confirm_btn',
                'set_link_password_btn', 'password_input_box', 'password_confirm_input_box', 'privacy_policy_tick_box', 'ok_btn'
            ]
            coords = {}
            if config.has_section('LINK_ACCOUNT'):
                for k in link_keys:
                    v = config.get('LINK_ACCOUNT', k, fallback=None)
                    if v:
                        try:
                            x, y = map(int, v.split(','))
                            coords[k] = (x, y)
                        except Exception:
                            coords[k] = (0, 0)
            time.sleep(2.3 + delay)
            if not self.is_macro_started: return
            if 'profile_btn' in coords:
                print(f'Clicking Profile Button at {coords["profile_btn"]}')
                self.Global_MouseClick(*coords['profile_btn'])
            time.sleep(3.5 + delay)
            for i in range(2):
                if not self.is_macro_started: return
                if 'next_btn' in coords:
                    print(f'Clicking Next Button ({i+1}/2) at {coords["next_btn"]}')
                    self.Global_MouseClick(*coords['next_btn'])
                time.sleep(1.2 + delay)
            time.sleep(2.3 + delay)
            if not self.is_macro_started: return
            forward_icon = config.get('INGAME_MENU', 'forward_icon', fallback=None)
            if not forward_icon:
                print('Config value for forward_icon in INGAME_MENU is missing!')
                return
            x_fwd, y_fwd = map(int, forward_icon.split(','))
            trainer_id = None
            if 'copy_trainer_id_btn' in coords:
                print(f'Clicking Copy Trainer ID Button at {coords["copy_trainer_id_btn"]}')
                self.Global_MouseClick(*coords['copy_trainer_id_btn'])
                time.sleep(1.5)
                try:
                    trainer_id = pyperclip.paste().strip()
                    print(f"Trainer ID (from clipboard): {trainer_id}")
                except Exception as e:
                    print(f"Failed to read Trainer ID from clipboard: {e}")
                time.sleep(0.75)
                for i in range(6):
                    if not self.is_macro_started: return
                    self.Global_MouseClick(x_fwd, y_fwd)
                    time.sleep(0.35 + delay)
            time.sleep(1.2 + delay)
            if not self.is_macro_started: return
            if 'data_link_btn' in coords:
                print(f'Clicking Data Link Button at {coords["data_link_btn"]}')
                self.Global_MouseClick(*coords['data_link_btn'])
            time.sleep(2.5 + delay)
            if not self.is_macro_started: return
            if 'data_link_confirm_btn' in coords:
                print(f'Clicking Data Link Confirm Button at {coords["data_link_confirm_btn"]}')
                self.Global_MouseClick(*coords['data_link_confirm_btn'])
            time.sleep(2 + delay)
            if not self.is_macro_started: return
            if 'set_link_password_btn' in coords:
                print(f'Clicking Set Link Password Button at {coords["set_link_password_btn"]}')
                self.Global_MouseClick(*coords['set_link_password_btn'])
            time.sleep(2.25 + delay)
            if not self.is_macro_started: return
            if 'data_link_confirm_btn' in coords:
                print(f'Clicking Data Link Confirm Button (again) at {coords["data_link_confirm_btn"]}')
                self.Global_MouseClick(*coords['data_link_confirm_btn'])
            time.sleep(2 + delay)
            if not self.is_macro_started: return
            password = config.get('SSR_REROLL', 'password', fallback='Qwerty4321')
            for i in range(3):
                if not self.is_macro_started: return
                if 'password_input_box' in coords:
                    print(f'Clicking Password Input Box ({i+1}/3) at {coords["password_input_box"]}')
                    self.Global_MouseClick(*coords['password_input_box'])
                    time.sleep(0.4 + delay)
            print(f'Typing password: {password}')
            autoit.send(password)
            time.sleep(0.8 + delay)
            for i in range(3):
                if not self.is_macro_started: return
                if 'password_confirm_input_box' in coords:
                    print(f'Clicking Password Confirm Input Box ({i+1}/3) at {coords["password_confirm_input_box"]}')
                    self.Global_MouseClick(*coords['password_confirm_input_box'])
                    time.sleep(0.4 + delay)
            print(f'Typing password (confirm): {password}')
            autoit.send(password)
            time.sleep(2 + delay)
            if not self.is_macro_started: return
            if 'privacy_policy_tick_box' in coords:
                print(f'Clicking Privacy Policy Tick Box at {coords["privacy_policy_tick_box"]}')
                self.Global_MouseClick(*coords['privacy_policy_tick_box'])
            time.sleep(2.2 + delay)
            if not self.is_macro_started: return
            if 'ok_btn' in coords:
                print(f'Clicking OK Button at {coords["ok_btn"]}')
                self.Global_MouseClick(*coords['ok_btn'])
            print('Link account process completed.')
            if trainer_id:
                try:
                    rotate_accounts_met_log()
                    with open('met_required_accounts.txt', 'a', encoding='utf-8') as f:
                        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        ssr_str = ', '.join([f'x{v} {k}' for k, v in total_ssr_counter.items()])
                        f.write(f'[{now}] Trainer ID: {trainer_id} | SSRs: {ssr_str}\n')
                except Exception as e:
                    print(f"Failed to log met account info: {e}")
            time.sleep(2.2 + delay)
            if not self.is_macro_started: return
            title_screen_btn = config.get('INGAME_MENU', 'title_screen_btn', fallback=None)
            x_btn, y_btn = map(int, title_screen_btn.split(','))
            for i in range(3):
                self.Global_MouseClick(x_btn, y_btn)
                time.sleep(1.2)
            # Send Discord notification with Trainer ID after linking (no SSR Cards Pulled field)
            screenshot_path = capture_and_resize_screenshot('linked_account')
            send_discord_notification(
                message=f'Linked account completed (Run {self.reroll_run_count})',
                webhook_url=self.webhook_url.get(),
                ping_user_id=self.ping_user_id.get(),
                status='linked',
                description=f'**Trainer ID:** `{trainer_id}`' if trainer_id else None,
                ping_on_embed=False,
                ssr_details=None,
                is_summary=False,
                screenshot_path=screenshot_path
            )
        except Exception as e:
            import sys, traceback
            exc_type, exc_obj, exc_tb = sys.exc_info()
            line_number = exc_tb.tb_lineno if exc_tb else 'unknown'
            tb_str = traceback.format_exc()
            self.logger.error(f"Error in link_met_required_account: {e} (Line {line_number})\n{tb_str}")
            print(f"Error in link_met_required_account: {e} (Line {line_number})\n{tb_str}")

    def start_connection_error_failsafe(self):
        def check_loop():
            while self.is_macro_started:
                try:
                    config = load_config()
                    if not config.has_section('SAFE_CHECK'):
                        time.sleep(5)
                        continue
                    conn_error_region = config.get('SAFE_CHECK', 'conn_error_region', fallback=None)
                    title_screen_btn = config.get('SAFE_CHECK', 'title_screen_btn', fallback=None)
                    if not conn_error_region or not title_screen_btn:
                        time.sleep(5)
                        continue
                    x, y, w, h = map(int, conn_error_region.split(','))
                    screenshot = ImageGrab.grab()
                    region = screenshot.crop((x, y, x+w, y+h))
                    img = cv2.cvtColor(np.array(region), cv2.COLOR_RGB2BGR)
                    result = self.ocr_reader.readtext(img)
                    text = ' '.join([r[1] for r in result]).strip().lower()
                    if 'connection error' in text:
                        print('[Failsafe] Connection error detected! Stopping macro and returning to title screen.')
                        self.is_macro_started = False
                        x_btn, y_btn = map(int, title_screen_btn.split(','))
                        for i in range(4):
                            self.Global_MouseClick(x_btn, y_btn)
                            time.sleep(1.2)
                        screenshot_path = capture_and_resize_screenshot('connection_error')
                        send_discord_notification(
                            'Connection error detected! Macro stopped and returned to title screen.',
                            self.webhook_url.get(),
                            self.ping_user_id.get(),
                            status=None,
                            screenshot_path=screenshot_path
                        )
                        break
                except Exception as e:
                    print(f'[Failsafe] Error in connection error check: {e}')
                time.sleep(5)
        self.conn_error_thread = threading.Thread(target=check_loop, daemon=True)
        self.conn_error_thread.start()

    def macro_sleep(self, multiplier=1.0):
        config = getattr(self, '_cached_config', None)
        if config is None:
            config = load_config()
            self._cached_config = config
        delay_ms = 800
        if config.has_section(ConfigSection.REROLL_SETTINGS.value):
            delay_ms = config.getint(ConfigSection.REROLL_SETTINGS.value, 'general_delay', fallback=800)
        time.sleep(max(0.01, delay_ms / 1000.0 * multiplier))

def send_discord_notification(message, webhook_url, ping_user_id=None, status=None, description=None, ping_on_embed=False, ssr_details=None, is_summary=False, screenshot_path=None):
    # status: 'start' or 'stop' or None
    now = datetime.datetime.now().strftime('%H:%M:%S')  # hh:mm:ss
    # Use more vibrant color for SSR pulls and summary
    if status == 'start':
        title = f'**__Reroll started__** ({now})'
        color = 0x57F287  # green
    elif status == 'stop':
        title = f'**__Reroll stopped__** ({now})'
        color = 0xED4245  # red
    elif status == 'summary':
        # Extract run number from message (e.g., '#1') or fallback to message
        import re
        m = re.search(r'#(\d+)', message)
        run_num = m.group(1) if m else message
        title = f'**__Total rolled SSR card for (Run {run_num})__**'
        color = 0xFFD700  # gold
    else:
        title = f'**{message}**'
        color = 0x7289DA  # Discord blurple
    content = ''
    if ping_on_embed and ping_user_id:
        content = f'<@{ping_user_id}>'
    embed = {
        "title": title,
        "color": color,
        "footer": {
            "text": "Umamusume Reroll Macro (Discord: @notwindybee_)",
            "icon_url": "https://i.postimg.cc/C5TXvNXH/image.png"
        }
    }
    # Add SSR details if present
    if ssr_details:
        ssr_lines = [f'**• {line}**' for line in ssr_details.split('\n') if line.strip()]
        embed["fields"] = [
            {"name": "SSR Cards Pulled", "value": '\n'.join(ssr_lines), "inline": False}
        ]
    # For summary, remove redundant description and only keep SSR Cards Pulled
    if status == 'summary':
        # Remove description (the summary_desc) if present
        if 'description' in embed:
            del embed['description']
    elif description:
        embed["description"] = description
    # If screenshot, embed it inside the embed
    files = None
    if screenshot_path and os.path.exists(screenshot_path):
        embed["image"] = {"url": f"attachment://{os.path.basename(screenshot_path)}"}
        try:
            with open(screenshot_path, "rb") as f:
                files = {"file": (os.path.basename(screenshot_path), f, "image/png")}
                response = requests.post(webhook_url, data={"payload_json": json.dumps({"content": content, "embeds": [embed]})}, files=files)
                response.raise_for_status()
            # Clean up screenshot after sending
            os.remove(screenshot_path)
            return
        except Exception as e:
            print(f"Failed to send screenshot with webhook: {e}")
    try:
        response = requests.post(webhook_url, json={"content": content, "embeds": [embed]})
        response.raise_for_status()
    except Exception as e:
        print(f"Failed to send Discord notification: {e}")

def rotate_accounts_met_log():
    log_path = 'met_required_accounts.txt'
    max_size = 10 * 1024 * 1024
    if os.path.exists(log_path) and os.path.getsize(log_path) > max_size:
        for i in range(2, 0, -1):
            prev = f'met_required_accounts.txt.{i-1}' if i > 1 else 'met_required_accounts.txt'
            next = f'met_required_accounts.txt.{i}'
            if os.path.exists(prev):
                shutil.move(prev, next)
        shutil.move('met_required_accounts.txt', 'met_required_accounts.txt.1')


if __name__ == '__main__':
    root = Tk()
    app = UMAPanel(root)
    root.mainloop()
