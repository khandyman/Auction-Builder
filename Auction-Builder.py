import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk
import ttkbootstrap.dialogs
import tksheet

import sys
import urllib.error
import urllib3
import platform
import threading
import math
import os
import requests


# ----------------------------------------
# ----------- software license -----------
# ----------------------------------------

# MIT License
#
# Copyright (c) 2024 khandyman
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# ----------------------------------------
# ----------- import functions -----------
# ----------------------------------------


# obtain item list and get prices to populate in sheet
def import_items():
    clear_form()
    if check_file(inventory_path):
        item_list = build_item_list()

        # update total items field with the number of items to import
        tot_items.set(str(len(item_list)))

        # start a thread so that sheet can be updated while importing items
        # this acts as a sort of progress bar
        thread = threading.Thread(target=lambda: build_price_list(item_list))
        thread.start()

        set_sheet_columns()
    else:
        show_app_info('Zeal outputfile specified does not exist.\n'
                      'Please check in settings and try again.',
                      'Missing File', 'error')


# read in user's zeal outputfile and built list of items
def build_item_list():
    with open(inventory_path) as file:
        line = file.readline()
        raw_item_list = []
        def_item_list = []

        while line:
            split = line.split('\t')
            line = file.readline()

            # make sure current item should not be excluded
            if check_exclusion(split[1]):
                continue

            # filter out undesired lines from outputfile
            if split[0] != 'Location' and split[0][:4] != 'Bank' and split[1] != 'Currency' and split[1] != 'Empty':
                if int(split[4]) < 1 and split[1][:4] != 'Nili' and split[1][:4] != 'Word' and split[1][:4] != 'Sali':
                    if split[1][:4] != 'Part' and split[1] != 'Empty':
                        # if all these filters pass, then grab item name and id and add to list
                        item_and_id = [split[1], split[2]]
                        raw_item_list.append(item_and_id)

    # check all items in raw list for duplicates;
    # if no duplicates found, add item to final list
    for raw_item in raw_item_list:
        duplicate = False

        for def_item in def_item_list:
            if raw_item == def_item:
                duplicate = True

        if duplicate is False:
            def_item_list.append(raw_item)

    return def_item_list


# get prices for each item from eq tunnel auctions web site
def build_price_list(def_item_list):
    # turn off ui buttons so user cannot click until finished
    disable_ui()

    def_price_list = []
    auction_list = []

    for index, item_plus_id in enumerate(def_item_list):
        # begin url string
        url = 'https://eqtunnelauctions.com/item.php?itemstr='
        # split item into individual words
        words = item_plus_id[0].split()
        # count the words
        num_words = len(words)
        item_string = ''

        curr_item.set(str(index + 1))

        # take the individual words of the item name, and
        # append them together with plus symbols between
        for i in range(num_words):
            if i == 0:
                item_string = f'{words[i]}'

                # if this is a spell, insert character code
                # for a colon
                if words[i][5:] == 'Spell':
                    item_string = item_string + '%3A'
            else:
                item_string = item_string + f'+{words[i]}'

        # add the formated item name to the url string
        url = url + item_string
        # run web scraper to get list of prices
        try:
            auction_list = scrape_page(url)
        except requests.exceptions.ReadTimeout as error:
            show_app_info(f'ReadTimeout error encountered: {error}\nContinuing to scan...',
                          'ReadTimeout Error', 'error')
        except urllib.error.HTTPError as error:
            show_app_info(f'HTTP error encountered: {error}\nContinuing to scan...',
                          'HTTP Error', 'error')
        except urllib.error.URLError as error:
            show_app_info(f'URL SSL error encountered: {error}\nContinuing to scan...',
                          'URL Error', 'error')

        # if auctions were found, run price calculator
        if len(auction_list) > 0:
            price = calculate_price(auction_list)
        # otherwise, just flag item as ?
        else:
            price = 'unknown'

        # append everything to the master price list
        def_price_list.append([item_plus_id[0], item_plus_id[1], price])

        # allow UI updates outside of thread
        app.after(0, update_sheet([item_plus_id[0], item_plus_id[1], price, False]))

    enable_ui()

    return def_price_list


# take in a url string, scrap html code from page,
# and assemble list of prices
def scrape_page(url):
    # open web page, read in the html, and translate into text
    page = requests.get(url, verify=False)

    html = page.text
    auction_list = []
    # look for the  start of the list of previous auctions
    data_start = html.find("data: [")

    # if data start point was found...
    if data_start != -1:
        # look for the end point
        data_end = html.find('],', data_start)
        # slice out the string of prices from the html code
        data_list = html[data_start + 7:data_end]
        # get rid of quotation marks
        num_list = data_list.replace("\"", "")
        # and split the entries into individual numbers
        separated_list = num_list.split(",")

        # add the numbers to the auction list
        for num in separated_list:
            auction_list.append(num)

    return auction_list


# take in a list of numbers and calculate
# an average price based on user settings
def calculate_price(def_auction_list):
    sum_price = 0
    price_count = 0
    divisor = int(auctions_count)

    # if there are fewer price points then
    # the user's desired number of auctions,
    # set the divisor to the number of prices
    # found
    if len(def_auction_list) < divisor:
        divisor = len(def_auction_list)

    # loop through price list, summing until
    # number of sums matches divisor
    for price in def_auction_list:
        sum_price = sum_price + int(price)
        price_count += 1

        if price_count == divisor:
            break

    # calculate average price by dividing sum by divisor,
    # then rounding it up to the nearest multiple of 50
    avg_price = round_to_50(sum_price // divisor)

    return avg_price


# take in a number and ceiling it
# up to the nearest multiple of 50
def round_to_50(num, base=50):
    return base * math.ceil(num / base)


# ----------------------------------------
# ----------- export functions -----------
# ----------------------------------------


# use the contents of the sheet to build a list of items
# to insert into character ini file
def build_file_list(price_list):
    if not check_file(character_path):
        show_app_info('Character ini file specified does not exist.\n'
                      'Please check in settings and try again.',
                      'Missing File', 'warning')
        return

    def_price_list = []

    # make sure there's data in the sheet first
    if len(sheet.get_sheet_data()) < 1:
        ttk.dialogs.Messagebox.show_error('Please import item data first.', 'No Data')
        return

    # build new item list with only items not
    # marked for exclusion, so that len of list
    # is accurate; items marked for exclusion
    # get added to exclusions list
    for item in price_list:
        if item[3] is False:
            def_price_list.append(item)
        else:
            add_exclusion(item[0])

    # print(def_price_list)
    # open the ini file
    with open(character_path) as file:
        file_contents = []
        item_index = 0
        line_num = 1
        button_num = int(hotkey_button)
        line_to_write = ''
        new_button = True
        total_buttons = math.ceil(len(def_price_list) / 30)

        # read in entire ini file, excluding entries that match
        # user's desired hotkey page and button, or contain
        # the word auction
        for line in file:
            skip_line = False

            # each macro can store 30 items, so ceil the quotient of
            # the item list count and 30, then loop to eliminate not
            # only the users start button number, but any additional
            # buttons that will be created
            for i in range(1, total_buttons + 1):
                if f'Page{hotkey_page}Button{i}' in line or 'uction' in line:
                    skip_line = True

            # if not skipping, then add line to list
            if skip_line is False:
                file_contents.append(line)

    # now loop through item list to add lines for auction
    # macros to bottom of file (EQ doesn't care what order
    # they are in)
    for index, item in enumerate(def_price_list):
        # if this is a new button, create the name and color lines,
        # then flip flag to false
        if new_button is True:
            write_line(file_contents, f'Page{hotkey_page}Button{button_num}Name=Auction{button_num}')
            write_line(file_contents, f'Page{hotkey_page}Button{button_num}Color=0')
            new_button = False

        # if this is the first auction line of the macro
        if item_index == 0:
            line_to_write = f'Page{hotkey_page}Button{button_num}Line{line_num}=/auction WTS '

        # format the item id to be 6 digits long
        if len(item[1]) == 5:
            format_id = f'0{item[1]}'
        else:
            format_id = f'00{item[1]}'

        # append the current item name, id, and price to the auction line,
        # including DC2 tags so that EQ will read it as an item link
        line_to_write = line_to_write + f'{format_id} {item[0]} {item[2]}'

        # if there are more items to add to this line,
        # then append a comma to the line
        if item_index < 5 and index < len(def_price_list) - 1:
            line_to_write = f'{line_to_write}, '

        # if current auction line contains 6 items,
        # then add it to file list and reset variables
        if item_index == 5:
            write_line(file_contents, line_to_write)

            line_to_write = ''
            item_index = 0
            line_num += 1
        else:
            item_index += 1

        # if there are already 5 lines of auctions in the
        # current macro button, flag to create a new
        # button on the next loop iteration
        if line_num == 6:
            new_button = True
            button_num += 1
            line_num = 1

    # a special clause for when there are no more items,
    # but the current macro line does not contain 6 items,
    # we still want the current item added
    if len(line_to_write) > 1:
        write_line(file_contents, line_to_write)

    # finally, call functions to perform the file write
    # operation, clear the sheet, and notify user
    write_new_file(file_contents)
    clear_form()
    show_app_info('Auction macro(s) successfully created in .ini file.\n'
                  'Please log into EverQuest to see the changes.',
                  'Write Successful', 'info')


# add line parameter, with a newline, to the items list
def write_line(items, line):
    line_to_write = line + f'\n'
    items.append(line_to_write)

    return items


# take the file list, with auction macro(s), assembled
# above and write it to disk, over-writing existing file
def write_new_file(file_contents):
    global character_path

    with open(character_path, 'w') as file:
        for line in file_contents:
            file.write(line)


# ----------------------------------------
# --------- exclusion functions ----------
# ----------------------------------------

# compare item parameter to exclusions list
def check_exclusion(item):
    with open('settings') as file:
        line = file.readline()

        while line:
            if item in line:
                return True

            line = file.readline()

    return False


# add item parameter to in memory exclusions list,
# then append to settings file
def add_exclusion(item):
    exclusions_list.append(item)

    with open('settings', 'a') as file:
        file.write(f'\n{item}')


# ----------------------------------------
# ------------- GUI functions ------------
# ----------------------------------------

# set next sheet row contents to list_item parameter
def update_sheet(list_item):
    index = sheet.get_total_rows()
    sheet.set_data(index, data=list_item)


# clear all data and reset column widths in sheet
def clear_form():
    for i in range(sheet.get_total_rows() - 1, -1, -1):
        sheet.delete_row(i)

    set_sheet_columns()
    curr_item.set('')
    tot_items.set('')


# set sheet columns to their proper widths
def set_sheet_columns():
    sheet.column_width(column=0, width=220)
    sheet.column_width(column=1, width=70)
    sheet.column_width(column=2, width=70)
    sheet.column_width(column=3, width=60)
    sheet.set_options(default_row_height=30)


# turn off all UI buttons
def disable_ui():
    file_menu.entryconfig('Settings', state='disabled')
    # settings_button.configure(state=ttk.DISABLED)
    import_button.configure(state=ttk.DISABLED)
    save_button.configure(state=ttk.DISABLED)


# turn on all UI buttons
def enable_ui():
    file_menu.entryconfig('Settings', state='normal')
    # settings_button.configure(state=ttk.NORMAL)
    import_button.configure(state=ttk.NORMAL)
    save_button.configure(state=ttk.NORMAL)


# ----------------------------------------
# --------- validation functions ---------
# ----------------------------------------

def show_app_info(msg_content, msg_title, msg_type):
    match msg_type:
        case 'error':
            message = ttk.dialogs.Messagebox.show_error(msg_content, msg_title)
        case 'warning':
            message = ttk.dialogs.Messagebox.show_warning(msg_content, msg_title)
        case 'info':
            message = ttk.dialogs.Messagebox.show_info(msg_content, msg_title)
        case 'yesno':
            message = ttk.dialogs.Messagebox.yesno(msg_content, msg_title, alert=True)
        case _:
            message = ''

    return message


def check_file(file):
    if os.path.exists(file):
        return True
    else:
        return False


# ----------------------------------------
# -------- setup/help functions ----------
# ----------------------------------------

# read in settings file and assign values to globals
def read_settings():
    global inventory_path, character_path, hotkey_page, hotkey_button, auctions_count, exclusions_list
    read = False

    # if the settings file doesn't exist, open
    # settings window and exit
    if not os.path.isfile('settings'):
        open_settings(False)
        return

    # if settings does exist
    with open('settings') as file:
        settings_count = 0
        line = file.readline()

        while line:
            # if flag for reading exclusions has been
            # flipped, assume all remaining lines need
            # added to exclusions list
            if read is True:
                exclusions_list.append(line.strip())

            # grab slice of current line to the right
            # of the equals sign
            setting = line[(line.find('=') + 1):]
            line_slice = line[:4]

            # determine which setting the current
            # line is, and store parameter to global
            match line_slice:
                case 'page':
                    hotkey_page = setting.strip()
                    settings_count += 1
                case 'butt':
                    hotkey_button = setting.strip()
                    settings_count += 1
                case 'auct':
                    auctions_count = setting.strip()
                    settings_count += 1
                case 'outp':
                    inventory_path = setting.strip()
                    settings_count += 1
                case 'mule':
                    character_path = setting.strip()
                    settings_count += 1
                case '[exc':
                    read = True
                    settings_count += 1

            line = file.readline()

        # if all cases above have not been found, assume
        # there is an error in the settings file and
        # pop settings window so user can fix
        if settings_count != 6:
            show_app_info('Invalid Settings File.\nPlease reconfigure.',
                          'Invalid Settings', 'warning')
            open_settings(False)


def open_readme():
    # ------------- readme window -------------
    readme = tk.Toplevel(app)
    readme.title('Readme')
    readme.transient(app)
    parent_x_pos = app.winfo_rootx()
    parent_y_pos = app.winfo_rooty()
    readme.geometry('520x620+%d+%d' % (parent_x_pos - adjust_x_pos, parent_y_pos - adjust_y_pos))
    readme.iconbitmap('EverQuest.ico')

    # ------------- readme layout (frames/separators/button) -------------
    readme_title = ttk.Label(readme, font=info_title_large)
    readme_title.configure(text='Auction Builder Readme')
    readme_title.pack(pady=10, padx=5, fill='both')

    # ------------- basic use widgets -------------
    use_title = ttk.Label(readme, font=info_title_small)
    use_title.configure(text='Basic Use')
    use_title.pack(pady=5, padx=5, fill='both')

    # basic_use_frame = ttk.Frame(readme)
    # basic_use_frame.pack(padx=5)

    use_text_1 = ttk.Label(readme, font=label_font_small)
    use_text_1.configure(text=' - To use Auction Builder, click the Import button, then wait while the program '
                              'assembles a list of items and prices, which it will print out.', wraplength=500)
    use_text_1.pack(pady=5, padx=10, fill='x')

    use_text_2 = ttk.Label(readme, font=label_font_small)
    use_text_2.configure(text=' - The list may then be examined and the calculated prices can be adjusted as '
                              'desired.', wraplength=500)
    use_text_2.pack(pady=5, padx=10, fill='x')

    use_text_3 = ttk.Label(readme, font=label_font_small)
    use_text_3.configure(text=' - Additionally, if an item should be excluded from this and any future macros, '
                              'check the exclude box for that item.', wraplength=500)
    use_text_3.pack(pady=5, padx=5, fill='x')

    use_text_4 = ttk.Label(readme, font=label_font_small)
    use_text_4.configure(text=' - Finally, click Save and the list will be written to the .ini file', wraplength=500)
    use_text_4.pack(pady=5, padx=5, fill='x')

    # ------------- settings widgets -------------
    settings_title = ttk.Label(readme, font=info_title_small)
    settings_title.configure(text='Settings')
    settings_title.pack(pady=5, padx=5, fill='both')

    settings_text_1 = ttk.Label(readme, font=label_font_small)
    settings_text_1.configure(text=' - Hotkey Starting Location: specify the page and button where auction macros '
                                   'should begin.  Each macro will store up to 30 items. If necessary, multiple '
                                   'macros will be created, incrementing the button by 1 each time.', wraplength=500)
    settings_text_1.pack(pady=5, padx=10, fill='x')

    settings_text_2 = ttk.Label(readme, font=label_font_small)
    settings_text_2.configure(text=' - Auctions to Count: indicate the desired number of auctions from '
                                   'www.eqtunnelauctions.com to use in calculating an average price. If an item has '
                                   'less than the number specified, the program will use as many as it can find. If '
                                   'an item has no auction data, the price will be \'unknown\'.', wraplength=500)
    settings_text_2.pack(pady=5, padx=10, fill='x')

    settings_text_3 = ttk.Label(readme, font=label_font_small)
    settings_text_3.configure(text=' - Outputfile Path: this is the path to a character\'s Zeal outputfile. '
                                   'Simply click the text field to change the file path.', wraplength=500)
    settings_text_3.pack(pady=5, padx=10, fill='x')

    settings_text_4 = ttk.Label(readme, font=label_font_small)
    settings_text_4.configure(text=' - Mule Ini Path: this is the path to a character\'s EQ .ini file. '
                                   'Simply click the text field to change the file path.', wraplength=500)
    settings_text_4.pack(pady=5, padx=10, fill='x')

    settings_text_4 = ttk.Label(readme, font=label_font_small)
    settings_text_4.configure(text=' - Item Exclusions: if items are present in a character\'s inventory that '
                                   'should not be sold, they can be marked for exclusion. Items in this list '
                                   'will be ignored.', wraplength=500)
    settings_text_4.pack(pady=5, padx=10, fill='x')


def open_about():
    # ------------- settings window -------------
    about = tk.Toplevel(app)
    about.title('About')
    about.transient(app)
    about.grab_set()
    parent_x_pos = app.winfo_rootx()
    parent_y_pos = app.winfo_rooty()
    about.geometry('300x200+%d+%d' % (parent_x_pos + 10, parent_y_pos + 55))
    about.iconbitmap('EverQuest.ico')

    # ------------- settings layout (frames/separators/button) -------------
    about_title = ttk.Label(about, font=info_title_large)
    about_title.configure(text='Auction Builder')
    about_title.pack(pady=20, padx=10, fill='both')

    about_version = ttk.Label(about, font=label_font_small)
    about_version.configure(text='Version 1.0\n\nCopyright (c) 2024 khandyman\nLicensed under the MIT License')
    about_version.pack(pady=20, padx=10, fill='both')


# -------------------------------------------------
# --------- settings window and functions ---------
# -------------------------------------------------

# open settings window, optional parameter is boolean
# to set whether user can freely click the close
# button to close the settings window
def open_settings(optional):
    global inventory_path, character_path, hotkey_page, hotkey_button, auctions_count

    # delete current selection from listbox
    def delete_exclusion():
        exclusions_listbox.delete('anchor')

    # add exclusions list to listbox
    def populate_exclusions():
        for item in exclusions_list:
            exclusions_listbox.insert('end', item)

    # pop an open file dialog to allow user to
    # select new outputfile
    def change_outputfile(event):
        file_path = filedialog.askopenfilename(title="Select Outputfile", filetypes=[("Text files", "*Inventory.txt")])

        if file_path:
            outputfile_path.set(file_path)

    # pop an open file dialog to allow user to
    # select new character ini file
    def change_mule_ini(event):
        file_path = filedialog.askopenfilename(title="Select Mule Ini", filetypes=[("Text files", "*_pq.proj.ini")])

        if file_path:
            mule_ini_path.set(file_path)

    # determine correct behavior for close button;
    # either allow user to close if optional is true
    # (i.e., if the user clicked settings on their own),
    # or show a prompt if optional is false (i.e., the system
    # forced the settings window open); in this latter case,
    # if user clicks yes to close, kill the whole program
    def handle_close():
        if optional is True:
            settings.destroy()
        else:
            # confirm = ttk.dialogs.Messagebox.yesno(
            #     'You cannot run Auction Builder\nwithout setting it up first.'
            #     '\n\nAre you sure you want to exit?', 'Exit Warning', alert=True)
            confirm = show_app_info('You cannot run Auction Builder\nwithout setting it up first.'
                                    '\n\nAre you sure you want to exit?', 'Exit Warning', 'yesno')

            if confirm == 'Yes':
                sys.exit()

    # write all settings to settings file on disk
    def save_settings():
        global inventory_path, character_path, hotkey_page, hotkey_button, auctions_count, exclusions_list

        if validate_settings():
            inventory_path = outputfile_path.get()
            character_path = mule_ini_path.get()
            hotkey_page = page.get()
            hotkey_button = button.get()
            auctions_count = auctions.get()
            exclusions_list = []

            with open('settings', 'w') as file:
                file.write(f'[config]')
                file.write(f'\npage={page.get()}')
                file.write(f'\nbutton={button.get()}')
                file.write(f'\nauctions={auctions.get()}')
                file.write(f'\noutputfile={outputfile_path.get()}')
                file.write(f'\nmule_ini={mule_ini_path.get()}')
                file.write(f'\n[exclusions]')

                for i in range(exclusions_listbox.size()):
                    item = exclusions_listbox.get(i)
                    exclusions_list.append(item)
                    file.write(f'\n{item}')

            # after performing write, close settings window
            settings.destroy()

    # check that all fields have been filled out
    def validate_settings():
        if len(page.get()) < 1:
            show_settings_error('Please enter a hotkey page number.', 'No Page')
            return False

        if len(button.get()) < 1:
            show_settings_error('Please enter a hotkey button number.', 'No Button')
            return False

        if len(auctions.get()) < 1:
            show_settings_error('Please enter the number of auctions to use.', 'No Auctions')
            return False

        if len(outputfile_path.get()) < 1:
            show_settings_error('Please select a Zeal outputfile to use.', 'No Outputfile')
            return False

        if len(mule_ini_path.get()) < 1:
            show_settings_error('Please select a character ini file to write to.', 'No Ini File')
            return False

        return True

    def show_settings_error(msg_content, msg_title):
        ttk.dialogs.Messagebox.show_error(msg_content, msg_title)

    outputfile_path = ttk.StringVar()
    mule_ini_path = ttk.StringVar()
    page = ttk.StringVar()
    button = ttk.StringVar()
    auctions = ttk.StringVar()

    outputfile_path.set(inventory_path)
    mule_ini_path.set(character_path)
    page.set(hotkey_page)
    button.set(hotkey_button)
    auctions.set(auctions_count)

    # ------------- settings window -------------
    settings = tk.Toplevel(app)
    settings.title('Settings')
    settings.transient(app)
    settings.grab_set()
    parent_x_pos = app.winfo_rootx()
    parent_y_pos = app.winfo_rooty()
    settings.geometry('480x500+%d+%d' % (parent_x_pos + 10, parent_y_pos + 55))
    settings.iconbitmap('EverQuest.ico')
    settings.protocol("WM_DELETE_WINDOW", handle_close)

    # ------------- settings layout (frames/separators/button) -------------
    numbers_frame = ttk.Frame(settings, width=20)
    numbers_frame.pack(pady=5, padx=10)

    top_separator = ttk.Separator(settings, orient='horizontal')
    top_separator.pack(pady=10, padx=20, fill='x')

    outputfile_frame = ttk.Frame(settings, width=20)
    outputfile_frame.pack(pady=5, padx=10)

    mule_ini_frame = ttk.Frame(settings, width=20)
    mule_ini_frame.pack(pady=5, padx=10)

    middle_separator = ttk.Separator(settings, orient='horizontal')
    middle_separator.pack(pady=10, padx=20, fill='x')

    exclusions_frame = ttk.Frame(settings)
    exclusions_frame.pack(pady=5)

    bottom_separator = ttk.Separator(settings, orient='horizontal')
    bottom_separator.pack(pady=10, padx=20, fill='x')

    config_button = ttk.Button(settings, text='Save', command=save_settings)
    config_button.configure(width=30, style='primary.Outline.TButton')
    config_button.pack(pady=10, padx=10)

    # ------------- numbers frame layout -------------
    hotkey_frame = ttk.Frame(numbers_frame, width=20)
    hotkey_frame.pack(padx=5, side='left')

    vert_separator = ttk.Separator(numbers_frame, orient='vertical')
    vert_separator.pack(pady=10, padx=10, side='left', fill='y')

    auction_frame = ttk.Frame(numbers_frame)
    auction_frame.pack(padx=5, side='left')

    # ------------- hotkey frame layout -------------
    hotkey_label = ttk.Label(hotkey_frame, style='primary.TLabel', text='Hotkey Starting Location')
    hotkey_label.pack(padx=5, pady=5)

    location_frame = ttk.Frame(hotkey_frame, width=20)
    location_frame.pack(padx=5, pady=5)

    # ------------- location frame layout -------------
    page_label = ttk.Label(location_frame, style='primary.TLabel', text='Page')
    page_label.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

    page_entry = ttk.Entry(location_frame, style='primary.TEntry', justify='center',
                           width=5, textvariable=page)
    page_entry.grid(row=0, column=1, sticky='nsew', padx=5, pady=5)

    button_label = ttk.Label(location_frame, style='primary.TLabel', text='Button')
    button_label.grid(row=0, column=2, sticky='nsew', padx=5, pady=5)

    button_entry = ttk.Entry(location_frame, style='primary.TEntry', justify='center',
                             width=5, textvariable=button)
    button_entry.grid(row=0, column=3, sticky='nsew', padx=5, pady=5)

    # ------------- auction frame layout -------------
    auction_label = ttk.Label(auction_frame, style='primary.TLabel', text='Auctions to Count')
    auction_label.pack(padx=5, pady=5)

    num_auctions_frame = ttk.Frame(auction_frame, width=20)
    num_auctions_frame.pack(padx=5, pady=5)

    auction_entry = ttk.Entry(num_auctions_frame, style='primary.TEntry', justify='center',
                              width=5, textvariable=auctions)
    auction_entry.pack(padx=5, pady=5)

    # ------------- outputfile frame layout -------------
    outputfile_label = ttk.Label(outputfile_frame, style='primary.TLabel', text='Outputfile Path', width=15)
    outputfile_label.pack(padx=10, side='left')

    outputfile_entry = ttk.Entry(outputfile_frame, style='primary.TEntry', width=45, textvariable=outputfile_path)
    outputfile_entry.bind("<Button-1>", change_outputfile)
    outputfile_entry.pack(padx=10, side='left')

    # ------------- mule ini frame layout -------------
    mule_ini_label = ttk.Label(mule_ini_frame, style='primary.TLabel', text='Mule Ini Path', width=15)
    mule_ini_label.pack(padx=10, side='left')

    mule_ini_entry = ttk.Entry(mule_ini_frame, style='primary.TEntry', width=45, textvariable=mule_ini_path)
    mule_ini_entry.bind("<Button-1>", change_mule_ini)
    mule_ini_entry.pack(padx=10, side='left')

    # ------------- exclusions frame layout -------------
    label_button_frame = ttk.Frame(exclusions_frame)
    label_button_frame.pack(padx=10, side='left')

    list_frame = ttk.Frame(exclusions_frame, width=20)
    list_frame.pack(pady=5, padx=10, side='left')

    # ------------- label button frame layout -------------
    exclusions_label = ttk.Label(label_button_frame, style='primary.TLabel', text='Item Exclusions')
    exclusions_label.pack(pady=10)

    exclusions_button = ttk.Button(label_button_frame, text='Delete',
                                   style='primary.Outline.TButton', command=delete_exclusion)
    exclusions_button.pack()

    # ------------- list frame layout -------------
    exclusions_listbox = tk.Listbox(list_frame)
    list_scroll = ttk.Scrollbar(list_frame)

    list_scroll.configure(orient='vertical', command=exclusions_listbox.yview)
    list_scroll.pack(side='right', fill='y')

    exclusions_listbox.configure(width=40, font=listbox_font, yscrollcommand=list_scroll.set)
    populate_exclusions()
    exclusions_listbox.pack(fill='both')


# -------------------------------------------------
# ------------- code main entry point -------------
# -------------------------------------------------

# ignore SSL certificate warnings because eqtunnelauctions
# is getting on my last nerve
urllib3.disable_warnings()

# set size and location parameters based on OS version
if platform.release() == '10':
    button_font = ('Inter', 12)
    entry_font_small = ('Inter', 10)
    entry_font_large = ('Inter', 12)
    label_font_title = ('Inter', 24)
    label_font_large = ('Inter', 12)
    label_font_small = ('Inter', 10)
    sheet_font = ('Inter', 10, 'normal')
    listbox_font = ('Inter', 10)
    info_title_large = ('Inter', 16, 'bold')
    info_title_small = ('Inter', 12, 'bold')
    adjust_x_pos = 9
    adjust_y_pos = 51
else:
    button_font = ('Inter', 10)
    entry_font_small = ('Inter', 8)
    entry_font_large = ('Inter', 10)
    label_font_title = ('Inter', 20)
    label_font_large = ('Inter', 10)
    label_font_small = ('Inter', 8)
    sheet_font = ('Inter', 8, 'normal')
    listbox_font = ('Inter', 8)
    info_title_large = ('Inter', 14, 'bold')
    info_title_small = ('Inter', 10, 'bold')
    adjust_x_pos = 10
    adjust_y_pos = 63

# ----------- global variables -----------
inventory_path = ''
character_path = ''
hotkey_page = ''
hotkey_button = ''
auctions_count = ''
exclusions_list = []

# ----------------------------------------
# ------------- main window --------------
# ----------------------------------------


app = ttk.Window(themename='flatly')
app.geometry('520x620')
app.title('Auction Builder')
app.resizable(False, False)
app.iconbitmap('EverQuest.ico')

style = ttk.Style()
style.configure('primary.Outline.TButton', font=button_font, width=12)
style.configure('primary.TEntry', font=entry_font_small)
style.configure('title.TLabel', font=label_font_title)
style.configure('primary.TLabel', font=label_font_large)
style.configure('secondary.TLabel', font=label_font_small)

curr_item = ttk.StringVar()
tot_items = ttk.StringVar()

# ------------- main settings layout (label/sheet/buttons) -------------

# ------------- menu setup -------------
main_menu = ttk.Menu(app)
app.configure(menu=main_menu)

# ------------- file menu setup -------------
file_menu = ttk.Menu(main_menu)
file_menu.add_command(label='Settings', command=lambda: open_settings(True))
file_menu.add_separator()
file_menu.add_command(label='Exit', command=sys.exit)

# ------------- help menu setup -------------
help_menu = ttk.Menu(main_menu)
help_menu.add_command(label='Readme', command=open_readme)
help_menu.add_command(label='About', command=open_about)

main_menu.add_cascade(label='File', menu=file_menu)
main_menu.add_cascade(label='Help', menu=help_menu)

title = ttk.Label(app, text='Project Quarm Auction Builder', style='title.TLabel')
title.pack(pady=10)

sheet = tksheet.Sheet(app, font=sheet_font)
sheet.set_options(default_row_height=30).height_and_width(width=470, height=400)
sheet.enable_bindings("single_select", "row_select", "right_click_popup_menu",
                      "rc_delete_row", "arrowkeys", "rc_select", "rc_insert_row",
                      "copy", "cut", "paste", "delete", "undo", "edit_cell")
sheet.set_header_data(('Item', 'ID', 'Price', 'Exclude')).set_sheet_data([])
sheet.checkbox("D")
set_sheet_columns()
sheet.pack(pady=10)

bottom_frame = ttk.Frame(app)
bottom_frame.pack(pady=5)

# ------------- bottom_frame layout -------------
status_frame = ttk.Frame(bottom_frame)
status_frame.grid(row=0, column=0, padx=10, sticky='nsew')

vert_separator = ttk.Separator(bottom_frame, orient='vertical')
vert_separator.grid(row=0, column=1, padx=25, sticky='nsew')

button_frame = ttk.Frame(bottom_frame)
button_frame.grid(row=0, column=2, padx=10, sticky='nsew')

# ------------- status_frame layout -------------
current_label = ttk.Label(status_frame, style='primary.TLabel', text='Item')
current_label.pack(pady=5, padx=5, side='left')

current_entry = ttk.Entry(status_frame, font=entry_font_large, width=5, justify='center',
                          foreground='black', textvariable=curr_item, state='disabled')
current_entry.pack(pady=5, padx=5, side='left')

total_label = ttk.Label(status_frame, style='primary.TLabel', text='Of')
total_label.pack(pady=5, padx=5, side='left')

total_entry = ttk.Entry(status_frame, font=entry_font_large, width=5, justify='center',
                        foreground='black', textvariable=tot_items, state='disabled')
total_entry.pack(pady=5, padx=5, side='left')

# ------------- button_frame layout -------------
import_button = ttk.Button(button_frame, text='Import', command=import_items)
import_button.configure(width=15, style='primary.Outline.TButton')
import_button.pack(pady=5)

save_button = ttk.Button(button_frame, text='Save', command=lambda: build_file_list(sheet.get_sheet_data()))
save_button.configure(width=15, style='primary.Outline.TButton')
save_button.pack(pady=5)

# read in settings file and store in globals
read_settings()

# ------------- tkinter main loop -------------
app.mainloop()
