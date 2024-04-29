# -*- coding: utf-8 -*-
"""
Created on Thu May 18 07:46:39 2023
"""

import PySimpleGUI as psg
import easyocr
import string
import cv2
import pandas as pd


def main():
    custom_theme = {'BACKGROUND': '#01826B',
                    'TEXT': '#000000',
                    'INPUT': '#FFFFFF',
                    'TEXT_INPUT': '#000000',
                    'SCROLL': '#D4C2A1',
                    'BUTTON': ('#000000', '#B8A581'),
                    'PROGRESS': ('#01826B', '#D0D0D0'),
                    'BORDER': 1,
                    'SLIDER_DEPTH': 0,
                    'PROGRESS_DEPTH': 0
                }
    psg.theme_add_new('custom', custom_theme)
    psg.theme('custom')

    item_list = ['Quick Read', 'Slow Accurate Read']
    layout = [
        [psg.Text("Choose images to read: "), psg.FilesBrowse(
            key='file_choice',
            file_types=(('PNG', '*.png'), ('JPEG', '*.jpg'), ('TIFF', '*.tiff')))],
        [psg.Output(size=(110, 7))],
        [psg.Text('Select type of read to use: '), psg.Combo(values=item_list, size=(16,10), key='dd_item')],
        [psg.B("Run Script", bind_return_key=True), psg.Exit()]]
    
    window = psg.Window("Postal OCR", layout, grab_anywhere=True)
    nad = 'NAD-BH.xlsx'
    address_list = parse(nad)
    reader = easyocr.Reader(['en'])

    while True:
        event, values = window.Read()
        files = values['file_choice'].split(';')
        if event == psg.WIN_CLOSED or event == "Exit":
            break
        if event == "Run Script":
            if values['file_choice'] == '':
                print('WARNING: No file selected!')
                continue
            elif values['dd_item'] not in item_list:
                print('WARNING: No run choice selected!')
                continue
            else:
                curr_attempt = 1
                for file in files:
                    length = str(len(files))
                    print("Image Name: " + file)
                    print("Reading image (" + str(curr_attempt) + "/" + length + ")...")
                    user(address_list, reader, file, values['dd_item'])
                    curr_attempt += 1
                    print('-'*192)
    window.close()


# Function that parses NAD and returns a list of all addresses as strings
def parse(nad):
    database = pd.read_excel(nad, usecols='A:F')
    database = database.fillna(0)
    data_list = database.values.flatten().tolist()
    address_list = []

    for i in range(0, len(data_list), 6):
        rough_address = data_list[i:i + 6]
        clean_address = ''
        for part in rough_address:
            if part != 0:
                clean_address += ' ' + str(part).upper()
        address_list.append(clean_address.strip())
    return address_list


# Function that talks to user and forwards values to correct reader
def user(address_list, reader, response, scan_choice):
    try:
        image = open(response)
        if scan_choice == 'Quick Read':
            ocr_string, top_address = original_reader(response,
                                                      address_list,
                                                      reader)
            print('-' * 192)
            print('OCR Reader Results:\n' + ocr_string)
            print('Predicted possible addresses:')
            count = 1
            for address in top_address:
                print(str(count) + '. ' + address)
                count += 1
        elif scan_choice == 'Slow Accurate Read':
            ocr_string, top_address = new_read(response,
                                                  address_list,
                                                  reader)
            print('-' * 192)
            print('OCR Reader Results:\n' + ocr_string)
            print('Predicted possible addresses:')
            count = 1
            for address in top_address:
                print(str(count) + '. ' + address)
                count += 1
    except Exception as e:
        print("ERROR: GeneralError in code")


# Function that reads addressList and compares original ocr result using LCS
def original_reader(image, address_list, reader):
    #Reads and puts into string format
    ocr_string = ''
    orig_result = reader.readtext(image, detail=0, paragraph=True)
    for line in orig_result:
        ocr_string += line.upper()
    #Compares string to addresses using LCS
    print("Predicting Address...")
    top_address, max_lcs = compare(ocr_string, address_list)
    return ocr_string, top_address


# Function that reads addressList and compares new ocr result using LCS
def new_read(image, address_list, reader):
    #Creates string of allowable characters
    alph = str(string.ascii_lowercase) + str(string.ascii_uppercase)
    allowed_char = '0123456789' + alph + ',' + '-' + ' ' + '+' + '&'

    #Converts image to grey scale
    img = cv2.imread(image)
    grey_image = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    
    #Reads and puts into string format
    ocr_string = ''
    sim_result = reader.readtext(grey_image,
                                         allowlist=allowed_char,
                                         decoder='wordbeamsearch',
                                         detail=0,
                                         paragraph=True)
    for line in sim_result:
        ocr_string += line.upper()
    #Compares string to addresses using LCS
    print("Predicting Address...")
    top_address,max_lcs = compare(ocr_string,address_list)
    return ocr_string,top_address


# Function that compares ocr result against all addresses using LCS
def compare(ocr_string, address_list):
    max_lcs = 0
    top_address = []
    for address in address_list:
        seq_score = lc_seq(ocr_string, address)
        sub_score = lc_sub(ocr_string, address)
        if sub_score > 6: #Very likely match, high reward
            sub_score = sub_score * 4
        elif sub_score > 3: #Good match, good reward
            sub_score = sub_score * 2
        curr_score = seq_score + sub_score  
        if curr_score > max_lcs: #Set highest found score so far
            max_lcs = curr_score
            top_address.clear()
            top_address.append(address)
        elif curr_score == max_lcs: #Creates a 'tie' if equally likely 'high' scores found
            top_address.append(address)
    return top_address, max_lcs


# Function that compares a string and an address using 
# Longest Common Subsequence
def lc_seq(ocr_string, address):
    ocr_axis = len(ocr_string)
    addr_axis = len(address)
    score_grid = []
    for i in range(ocr_axis + 1): #Creates empty array with 1 extra row/column
        score_grid.append([None] * (addr_axis + 1))
    for i in range(ocr_axis + 1):
        for j in range(addr_axis + 1):
            if i == 0 or j == 0: #Fills extra row/column with zeroes
                score_grid[i][j] = 0
            elif ocr_string[i - 1] == address[j - 1]: #If characters match, add score  
                score_grid[i][j] = score_grid[i - 1][j - 1] + 1 #from previous diagonal

            else: #If no match, take highest score from neighbor (left or up)
                score_grid[i][j] = max(score_grid[i - 1][j],
                                      score_grid[i][j - 1])
    score = score_grid[i][j]
    return score


# Function that compares a string and an address using
# Longest Common Substring
def lc_sub(ocr_string, address):
    ocr_axis = len(ocr_string)
    addr_axis = len(address)
    score_grid = []
    score = 0
    for i in range(ocr_axis + 1): #Creates empty array with 1 extra row/column
        score_grid.append([None] * (addr_axis + 1))
    for i in range(ocr_axis + 1):
        for j in range(addr_axis + 1):
            if i == 0 or j == 0: #Fills extra row/column with zeroes
                score_grid[i][j] = 0
            elif ocr_string[i - 1] == address[j - 1]: #If characters match, add score
                score_grid[i][j] = score_grid[i - 1][j - 1] + 1 #from previous diagonal
                score = max(score, score_grid[i][j])
            else: #If no match, set score to 0 (Only want consecutive matches)
                score_grid[i][j] = 0
    return score


main()

# newline to end file
