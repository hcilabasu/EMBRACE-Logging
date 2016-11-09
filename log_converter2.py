# This script is a refactored version of log_converter.py and includes ITS logging.

import os
import glob
import xml.etree.ElementTree as ET
import xlsxwriter

from datetime import datetime

PATH_TO_LOG_FILES = "./Log Data"
EXCEL_FILE_NAME = "log_data2.xlsx"

# Labels used in Excel file
WORKSHEET_LABELS = [
                        "Actor",
                        "Action ID",
                        "Selection",
                        "Action",
                        "Input",
                        "All Input",
                        "Book Title",
                        "Chapter Number",
                        "Chapter Title",
                        "Page Language",
                        "Page Mode",
                        "Page Number",
                        "Page Complexity",
                        "Sentence Number",
                        "Sentence Order",
                        "Sentence Complexity",
                        "Sentence Text",
                        "Manipulation Sentence",
                        "Step Number",
                        "Idea Number",
                        "Question Number",
                        "Verification",
                        "Error Type",
                        "User Step",
                        "Chapter Status",
                        "Assessment Status",
                        "School Code",
                        "App Mode",
                        "Condition",
                        "Study Day",
                        "Participant Code",
                        "Experimenter Name",
                        "Study Language",
                        "Date",
                        "Time",
                        "Sentence Time"
                    ]

# Current row to write in Excel file
row = 0

# Actor
def write_actor(child, worksheet):
    col = WORKSHEET_LABELS.index("Actor")
    
    if child.tag == "User_Action":
        worksheet.write(row, col, "User")
    elif child.tag == "System_Action":
        worksheet.write(row, col, "System")

# Action ID
def write_action_ID(child, worksheet):
    col = WORKSHEET_LABELS.index("Action ID")
    worksheet.write(row, col, child.find("User_Action_ID").text)

# Selection
def write_selection(child, worksheet):
    col = WORKSHEET_LABELS.index("Selection")
    worksheet.write(row, col, child.find("Selection").text)

# Action
def write_action(child, worksheet):
    col = WORKSHEET_LABELS.index("Action")
    worksheet.write(row, col, child.find("Action").text)

# Input for library logging
def write_library_input(action, input, child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    if action == "Start Session":
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Type: {0}".format(button_type))
    elif action == "End Session":
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Type: {0}".format(button_type))
    elif action == "Show Books":
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Type: {0}".format(button_type))
    elif action == "Unlock Library Item":
        button_type = input.find("Button_Type").text
        
        if button_type == "Book":
            book_title = input.find("Book_Title").text
            book_status = input.find("Book_Status").text
            worksheet.write(row, col, "Button Type: {0}\nBook Title: {1}\nBook Status: {2}".format(button_type, book_title, book_status))
        elif button_type == "Chapter":
            chapter_title = input.find("Chapter_Title").text
            book_title = input.find("Book_Title").text
            chapter_status = input.find("Chapter_Status").text
            worksheet.write(row, col, "Button Type: {0}\nChapter Title: {1}\nBook Title: {2}\nChapter Status: {3}".format(button_type, chapter_title, book_title, chapter_status))
    elif action == "Load Book":
        button_type = input.find("Button_Type").text
        book_title = input.find("Book_Title").text
        worksheet.write(row, col, "Button Type: {0}\nBook Title: {1}".format(button_type, book_title))
    elif action == "Load Chapter":
        button_type = input.find("Button_Type").text
        chapter_title = input.find("Chapter_Title").text
        book_title = input.find("Book_Title").text
        worksheet.write(row, col, "Button Type: {0}\nChapter Title: {1}\nBook Title:{2}".format(button_type, chapter_title, book_title))

# Input for manipulation logging
def write_manipulation_input(action, input, child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    if action == "Move Object":
        object = input.find("Object").text
        destination = input.find("Destination").text
        destination_type = input.find("Destination_Type").text
        start_position = input.find("Start_Position").text
        end_position = input.find("End_Position").text
        worksheet.write(row, col, "Object: {0}\nDestination: {1}\nDestination Type: {2}\nStart Position: {3}\nEnd Position: {4}".format(object, destination, destination_type, start_position, end_position))
    elif action == "Group Objects" or action == "Ungroup Objects" or action == "Ungroup and Stay Objects":
        object_1 = input.find("Object_1").text
        object_2 = input.find("Object_2").text
        hotspot = input.find("Hotspot").text
        worksheet.write(row, col, "Object 1: {0}\nObject 2: {1}\nHotspot: {2}".format(object_1, object_2, hotspot))
    elif action == "Display Menu Items":
        menu_items = input.find("Menu_Items")
        menu_items_list = []
        
        menu_item_0 = menu_items.find("Menu_Item_0")
        interaction_0 = menu_item_0.find("Interaction_0")
        object_1 = interaction_0.find("Object_1").text
        object_2 = interaction_0.find("Object_2").text
        hotspot = interaction_0.find("Hotspot").text
        interaction_type = interaction_0.find("Interaction_Type").text
        menu_items_list.append("Menu Item 0: Object 1: {0} Object 2: {1} Hotspot: {2} Interaction Type: {3}".format(object_1, object_2, hotspot, interaction_type))
        
        menu_item_1 = menu_items.find("Menu_Item_1")
        interaction_0 = menu_item_1.find("Interaction_0")
        object_1 = interaction_0.find("Object_1").text
        object_2 = interaction_0.find("Object_2").text
        hotspot = interaction_0.find("Hotspot").text
        interaction_type = interaction_0.find("Interaction_Type").text
        menu_items_list.append("Menu Item 1: Object 1: {0} Object 2: {1} Hotspot: {2} Interaction Type: {3}".format(object_1, object_2, hotspot, interaction_type))
        
        menu_item_2 = menu_items.find("Menu_Item_2")
        
        if menu_item_2 != None:
            interaction_0 = menu_item_2.find("Interaction_0")
            object_1 = interaction_0.find("Object_1").text
            object_2 = interaction_0.find("Object_2").text
            hotspot = interaction_0.find("Hotspot").text
            interaction_type = interaction_0.find("Interaction_Type").text
            menu_items_list.append("Menu Item 2: Object 1: {0} Object 2: {1} Hotspot: {2} Interaction Type: {3}".format(object_1, object_2, hotspot, interaction_type))
        
        worksheet.write(row, col, "\n".join(menu_items_list))
    elif action == "Select Menu Item":
        menu_item = list(input)[0]
        
        if menu_item.tag == "Menu_Item":
            worksheet.write(row, col, "Menu Item: {0}".format(menu_item.text)) # NULL
        else:
            menu_item_number = 0

            if menu_item.tag == "Menu_Item_0":
                menu_item_number = 0
            elif menu_item.tag == "Menu_Item_1":
                menu_item_number = 1
            elif menu_item.tag == "Menu_Item_2":
                menu_item_number = 2

            interaction_0 = menu_item.find("Interaction_0")
            object_1 = interaction_0.find("Object_1").text
            object_2 = interaction_0.find("Object_2").text
            hotspot = interaction_0.find("Hotspot").text
            interaction_type = interaction_0.find("Interaction_Type").text
            worksheet.write(row, col, "Menu Item {0}: Object 1: {1} Object 2: {2} Hotspot: {3} Interaction Type: {4}".format(menu_item_number, object_1, object_2, hotspot, interaction_type))
    elif action == "Verify Action":
        verification = input.find("Verification").text
        worksheet.write(row, col, "Verification: {0}".format(verification))
    elif action == "Maximum Attempts Reached":
        worksheet.write(row, col, "NULL") # No input is logged
    elif action == "Reset Object":
        object = input.find("Object").text
        start_position = input.find("Start_Position").text
        end_position = input.find("End_Position").text
        worksheet.write(row, col, "Object: {0}\nStart Position: {1}\nEnd Position: {2}".format(object, start_position, end_position))
    elif action == "Appear Object" or action == "Disappear Object":
        object = input.find("Object").text
        worksheet.write(row, col, "Object: {0}".format(object))
    elif action == "Swap Image":
        object = input.find("Object").text
        alternative_image = input.find("Alternative_Image").text
        worksheet.write(row, col, "Object: {0}\nAlternative Image: {1}".format(object, alternative_image))
    elif action == "Animate Object":
        object = input.find("Object").text
        animate_action = input.find("Animate_Action").text
        worksheet.write(row, col, "Object: {0}\nAnimate Action: {1}".format(object, animate_action))
    elif action == "Tap Object":
        object = input.find("Object").text
        worksheet.write(row, col, "Object: {0}".format(object))
    elif action == "Tap Word":
        word = input.find("Word").text
        worksheet.write(row, col, "Word: {0}".format(word))
    elif action == "Error Feedback Noise" or action == "Play Error Noise" or action == "Play Sound" or action == "Play Word" or action == "Play Word with Definition" or action == "Post-Sentence Script Audio" or action == "Pre-Sentence Script Audio":
        audio_name = input.find("Audio_Name").text
        audio_language = input.find("Audio_Language").text
        worksheet.write(row, col, "Audio Name: {0}\nAudio Language: {1}".format(audio_name, audio_language))

# Input for manipulation navigation logging
def write_manipulation_navigation_input(action, input, child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    if action == "Press Next":
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Type: {0}".format(button_type))
    elif action == "Skip Content":
        gesture_type = input.find("Gesture_Type").text
        worksheet.write(row, col, "Gesture Type: {0}".format(gesture_type))
    elif action == "Load Step":
        step_number = input.find("Step_Number").text
        step_type = input.find("Step_Type").text
        worksheet.write(row, col, "Step Number: {0}\nStep Type: {1}".format(step_number, step_type))
    elif action == "Load Sentence":
        sentence_number = input.find("Sentence_Number").text
        sentence_complexity = input.find("Sentence_Complexity").text
        sentence_text = input.find("Sentence_Text").text
        manipulation_sentence = input.find("Manipulation_Sentence").text
        worksheet.write(row, col, "Sentence Number: {0}\nSentence Complexity: {1}\nSentence Text: {2}\nManipulation Sentence: {3}".format(sentence_number, sentence_complexity, sentence_text, manipulation_sentence))
    elif action == "Load Page":
        page_language = input.find("Page_Language").text
        page_mode = input.find("Page_Mode").text
        page_number = input.find("Page_Number").text
        page_complexity = input.find("Page_Complexity").text
        worksheet.write(row, col, "Page Language: {0}\nPage Mode: {1}\nPage Number: {2}\nPage Complexity: {3}".format(page_language, page_mode, page_number, page_complexity))
    elif action == "Return to Library":
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Type: {0}".format(button_type))
    elif action == "Completed Manipulation Activity":
        worksheet.write(row, col, input.text) # NULL

# Input for assessment logging
def write_assessment_input(action, input, child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    if action == "Display Assessment Question":
        question_text = input.find("Question_Text").text
        answer_options = input.find("Answer_Options").text
        worksheet.write(row, col, "Question Text: {0}\nAnswer Options: {1}".format(question_text, answer_options))
    elif action == "Select Assessment Answer":
        selected_answer = input.find("Selected_Answer").text
        worksheet.write(row, col, "Selected Answer: {0}".format(selected_answer))
    elif action == "Verify Assessment Answer":
        verification = input.find("Verification").text
        worksheet.write(row, col, "Verification: {0}".format(verification))
    elif action == "Play Answer Audio" or action == "Play Question Audio":
        audio_name = input.find("Audio_Name").text
        audio_language = input.find("Audio_Language").text
        worksheet.write(row, col, "Audio Name: {0}\nAudio Language: {1}".format(audio_name, audio_language))
    elif action == "Skip Content":
        gesture_type = input.find("Gesture_Type").text
        worksheet.write(row, col, "Gesture Type: {0}".format(gesture_type))

# Input for assessment navigation logging
def write_assessment_navigation_input(action, input, child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    if action == "Tap Assessment Audio":
        button_name = input.find("Button_Name").text
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Name: {0}\nButton Type: {1}".format(button_name, button_type))
    elif action == "Press Next":
        button_type = input.find("Button_Type").text
        worksheet.write(row, col, "Button Type: {0}".format(button_type))
    elif action == "Load Assessment Step":
        assessment_step_number = input.find("Assessment_Step_Number").text
        worksheet.write(row, col, "Assessment Step Number: {0}".format(assessment_step_number))
    elif action == "Completed Assessment Activity":
        worksheet.write(row, col, input.text) # NULL

# Input for ITS logging
def write_ITS_input(action, input, child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    if action == "Updated Skill Value":
        skill_name = input.find("Skill_Name").text
        previous_skill_value = input.find("Previous_Skill_Value").text
        new_skill_value = input.find("New_Skill_Value").text
        worksheet.write(row, col, "Skill Name: {0}\nPrevious Skill Value: {1}\nNew Skill Value: {2}".format(skill_name, previous_skill_value, new_skill_value))
    elif action == "Adapted Vocabulary Introduction":
        extra_vocabulary = input.find("Extra_Vocabulary").text
        worksheet.write(row, col, "Extra Vocabulary: {0}".format(extra_vocabulary))
    elif action == "Adapted Chapter Syntax":
        previous_complexity = input.find("Previous_Complexity").text
        new_complexity = input.find("New_Complexity").text
        worksheet.write(row, col, "Previous Complexity: {0}\nNew Complexity: {1}".format(previous_complexity, new_complexity))
    elif action == "Provide Vocabulary Error Feedback":
        highlighted_items = input.find("Highlighted_Items").text
        worksheet.write(row, col, "Highlighted Items: {0}".format(highlighted_items))
    elif action == "Provide Syntax Error Feedback":
        simpler_sentence = input.find("Simpler_Sentence").text
        worksheet.write(row, col, "Simpler Sentence: {0}".format(simpler_sentence))
    elif action == "Provide Usability Error Fedback":
        animated_items = input.find("Animatd_Items").text
        worksheet.write(row, col, "Animated Items: {0}".format(animated_items))

# Input
def write_input(child, worksheet):
    col = WORKSHEET_LABELS.index("Input")
    
    action = child.find("Action").text
    input = child.find("Input")
    context = child.find("Context")
    
    # Manipulation Activity
    if context.find("Manipulation_Context") != None:
        write_manipulation_input(action, input, child, worksheet)
        write_manipulation_navigation_input(action, input, child, worksheet)
        write_ITS_input(action, input, child, worksheet)
    # Assessment Activity
    elif context.find("Assessment_Context") != None:
        write_assessment_input(action, input, child, worksheet)
        write_assessment_navigation_input(action, input, child, worksheet)
    # Library
    else:
        write_library_input(action, input, child, worksheet)

    col = WORKSHEET_LABELS.index("All Input")
    input_list = []
    
    for input in list(child.find("Input")):
        if input.text != None:
            input_list.append(input.text)

    worksheet.write(row, col, ", ".join(input_list))

# Manipulation Context
def write_manipulation_context(context, worksheet):
    manipulation_context = context.find("Manipulation_Context")

    if manipulation_context != None:
        col = WORKSHEET_LABELS.index("Book Title")
        worksheet.write(row, col, manipulation_context.find("Book_Title").text)

        col = WORKSHEET_LABELS.index("Chapter Number")
        worksheet.write(row, col, manipulation_context.find("Chapter_Number").text)

        col = WORKSHEET_LABELS.index("Chapter Title")
        worksheet.write(row, col, manipulation_context.find("Chapter_Title").text)

        col = WORKSHEET_LABELS.index("Page Language")
        worksheet.write(row, col, manipulation_context.find("Page_Language").text)

        col = WORKSHEET_LABELS.index("Page Mode")
        worksheet.write(row, col, manipulation_context.find("Page_Mode").text)

        col = WORKSHEET_LABELS.index("Page Number")
        worksheet.write(row, col, manipulation_context.find("Page_Number").text)

        col = WORKSHEET_LABELS.index("Page Complexity")
        worksheet.write(row, col, manipulation_context.find("Page_Complexity").text)

        col = WORKSHEET_LABELS.index("Sentence Number")
        worksheet.write(row, col, manipulation_context.find("Sentence_Number").text)

        col = WORKSHEET_LABELS.index("Sentence Complexity")
        worksheet.write(row, col, manipulation_context.find("Sentence_Complexity").text)

        col = WORKSHEET_LABELS.index("Sentence Text")
        worksheet.write(row, col, manipulation_context.find("Sentence_Text").text)

        col = WORKSHEET_LABELS.index("Manipulation Sentence")
        worksheet.write(row, col, manipulation_context.find("Manipulation_Sentence").text)

        col = WORKSHEET_LABELS.index("Step Number")
        worksheet.write(row, col, manipulation_context.find("Step_Number").text)

        col = WORKSHEET_LABELS.index("Idea Number")
        worksheet.write(row, col, manipulation_context.find("Idea_Number").text)

# Assessment Context
def write_assessment_context(context, worksheet):
    assessment_context = context.find("Assessment_Context")

    if assessment_context != None:
        col = WORKSHEET_LABELS.index("Book Title")
        worksheet.write(row, col, assessment_context.find("Book_Title").text)
        
        col = WORKSHEET_LABELS.index("Chapter Title")
        worksheet.write(row, col, assessment_context.find("Chapter_Title").text)
        
        col = WORKSHEET_LABELS.index("Question Number")
        worksheet.write(row, col, assessment_context.find("Assessment_Step_Number").text)

# Study Context
def write_study_context(context, worksheet):
    study_context = context.find("Study_Context")
    
    if study_context != None:
        col = WORKSHEET_LABELS.index("School Code")
        worksheet.write(row, col, study_context.find("School_Code").text)
        
        col = WORKSHEET_LABELS.index("App Mode")
        worksheet.write(row, col, study_context.find("App_Mode").text)
        
        col = WORKSHEET_LABELS.index("Condition")
        worksheet.write(row, col, study_context.find("Condition").text)
        
        col = WORKSHEET_LABELS.index("Study Day")
        worksheet.write(row, col, study_context.find("Study_Day").text)
        
        col = WORKSHEET_LABELS.index("Participant Code")
        worksheet.write(row, col, study_context.find("Participant_Code").text)
        
        col = WORKSHEET_LABELS.index("Experimenter Name")
        worksheet.write(row, col, study_context.find("Experimenter_Name").text)
        
        col = WORKSHEET_LABELS.index("Study Language")
        worksheet.write(row, col, study_context.find("Language").text)
        
        timestamp = study_context.find("Timestamp")
        
        if timestamp != None:
            col = WORKSHEET_LABELS.index("Date")
            worksheet.write(row, col, timestamp.find("Date").text)
            
            col = WORKSHEET_LABELS.index("Time")
            worksheet.write(row, col, timestamp.find("Time").text)

# Context
def write_context(child, worksheet):
    context = child.find("Context")
    
    write_manipulation_context(context, worksheet)
    write_assessment_context(context, worksheet)
    write_study_context(context, worksheet)

# Reads log files and writes to Excel file
def read_log_file(worksheet, log_file):
    global row
    
    print "*** file: " + log_file

    tree = ET.parse(log_file)
    root = tree.getroot()

    for child in root:
        write_actor(child, worksheet)
        write_action_ID(child, worksheet)
        write_selection(child, worksheet)
        write_action(child, worksheet)
        write_input(child, worksheet)
        write_context(child, worksheet)

        row += 1

# Checks if specified file is a log file
def is_log_file(file_name):
    if file_name == ".DS_Store":
        return False
    
    if "progress.xml" in file_name:
        return False
    
    if ".xml" not in file_name:
        return False
    
    return True

# Sets up Excel file
def create_workbook():
    # Create Excel file
    workbook = xlsxwriter.Workbook(EXCEL_FILE_NAME)
    worksheet = workbook.add_worksheet()
    
    # Add a bold format
    bold = workbook.add_format({"bold": True})
    
    for index, label in enumerate(WORKSHEET_LABELS):
        worksheet.write(row, index, label, bold)
    
    return workbook, worksheet

def main():
    global row
    
    workbook, worksheet = create_workbook()
    
    row += 1
    
    for root, dirs, files in os.walk(PATH_TO_LOG_FILES):
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            
            for file in os.listdir(dir_path):
                file_path = os.path.join(dir_path, file)
                
                if is_log_file(file_path):
                    read_log_file(worksheet, file_path)

    workbook.close()

if __name__ == "__main__":
    main()
