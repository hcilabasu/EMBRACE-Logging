import os
import glob
import xml.etree.ElementTree as ET
import xlsxwriter

from datetime import datetime

FMT = '%H:%M:%S'

# Create Excel file
workbook = xlsxwriter.Workbook('log_data.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format
bold = workbook.add_format({'bold': True})

# Write labels
worksheet.write('A3', 'Actor', bold) # User or Computer
actorIndex = 0
worksheet.write('B3', 'Action ID', bold) # Number associated with a specific user or computer action
actionIDIndex = 1
worksheet.write('C3', 'Selection', bold) # The type of object the user interacted with (button, menu, image...)
selectionIndex = 2
worksheet.write('D3', 'Action', bold)
actionIndex = 3
worksheet.write('E3', 'Input', bold)
inputIndex = 4
worksheet.write('F3', 'All Input', bold)
allInputIndex = 5

# Manipulation Context Data
worksheet.write('G3', 'Book Title', bold)
BookTitleIndex = 6
worksheet.write('H3', 'Chapter Number', bold)
ChapterNumberIndex = 7
worksheet.write('I3', 'Chapter Title', bold)
ChapterTitleIndex = 8
worksheet.write('J3', 'Page Language', bold)
PageLanguageIndex = 9
worksheet.write('K3', 'Page Mode', bold)
PageModeIndex = 10
worksheet.write('L3', 'Page Number', bold)
PageNumberIndex = 11
worksheet.write('M3', 'Sentence Number', bold)
SentenceNumberIndex = 12
worksheet.write('N3', 'Sentence Order', bold)
SentenceOrderIndex = 13
worksheet.write('O3', 'Sentence Text', bold)
SentenceTextIndex = 14
worksheet.write('P3', 'Manipulation Sentence', bold)
ManipulationSentenceIndex = 15
worksheet.write('Q3', 'Step Number', bold)
StepNumberIndex = 16
worksheet.write('R3', 'Idea Number', bold)
IdeaNumberIndex = 17
worksheet.write('S3', 'Question Number', bold)
QuestionNumberIndex = 18

worksheet.write('T3', 'Verification', bold)
VerificationIndex = 19
worksheet.write('U3', 'Error Type', bold)
ErrorTypeIndex = 20
worksheet.write('V3', 'User Step', bold)
UserStepIndex = 21
worksheet.write('W3', 'Chapter Status', bold)
ChapterStatusIndex = 22
worksheet.write('X3', 'Assessment Status', bold)
AssessmentStatusIndex = 23

# Study Context Data
worksheet.write('Y3', 'School Code', bold)
SchoolCodeIndex = 24
worksheet.write('Z3', 'Condition', bold)
ConditionIndex = 25
worksheet.write('AA3', 'Study Day', bold)
StudyDayIndex = 26
worksheet.write('AB3', 'Participant Code', bold)
ParticipantCodeIndex = 27
worksheet.write('AC3', 'Experimenter Name', bold)
ExperimenterNameIndex = 28
worksheet.write('AD3', 'Study Language', bold)
StudyLanguageIndex = 29
worksheet.write('AE3', 'Date', bold)
DateIndex = 30
worksheet.write('AF3', 'Time', bold)
TimeIndex = 31
worksheet.write('AG3', 'Sentence Time', bold)
SentenceTimeIndex = 32

# Start writing data below labels
row = 3
col = 0

# Go through every .xml log file in each directory
for root, dirs, files in os.walk("./All Data"):
	for dir in dirs:
		print "*** dir: " + dir
		print glob.glob(os.path.join(dir, "*.xml"))
        
		dirPath = os.path.join(root, dir)

		prevParticipantID = ""
		sentenceOrder = -1; # increments across all stories for a single participant
		
		for file in glob.glob(os.path.join(dirPath, "*.xml")):
			print "*** file: " + file
            
			tree = ET.parse(file)
			root = tree.getroot()
			
			prevAction = "" # previous action (used for writing Error Type column)
			userStep = -1 # user step number
			prevSentence = -1 # previous sentence number
			ChapterBeginRow = row # row where new chapter starts
			AssessmentBeginRow = row # row where new assesment starts
			ChapterEndRow = -1 # row where chapter ends (if completed)
			AssessmentEndRow = -1 # row where the assessment ends (if completed)
			SentenceBeginRow = row # row where new sentence starts
			sentenceStartTime = ""
			sentenceElapsedTime = ""
				
			for child in root:
				displayMenuCount = 0 # number of "Display Menu" actions read
				verificationCol = 0;
				correctVerification = False # whether verification is correct (used for incrementing user steps)
				
				# Write Actor column
				if child.tag == 'User_Action' or child.tag == 'userAction':
					worksheet.write(row, actorIndex, 'User')
                
				elif child.tag == 'System_Action':
					worksheet.write(row, actorIndex, 'System')

				# Write Action ID column
				actionID = child.find('User_Action_ID').text if child.find('User_Action_ID').text is not None else ""
				worksheet.write(row, actionIDIndex, actionID)
				
				# Write Selection column
				selection = child.find('Selection').text if child.find('Selection').text is not None else ""
				worksheet.write(row, selectionIndex, selection)
				
				# Write Action column
				action = child.find('Action').text if child.find('Action').text is not None else ""
				worksheet.write(row, actionIndex, action)
				inputText = ""
				
				# Read a "Display Menu Items" action
				if action == "Display Menu Items":
					displayMenuCount += 1
					input = child.find('Input')
					input = input.find('Menu_Items')
					menuItem0 = input.find('Menu_Item_0')
					interaction0 = menuItem0.find('Interaction_0')
					object1 = interaction0.find('Object_1').text
					object2 = interaction0.find('Object_2').text
					hotspot = interaction0.find('Hotspot').text
					interactionType = interaction0.find('Interaction_Type').text
					menu1 = object1 + " " + interactionType + " with " + object2 + " at hotspot: " + hotspot
					
					menuItem1 = input.find('Menu_Item_1')
					interaction1 = menuItem1.find('Interaction_0')
					object1 = interaction1.find('Object_1').text
					object2 = interaction1.find('Object_2').text
					hotspot = interaction1.find('Hotspot').text
					interactionType = interaction1.find('Interaction_Type').text
					menu2 = object1 + " " + interactionType + " with " + object2 + " at hotspot: " + hotspot
					inputText = "Menu Item 1: " + menu1 + "\nMenu Item 2: " + menu2
                
				elif action == "Move Object":
					input = child.find('Input')
					object = input.find('Object').text
					destination = input.find('Destination').text
					destinationType = input.find('Destination_Type').text
					startPos = input.find('Start_Position').text
					startCord = startPos.split(", ", 2)
					startPosX = ("%.2f" % round(float(startCord[0]),2)) 
					startPosY = ("%.2f" % round(float(startCord[1]),2)) 
					startPos = "(" + startPosX + ", " + startPosY + ")"
					endPos =  input.find('End_Position').text
					endCord = endPos.split(", ", 2)
					endPosX = ("%.2f" % round(float(endCord[0]),2)) 
					endPosY = ("%.2f" % round(float(endCord[1]),2)) 
					endPos = "(" + endPosX + ", " + endPosY + ")"
					inputText = "Move Object: " + object + " to " + destinationType + ": " + destination + ". From: " + startPos + " -> " + endPos

				elif action == "Group Objects":
					input = child.find('Input')
					object1 = input.find('Object_1').text
					object2 = input.find('Object_2').text
					hotspot = input.find('Hotspot').text
					inputText = "Group " + object1 + " with " + object2 + " at hotspot: " + hotspot
                        
				elif action == "Animate Object":	
					input = child.find('Input')
					object = input.find('Object').text
					animateAction =  input.find('Animate_Action').text
					inputText = "Animate " + object + " with " + animateAction
                
				elif action == "Appear Object":
					input = child.find('Input')
					object = input.find('Object').text
					inputText = "Appear " + object
                
				elif action == "Disappear Object":
					input = child.find('Input')
					object = input.find('Object').text
					inputText = "Disappear " + object
                
				elif (action == "Completed Assessment Activity") | (action == "Completed Manipulation Activity"):
					if (action == "Completed Assessment Activity"):
						AssessmentEndRow = row
                    
					elif (action == "Completed Manipulation Activity"):
						ChapterEndRow = row
                    
					inputText = "NULL"
                        
				elif action == "Display Assessment Question":
					input = child.find('Input')
					questionText = input.find('Question_Text').text
					answerOptions = input.find('Answer_Options').text
					inputText = "Question: " + questionText + "\nOptions: " + answerOptions
                
				elif (action == "End Session") | (action == "Load Book") | (action == "Load Chapter") | (action == "Press Next") | (action == "Show Books") | (action == "Return to Library") | (action == "Start Session") | (action == "Unlock Library Item"):
					input = child.find('Input')
					buttonType = input.find('Button_Type').text
					inputText = "Button Type: " + buttonType
					bookTitle = input.find('Book_Title')
                    
					if bookTitle is not None:
						inputText += "\nBook Title: " + bookTitle.text
                    
					chapterTitle = input.find('Chapter_Title')
                    
					if chapterTitle is not None:
						inputText += "\nChapter Title: " + chapterTitle.text
                    
					bookStatus = input.find('Book_Status')
                    
					if bookStatus is not None:
						inputText += "\nBook Status: " + bookStatus.text

				elif action ==  "Tap Assessment Audio":
					input = child.find('Input')
					buttonType = input.find('Button_Type').text
					butonName = input.find('Button_Name').text
					inputText = "Button Type: " + buttonType + "\nButton Name: " + butonName
                
				elif (action == "Play Answer Audio") | (action == "Play Error Noise") | (action == "Play Question Audio") | (action == "Play Word") | (action == "Play Word Definition") | (action == "Pre-Sentence Script Audio") | (action == "Post-Sentence Script Audio"):
					input = child.find('Input')
					audioName = input.find('Audio_Name').text
					audioLanguage = input.find('Audio_Language').text
					inputText = "Audio File: " + audioName + "\nAudio Language: " + audioLanguage

				elif action == "Load Assessment Step":
					input = child.find('Input')
					assessmentStepNumber = input.find('Assessment_Step_Number').text
					inputText = "Question #: " + assessmentStepNumber

				elif action == "Select Assessment Answer":
					input = child.find('Input')
					selectedAnswer = input.find('Selected_Answer').text
					inputText = "Selected Answer: " + selectedAnswer

				elif action == "Tap Word":
					input = child.find('Input')
					word = input.find('Word').text
					inputText = "Word: " + word

				elif action == "Select Menu Item":
					input = child.find('Input')
					menuItem0 = input.find('Menu_Item_0')
					menuItem = menuItem0
                    
					if menuItem0 is None:
						menuItem1 = input.find('Menu_Item_1')
						menuItem = menuItem1
                    
					if menuItem is not None:
						interaction0 = menuItem.find('Interaction_0')
						object1 = interaction0.find('Object_1').text
						object2 = interaction0.find('Object_2').text
						hotspot = interaction0.find('Hotspot').text
						interactionType = interaction0.find('Interaction_Type').text
						inputText = object1 + " " + interactionType + " with " + object2 + " at hotspot: " + hotspot

				elif (action == "Ungroup Objects") | (action == "Ungroup and Stay Objects"):
					input = child.find('Input')
					object1 = input.find('Object_1').text
					object2 = input.find('Object_2').text
					hotspot = input.find('Hotspot').text
					inputText = "Ungroup " + object1 + " from " + object2 + " at hotspot " + hotspot
                
				elif (action == "Verify Action") | (action == "Verify Assessment Answer"):
					input = child.find('Input')
					verification = input.find('Verification').text
                    
					# Write the Verification (Correct or Incorrect) to the previous row. Note: "Display Menu" actions repeat verification unnecessarily.
							# Record that this verification was correct
					if verification == "Correct":
						correctVerification = True
							
						worksheet.write(row - 1, VerificationIndex, verification)
						worksheet.write(row - 1, ErrorTypeIndex, prevAction) # write the previous action in the Error Type column
					
					inputText = "Verification: " + verification

				elif action == "Reset Object":
					input = child.find('Input')
					object = input.find('Object').text
					startPos = input.find('Start_Position').text
					startCord = startPos.split(", ", 2)
					startPosX = ("%.2f" % round(float(startCord[0]),2)) 
					startPosY = ("%.2f" % round(float(startCord[1]),2)) 
					startPos = "(" + startPosX + ", " + startPosY + ")"
					endPos = input.find('End_Position').text
					inputText = "Reset " + object + " to " + startPos

				elif action == "Skip Content":
					input = child.find('Input')
					gestureType = input.find('Gesture_Type').text
					inputText = "Gesture: " + gestureType

				elif action == "Swap Image":
					input = child.find('Input')
					object = input.find('Object').text
					altImage = input.find('Alternative_Image').text
					inputText = "Swap " + object + " with " + altImage

				elif action == "Tap Object":
					input = child.find('Input')
					object = input.find('Object').text
                    
					if object is None:
						object = ""
                    
					inputText = "Object: " + object

				elif action == "Load Page":
					input = child.find('Input')
					pageLang = input.find('Page_Language').text
					pageMode = input.find('Page_Mode').text
					pageNum = input.find('Page_Number').text
					inputText = "Page Language: " + pageLang + "\nPage Mode: " + pageMode + "\nPage Number: " + pageNum

				elif action == "Load Sentence":
					input = child.find('Input')
					sentenceNum = input.find('Sentence_Number').text
					sentenceText = input.find('Sentence_Text').text
					manipulationSentence = input.find('Manipulation_Sentence').text
					inputText = "Manipulation Sentence: " + manipulationSentence + "\nSentence Number: " + sentenceNum + "\nSentence Text: " + sentenceText

					if sentenceElapsedTime != "":
						for i in range(SentenceBeginRow, row):
							worksheet.write(i, SentenceTimeIndex, str(sentenceElapsedTime))

					SentenceBeginRow = row
					sentenceStartTime = ""
					sentenceElapsedTime = ""

				elif action == "Load Step":
					input = child.find('Input')
					stepNum = input.find('Step_Number').text
					stepType = input.find('Step_Type').text
                    
					if stepType is None:
						stepType = "NULL"
                    
					inputText = "Step Number: " + stepNum + "\nStep Type: " + stepType

				allInput = ""
                
				for input in list(child.find('Input')):
					if input.text is not None:
						allInput+= input.text + ", "
				
				# Record this action as the previous action
				prevAction = action
				
				worksheet.write(row, allInputIndex, allInput)
				worksheet.write(row, inputIndex, inputText)
								
				# Write Context section
				context = child.find('Context')
				
				# Write the Assessment Context section
				assessmentContext = context.find('Assessment_Context')
				
				if assessmentContext is not None:
					# Write Book Title column
					bookTitle = assessmentContext.find('Book_Title').text
					worksheet.write(row, BookTitleIndex, bookTitle)
					
					# TODO: Set vars to get last manipulationContext to populate missing context data
					
					# Write Chapter Number column
					chapterTitle = assessmentContext.find('Chapter_Title').text
					worksheet.write(row, ChapterTitleIndex, chapterTitle)
					
					# Write Question Number column
					questionNumber = assessmentContext.find('Assessment_Step_Number').text
					worksheet.write(row, QuestionNumberIndex, questionNumber)
				
				# Write the Manipulation Context section
				manipulationContext = context.find('Manipulation_Context')
				
				if manipulationContext is not None:
					# Write Book Title column
					bookTitle = manipulationContext.find('Book_Title').text
					worksheet.write(row, BookTitleIndex, bookTitle)

					# Write Chapter Number column
					chapterNumber = manipulationContext.find('Chapter_Number').text
					worksheet.write(row, ChapterNumberIndex, chapterNumber)

					# Write Chapter Title column
					chapterTitle = manipulationContext.find('Chapter_Title').text
					worksheet.write(row, ChapterTitleIndex, chapterTitle)
					
					# Write Page Language column
					pageLanguage = manipulationContext.find('Page_Language').text
					worksheet.write(row, PageLanguageIndex, pageLanguage)
					
					# Write Page Mode column
					pageMode = manipulationContext.find('Page_Mode').text
					worksheet.write(row, PageModeIndex, pageMode)
					
					# Write Page Number column
					pageNumber = manipulationContext.find('Page_Number').text
					worksheet.write(row, PageNumberIndex, pageNumber)
					
					# Write Sentence Number column
					sentenceNumber = manipulationContext.find('Sentence_Number').text
					worksheet.write(row, SentenceNumberIndex, sentenceNumber)
					
					# Write Sentence Text column
					sentenceText = manipulationContext.find('Sentence_Text').text
					worksheet.write(row, SentenceTextIndex, sentenceText)
					
					# Write Manipulation Sentence column
					manipulationSentence = manipulationContext.find('Manipulation_Sentence').text
					worksheet.write(row, ManipulationSentenceIndex, manipulationSentence)
					
					# Write Step Number column
					stepNumber = manipulationContext.find('Step_Number').text
					worksheet.write(row, StepNumberIndex, stepNumber)
					
					# Write Idea Number column
					ideaNumber = manipulationContext.find('Idea_Number').text
					worksheet.write(row, IdeaNumberIndex, ideaNumber)
					
					# Write User Step column
					if sentenceNumber == "NULL":
						worksheet.write(row, SentenceNumberIndex, "NULL")
						worksheet.write(row, SentenceOrderIndex, "NULL")

					else:
						# Increase user step number if verification is correct
						if correctVerification:
							correctVerification = False
							userStep += 1
                        
						# Decrease user step number if action is "Display Menu" because we want user moving to an object/hotspot and tapping a menu to count as one user step
						if displayMenuCount == 1:
							userStep -= 1
							# Rewrite the previous row's user step with this correction
							worksheet.write(row - 1, UserStepIndex, userStep)
                        
						elif displayMenuCount == 2:
							displayMenuCount = 0
                        
						# Reset user step to 1 if we have changed sentences
						if int(sentenceNumber) != prevSentence:
							prevSentence = int(sentenceNumber) # current sentence is the now previous sentence
							userStep = 1
							sentenceOrder += 1;

						# Write Sentence Order column
						worksheet.write(row, SentenceOrderIndex, sentenceOrder)
						
						worksheet.write(row, UserStepIndex, userStep)
								
				# Write the Study Context section
				studyContext = context.find('Study_Context')

				if studyContext is not None:
					# Write School Code column
					schoolCode = studyContext.find('School_Code').text
					worksheet.write(row, SchoolCodeIndex, schoolCode)
					
					# Write Condition column
					condition = studyContext.find('Condition').text
					worksheet.write(row, ConditionIndex, condition)

					# Write Day column
					day = studyContext.find('Study_Day').text
					worksheet.write(row, StudyDayIndex, day)

					# Write Participant Code column
					participantID = studyContext.find('Participant_Code').text
					worksheet.write(row, ParticipantCodeIndex, participantID)
					
					if prevParticipantID != participantID:
						sentenceOrder = -1; # Reset sentence order for new participant ID
						prevParticipantID = participantID

					# Write Experimenter column
					experimenter = studyContext.find('Experimenter_Name').text
					worksheet.write(row, ExperimenterNameIndex, experimenter)

					# Write Language column
					language = studyContext.find('Language').text
					worksheet.write(row, StudyLanguageIndex, language)
					
					# Write Timestamp columns
					timeStamp = studyContext.find('Timestamp')
					
					if timeStamp is not None:
						# Write Date column
						date = timeStamp.find('Date').text
						worksheet.write(row, DateIndex, date)
					
						# Write Time column
						time = timeStamp.find('Time').text
						worksheet.write(row, TimeIndex, time)

						if sentenceStartTime == "":
							sentenceStartTime = time
						
						else:
							sentenceElapsedTime = datetime.strptime(time, FMT) - datetime.strptime(sentenceStartTime, FMT)
				
				# Write Chapter Status column
				if ChapterEndRow != -1:
					for x in range(ChapterBeginRow, ChapterEndRow + 1):
						worksheet.write(x, ChapterStatusIndex, 'Complete')
				
					ChapterBeginRow = ChapterEndRow + 1
					ChapterEndRow = -1

				else:
					worksheet.write(row, ChapterStatusIndex, 'Incomplete')
				
				# Write Assessment Status column
				if AssessmentEndRow != -1:
					for x in range(AssessmentBeginRow, AssessmentEndRow + 1):
						worksheet.write(x, AssessmentStatusIndex, 'Complete')
						worksheet.write(x, ChapterStatusIndex, 'Complete')
				
					AssessmentBeginRow = AssessmentEndRow + 1
					ChapterBeginRow = AssessmentEndRow + 1
					AssessmentEndRow = -1
					ChapterEndRow = -1

				else:
					worksheet.write(row, AssessmentStatusIndex, 'Incomplete')

				#Go to next row and reset to first column
				row += 1
				col = 0

# Close the workbook
workbook.close()
