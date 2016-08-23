import os
import glob
import xml.etree.ElementTree as ET
import xlsxwriter

worksheet = None

# Contains "<bookKey>" -> column number to write order in which book was read
bookOrderCol = {}

# Records whether book was completed or not
bookCompletion = {  "Monkey1": False,
					"Bottled": False,
					"Monkey2": False,
					"BestFarm": False,
					"Native": False,
					"Physics": False,
					"Celebration": False,
					"Disasters": False,
					"House": False
				 }


# Contains "<bookKey> CH<chapter number>" -> column number to write condition in which chapter was read
chapterConditionCol = {}

# Contains "<bookKey> CH<chapter number> Q<question number>" -> column number to write number of attempts
chapterIMMenuCol = {}

# Each book key corresponds to array containing number of immenus per chapter
numIMMenusPerChapterByBook = {    	"Monkey1": [2],
									"Bottled": [2, 2, 2, 2, 2],
									"Monkey2": [2],
									"BestFarm": [0, 0, 0, 2, 2, 2, 2],
									"Native": [0, 0, 0, 2, 2, 2, 2],
									"Physics": [0, 0, 0, 2, 2, 2, 2],
									"Celebration": [2, 0, 2, 2, 2, 2],
									"Disasters": [0, 0, 2, 2, 2, 2],
									"House": [2, 2, 2, 2, 2, 2]
								}
								
								# Each book key corresponds to array containing number of assessment questions per chapter
numQuestionsPerChapterByBook = {    "Monkey1": [2],
                                    "Bottled": [8, 7, 8, 8, 7],
                                    "Monkey2": [2],
                                    "BestFarm": [7, 7, 6, 6, 6, 6, 6],
                                    "Native": [3, 7, 7, 6, 9, 6, 6],
                                    "Physics": [5, 5, 6, 7, 6, 6, 7],
                                    "Celebration": [7, 7, 6, 7, 7, 7],
                                    "Disasters": [3, 7, 5, 6, 8, 8],
                                    "House": [7, 6, 8, 8, 6, 7]
                                }

# Writes labels in the first row of the Excel sheet
def createLabels():
	# Write labels
	worksheet.write(0, 0, "Participant ID", bold)
	
	labelCol = 1
	
	for bookKey, numImMenusPerChapter in numIMMenusPerChapterByBook.items():
		# Generate label
		label = bookKey + " Order"
		
		# Write label
		worksheet.write(0, labelCol, label, bold)
		
		# Record column number
		bookOrderCol[label] = labelCol
		
		labelCol += 1
		
		for chapterNum in range(1, len(numImMenusPerChapter) + 1):
			# Generate label
			label = "{0} CH{1} Condition".format(bookKey, chapterNum)
			
			# Write label
			worksheet.write(0, labelCol, label, bold)
			
			# Record column number
			chapterConditionCol[label] = labelCol
			
			labelCol += 1
			
			numIMMenus = numImMenusPerChapter[chapterNum - 1]
			
			for IMMenuNum in range(1, numIMMenus + 1):
				# Generate label
				label = "{0} CH{1} Menu{2}".format(bookKey, chapterNum, IMMenuNum)
				
				# Write label
				worksheet.write(0, labelCol, label, bold)
				
				# Record column number
				chapterIMMenuCol[label] = labelCol
				
				labelCol += 1

# Returns the appropriate book key for the given book title
def getBookKey(bookTitle):
	bookKey = ""
	
	if bookTitle == "Introduction to EMBRACE":
		bookKey = "Monkey1"
			
	elif bookTitle == "Bottled Up Joy":
		bookKey = "Bottled"
					
	elif bookTitle == "Second Introduction to EMBRACE":
		bookKey = "Monkey2"
							
	elif bookTitle == "The Best Farm":
		bookKey = "BestFarm"
					
	elif bookTitle == "Native American Homes":
		bookKey = "Native"
							
	elif bookTitle == "How Objects Move":
		bookKey = "Physics"
				
	elif bookTitle == "A Celebration to Remember":
		bookKey = "Celebration"
							
	elif bookTitle == "Natural Disasters":
		bookKey = "Disasters"
						
	elif bookTitle == "The Lopez Family Mystery":
		bookKey = "House"

	return bookKey

# Resets all book completions to false
def resetBookCompletion():
	for bookKey in bookCompletion:
		bookCompletion[bookKey] = False

if __name__ == "__main__":
	# Create Excel file
	workbook = xlsxwriter.Workbook("immenu_data.xlsx")
	worksheet = workbook.add_worksheet()

	# Add a bold format
	bold = workbook.add_format({"bold": True})

	createLabels()

	# Start writing data below labels
	row = 1

	# Go through every .xml log file in each directory
	for root, dirs, files in os.walk("./Log Data"):
		for dir in dirs:
			# Write participant ID (comes from directory name)
			participantID = dir
			worksheet.write(row, 0, participantID)
			
			print "\n*** participantID: " + participantID
			
			bookNum = 0
			prevBookTitle = ""
			
			dirPath = os.path.join(root, dir)
			
			for file in glob.glob(os.path.join(dirPath, "*.xml")):
				print "*** file: " + file
				
				tree = ET.parse(file)
				treeRoot = tree.getroot()
				
				bookKey = ""
				numChapters = 0
				
				chapterNum = 0
				prevChapterTitle = ""
				chapterKey = ""
				chapterCondition = ""
				updatedChapterCondition = False
				
				numAttempts = 0
				currentNumMenus = 0
				
				for child in treeRoot:
					action = child.find("Action").text
					selection = child.find("Selection").text
					
					if action == "Load Book":
						bookTitle = child.find("Input").find("Book_Title").text
						bookKey = getBookKey(bookTitle)
						
						# Ignore any extra log data if book was already completed
						if bookCompletion[bookKey] == False:
							# NOTE: This check was added in case the user opened the book more than once without completing it
							if bookTitle != prevBookTitle:
								prevBookTitle = bookTitle
								bookNum += 1
							
								# Write book order in appropriate column
								key = bookKey + " Order"
								col = bookOrderCol[key]
								worksheet.write(row, col, bookNum)
								
								print "*** key: {0}   bookNum: {1}".format(key, bookNum)
								
								# Record number of chapters in this book
								numChapters = len(numIMMenusPerChapterByBook[bookKey])
								
								chapterNum = 0
								chapterTitle = ""
								chapterKey = ""
								chapterCondition = ""
								
								numAttempts = 0
				
					# Ignore any extra log data if book was already completed
					elif bookKey != "" and bookCompletion[bookKey] == False:
						if action == "Load Chapter":
							chapterTitle = child.find("Input").find("Chapter_Title").text
							
							# NOTE: This check was added in case the user opened the chapter more than once without completing it
							if chapterTitle != prevChapterTitle:
								
								prevChapterTitle = chapterTitle
								chapterNum += 1
								chapterKey = "CH" + str(chapterNum)
								
								# Write initial chapter condition in appropriate column
								chapterCondition = child.find("Context").find("Study_Context").find("Condition").text
								key = "{0} {1} Condition".format(bookKey, chapterKey)
								col = chapterConditionCol[key]
								worksheet.write(row, col, chapterCondition)
					
								print "*** key: {0}   chapterCondition: {1}".format(key, chapterCondition)
									
								updatedChapterCondition = False
					
						# Sentence audio indicates that the system was reading to the user
						elif action == "Sentence Audio":
							# Make sure chapter condition is only updated once
							if updatedChapterCondition == False:
								updatedChapterCondition = True
								
								if chapterCondition == "R":
									# User was listening only
									chapterCondition = "L"

								else:
									# LPM = Listen + PM; LIM = Listen + IM
									chapterCondition = "L" + chapterCondition
							
								# Write updated chapter condition in appropriate column
								key = "{0} {1} Condition".format(bookKey, chapterKey)
								col = chapterConditionCol[key]
								worksheet.write(row, col, chapterCondition)
								
								print "*** key: {0}   chapterCondition: {1}".format(key, chapterCondition)
					
						elif action == "Verify Action" and selection == "Select Menu Item":
							# Record attempt
							numAttempts += 1
						
							# Get verification
							verification = child.find("Input").find("Verification").text
							
							if verification == "Correct":

								# Get menu number
								currentNumMenus += 1
								imMenuNumKey = "Menu" + str(currentNumMenus)
								
								# Create key
								key = "{0} {1} {2}".format(bookKey, chapterKey, imMenuNumKey)
								
								# Write number of attempts in appropriate column
								col = chapterIMMenuCol[key]
								worksheet.write(row, col, numAttempts)
								
								print "*** key: {0}   numAttempts: {1}".format(key, numAttempts)
			
								numAttempts = 0
								if currentNumMenus == 2:
									currentNumMenus = 0
								
						elif action == "Verify Assessment Answer":
						
							# Get verification
							verification = child.find("Input").find("Verification").text
							
							if verification == "Correct":

								# Get question number
								questionNum = child.find("Context").find("Assessment_Context").find("Assessment_Step_Number").text
								
								# Record book as completed
								if chapterNum == numChapters and int(questionNum) == numQuestionsPerChapterByBook[bookKey][chapterNum - 1]:
									bookCompletion[bookKey] = True
			resetBookCompletion()
			
			# New row for next participant
			row += 1

	# Close the workbook
	workbook.close()