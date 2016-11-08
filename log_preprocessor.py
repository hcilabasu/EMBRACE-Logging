import os
import re

# Check that path is correct before running this script!
PATH_TO_LOG_FILES = "."

for root, dirs, files in os.walk(PATH_TO_LOG_FILES):
    for dir in dirs:
        dirPath = os.path.join(root, dir)
        
        for file in os.listdir(dirPath):
            if file != ".DS_Store" and ".txt" in file:
                newFile = file.replace(".txt", ".xml")

                filePath = os.path.join(dirPath, file)
                newFilePath = os.path.join(dirPath, newFile)
                
                iFile = open(filePath, "r")
                oFile = open(newFilePath, "w")

                lineNumber = 0

                for line in iFile:
                    if lineNumber != 0:
                        # Find index of first > symbol
                        openR = line.find(">")
                        
                        # Replace spaces inside opening XML tag with underscores
                        line = re.sub(r'([^\s])\s([^\s])', r'\1_\2', line[:openR]) + line[openR:]
                        
                        # Find index of second < symbol
                        closeL = line.rfind("<")
                        
                        # Replace sapces inside closing XML tag with underscores
                        line = line[:closeL] + line[closeL:].replace(" ", "_")
                    
                    # Write modified line to file
                    oFile.write(line)

                    lineNumber += 1

                iFile.close()
                oFile.close()
