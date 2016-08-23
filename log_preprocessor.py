import os
import re

for root, dirs, files in os.walk("."):
    for dir in dirs:
        for file in os.listdir(dir):
            newFile = file.replace(".txt", ".xml")

            iFile = open(os.path.join(dir, file), "r")
            oFile = open(os.path.join(dir, newFile), "w")

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