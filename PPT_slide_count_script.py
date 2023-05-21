import os
import re
import zipfile
import argparse


parser = argparse.ArgumentParser(description="Count the number of slides in .pptx decks and return a count")
parser.add_argument("-s", "--summary", help="Generate a summary of all power point decks processed", action="store_true")
parser.add_argument("-f", "--files", help="List of files to process", type=argparse.FileType("r"), nargs="*")
args = parser.parse_args()

# Initializing an empty dictionary to hold the deck name and  count of the number of slides in the deck
power_point_decks = {}

# Iterating through the files
if(args.files):
    for file in args.files:
        # If the file ends with ".pptx", will be added to the power_point_decks dictionary
        if os.path.abspath(file.name).endswith('.pptx'):
                power_point_decks[(os.path.abspath(file.name))] = 0
        # If the file does not end with ".pptx", will ignore it and print a message
        else:
            print("The file %s is not a .pptx file and will be ignored." % (file.name))

# Iterating through the items in the power_point_decks dictionary
for power_point_deck, slide_count in power_point_decks.items():
    try:
        # Read the deck as a zip file
        archive = zipfile.ZipFile(power_point_deck, 'r')
        # Get the list of files of the deck
        list_of_files = archive.namelist()
    # If the file cannot be open as a zip or if there is no namelist inside it, will print an error message and skip the file
    except Exception as e:
        print("Error reading %s (%s). Count will be 0." % (os.path.basename(power_point_deck), e))
    else:
        # Iterate through the list of files in the zip file's namelist
        for file_entry in list_of_files:
                # For each entry that matches the path "ppt/slides/slide", will increment the count of slides for that deck by 1
                if(re.findall("ppt/slides/slide", file_entry)):
                        power_point_decks[power_point_deck] += 1

# Print out the header for the table of slide counts
print("Slides\tDeck")

# Iterate through a sorted version of the power_point_decks dictionary, print out the slide count and the file name of each deck
for power_point_deck, slide_count in sorted(power_point_decks.items()):
    print("%s\t%s" % (slide_count, os.path.basename(power_point_deck)))

# Display summary, count the total number of slides by summing up the keys in the power_point_decks dictionary
# and report that as well as the the length of the power_point_decks dictionary for an overall summary of the total number of slides
# and decks that were processed
if (args.summary):
    print("- - - - -")
    total = 0
    # Iterate through the values in the power_point_decks dictionary and add up all of the counts
    for slide_count in power_point_decks.values():
        total += slide_count
    print("There are %s total slides in %s decks." % (total, len(power_point_decks)))