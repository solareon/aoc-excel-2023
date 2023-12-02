# aoc-excel-2023
Advent of Code 2023 with only Excel
Excel is a still programming language. This repo archives my attempt to do all of the Advent of Code 2023 entirely in Excel without the use of VBA.

# Day 1
Alright lets break out the maps and lambda right off the rip. Of special note is the anti-AI samples that don't include all the corner cases. Well played Mr. AoC man.
Part 1 is a simple search from left to right to find the first number. To find the second number we reverse the string and search again.
Part 2 we need to still find the location of the numbers but we then need to search for the words within the string to replace with numbers. After finding the first number we prepend it to the original string and then proceed to reverse and search for the words again but this time with everything reversed. After completing the search and replace we then reassemble the first and last number and perform the addition.

# Day 2
TEXTSPLIT, TEXTAFTER, and TEXTBEFORE to the rescue. Some clever error checking with IFNA allows to correct for cases where the found text is in a weird spot. Only dragged formula was the initial split since for some reason trying to BYROW it wouldn't fill to the right. In my input there was no more than 6 pulls from the bag so thats as wide as the formula gets. Part 1 ANDs the checks then finds the game id from the first split. This could be optimized a bit but overall much cleaner than day 1.