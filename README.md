# aoc-excel-2023
Advent of Code 2023 with only Excel
Excel is a still programming language. This repo archives my attempt to do all of the Advent of Code 2023 entirely in Excel without the use of VBA.

# Day 1
Alright lets break out the maps and lambda right off the rip. Of special note is the anti-AI samples that don't include all the corner cases. Well played Mr. AoC man.
Part 1 is a simple search from left to right to find the first number. To find the second number we reverse the string and search again.
Part 2 we need to still find the location of the numbers but we then need to search for the words within the string to replace with numbers. After finding the first number we prepend it to the original string and then proceed to reverse and search for the words again but this time with everything reversed. After completing the search and replace we then reassemble the first and last number and perform the addition.
