# aoc-excel-2023
Advent of Code 2023 with only Excel.

Excel is a still programming language. This repo archives my attempt to do all of the Advent of Code 2023 entirely in Excel without the use of VBA.

# Day 1
Alright lets break out the maps and lambda right off the rip. Of special note is the anti-AI samples that don't include all the corner cases. Well played Mr. AoC man.
Part 1 is a simple search from left to right to find the first number. To find the second number we reverse the string and search again.
Part 2 we need to still find the location of the numbers but we then need to search for the words within the string to replace with numbers. After finding the first number we prepend it to the original string and then proceed to reverse and search for the words again but this time with everything reversed. After completing the search and replace we then reassemble the first and last number and perform the addition.

# Day 2
TEXTSPLIT, TEXTAFTER, and TEXTBEFORE to the rescue. Some clever error checking with IFNA allows to correct for cases where the found text is in a weird spot. Only dragged formula was the initial split since for some reason trying to BYROW it wouldn't fill to the right. In my input there was no more than 6 pulls from the bag so thats as wide as the formula gets. Part 1 ANDs the checks then finds the game id from the first split. This could be optimized a bit but overall much cleaner than day 1.

# Day 3
Recursion. The enemy of Excel and spreadsheets. Part 1 wasn't too difficult but getting the pairs for part 2 has eluded me at this point. I've paired the data down but just need to match them up correctly. I have some ideas but not the time to implement

# Day 4
SEQUENCE go brrrrrrr. Unpack the cards, XMATCH and SUM to get the hits per card. MAP down the array with Power subtracting 1 for the first card. Sum this all together and Part 1 down. Part 2 required a bit more thinking to find the correct offset variables for walking across the cards but once solved fill down the array and sum it up and add the number of cards in the original set. My input was 220 so I hard coded it.

# Day 5
Words critically hit you for over 9000. Got too busy to tinker with this one

# Day 6
Big numbers you say? Part 1 solution was simple and elegant. Part 2 required some outside of the box thinking to break it into 8 smaller problems that are within Excel's calculation limits. Sure it isn't golfed but is beautiful in it's own way.
