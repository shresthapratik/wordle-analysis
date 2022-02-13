#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# =============================================================================
# Created By  : Pratik Shrestha
# Created Date: Feb 13 08:42:13 2022
# =============================================================================
"""The Python File has been for:
    Analysing Wordle List to find best opening word"""
# =============================================================================

    
# =============================================================================
# Imports
# =============================================================================

import sys
import json
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
# =============================================================================
# Loading and converting data to appropriate format
# =============================================================================


# Opening JSON file
jsonObj = open('wordlist.json',)

# Returns JSON object as a dictionary
words = json.load(jsonObj)

#Converting all words into upper case
for i in range(len(words['list'])):
    words['list'][i] = words['list'][i].upper()
    
#Printing Total # of Words in the List
total_words = len(words['list'])
print("Total No. of Words in the List: " +str(total_words))
print("")


# =============================================================================
# Analysis 1 - Find the no. of occurences of each letter in the list
# =============================================================================

#Dictionary to store count of letters in all the word list
lettercount = {}

#Loop to count no. of each letter in all the world list
for word in words['list']:
    for letter in word:
        if letter in lettercount:
            lettercount[letter] += 1
        else:
            lettercount[letter] = 1

#Converting to dataframe
lettercount_df = pd.DataFrame(lettercount, index=[0])

#Transforming data to required format
lettercount_df = lettercount_df.T.reset_index()

#Renaming columns to understandable format
lettercount_df.columns = ['letter', 'count']

#Print Count of each letter
print("Which letter occurs the most?")
print(lettercount_df.sort_values('count', ascending=False))

with pd.ExcelWriter("Wordle Analysis.xlsx") as writer:
    lettercount_df.to_excel(writer, sheet_name = "OccurencesByLetter", index=False)

# =============================================================================
# Analysis 2 - Find which letters do words in the list mostly start or end with
# =============================================================================

print("Which letter does most words start with?")

##---------STARTING LETTER ANALYSIS ----------------------##

#Dictionary to store count of letters in all the word list
start_letter_count = {}

#Loop to count no. of each letter in all the world list
for word in words['list']:
    if word[0:1] in start_letter_count:
            start_letter_count[word[0:1]] += 1
    else:
        start_letter_count[word[0:1]] = 1

#Converting to dataframe
start_letter_count_df = pd.DataFrame(start_letter_count, index=[0])

#Transforming data to required format
start_letter_count_df = start_letter_count_df.T.reset_index()

#Renaming columns to understandle format
start_letter_count_df.columns = ['letters', 'count']

#Sorting values by letter count and printing
print(start_letter_count_df.sort_values('count', ascending=False))
print("")

#Exporting result to Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:
    start_letter_count_df.sort_values('count', ascending=False).to_excel(writer, sheet_name = "StartingLetterAnalysis", index=False)

##---------ENDING LETTER ANALYSIS ----------------------##

#Dictionary to store count of letters in all the word list
end_letter_count = {}

print("Which letter does most words end with?")

#Loop to count no. of each letter in all the world list
for word in words['list']:
    if word[-1] in end_letter_count:
            end_letter_count[word[-1]] += 1
    else:
        end_letter_count[word[-1]] = 1

#Converting to dataframe
end_letter_count_df = pd.DataFrame(end_letter_count, index=[0])

#Transforming data to required format
end_letter_count_df = end_letter_count_df.T.reset_index()

#Renaming columns to understandle format
end_letter_count_df.columns = ['letters', 'count']

#Sorting values by letter count and printing
print(end_letter_count_df.sort_values('count', ascending=False))
print("")

#Exporting result to Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:
    end_letter_count_df.sort_values('count', ascending=False).to_excel(writer, sheet_name = "EndingLetterAnalysis", index=False)


# =============================================================================
# Analysis 3 - Algorithm to find the average letter revealed for each word
# and the probability that each word reveals at least 2 words
# =============================================================================

#Starting the clock to calculate how long the algorithm takes to compute
start_time = datetime.now()

print("Calculating avg number of letters revealed for each word")
print("along with probability of each word revealing at least 2 letters")

#Dictionary to store average letter revealed for each word
avg_letter_revealed = {}

#Dictionary to store probability of at least two letters revealed for each word
prob_atleast_two_letters = {}

#Counter for checking which word is currently being processed on cmd line
i = 0

#Counter to find the complexity of algorithm by counting no. of times in the loop
no_of_step_in_loop = 0

#Main loop to find the avg no. of letters revealed and probability
##Initially we label each word that we start to check with other words as 'check_word'
for check_word in words['list']:
    i = i+1
    p = round(i/total_words * 100, 2) #Converting to percentage
    
    #printing on screen to check with word is being processed, and how many % is completed
    sys.stdout.write("\rProcessing %s" % check_word + " Completed %s" % p + "%")
    sys.stdout.flush()
    
    #Array to store no. of letters revealed for each check word
    revealed_list = []
    
    #Second loop that runs through each word list and we label this as 'compare_word'
    for compare_word in words['list']:
        no_of_letters_revealed = 0
        
        #Loop to go through each letter in the check word
        for letter_check_word in check_word:
            no_of_step_in_loop+=1
            
            #Checking if the letter of check word is in compare word
            if(letter_check_word in compare_word):
                no_of_letters_revealed += 1
                #Removing the letter from compare word so that check word doesn't check for same letter again
                compare_word = compare_word.replace(letter_check_word, "", 1)
        
        #Adding no. of letters revealed to our array
        revealed_list.append(no_of_letters_revealed)
        
    #Finding the average letter revealed for each check word    
    avg_letter_revealed[check_word] = np.mean(revealed_list)
    
    #Converting revealed list array to numpy array
    numpy_revealed_list = np.array(revealed_list)
    
    #Finding probability that at least two letters are revealed
    prob_atleast_two_letters[check_word] = len(numpy_revealed_list[numpy_revealed_list>1])/len(words['list'])
                    
end_time = datetime.now() #To indicate the loop has ended
loop_time = end_time - start_time #Calculating time to calculate
print("\nLoop Time: "+ str(loop_time.total_seconds()) + " seconds")                    
print("No. of steps spent in the loop: " + str(no_of_step_in_loop))


#Converting to dataframe
average_letter_revealed_df = pd.DataFrame(avg_letter_revealed, index=[0])
prob_atleast_two_letters_df = pd.DataFrame(prob_atleast_two_letters, index=[0])

#Transforming data to required format
average_letter_revealed_df = average_letter_revealed_df.T.reset_index()
prob_atleast_two_letters_df= prob_atleast_two_letters_df.T.reset_index()

#Renaming columns to understandable format
average_letter_revealed_df.columns = ['letter', 'avg_letter_revealed']
prob_atleast_two_letters_df.columns = ['letter', 'prob_at_least_two_letters_revealed']

#Printing top 5 words by average letter revealed
print(average_letter_revealed_df.sort_values('avg_letter_revealed', ascending=False)[0:5])
print("")

#Printing top 5 words by probability that at least 2 letters will be revealed
print(prob_atleast_two_letters_df.sort_values('prob_at_least_two_letters_revealed', ascending=False)[0:5])
print("")

#Storing data in Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:    average_letter_revealed_df.sort_values('avg_letter_revealed', ascending=False).to_excel(writer, sheet_name = 'AvgLetterRevealed', index=False)

with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:    prob_atleast_two_letters_df.sort_values('prob_at_least_two_letters_revealed', ascending=False).to_excel(writer, sheet_name = 'Prob_AtLeast2', index=False)

# =============================================================================
# From Analysis 3, we find that 'SOARE', 'AROSE', 'AEROS' are top 3 words.
# But we want to find the best among them.
#
# Analysis 4 - Algorithm to find the probability that at least one letter is
# revealed in its correct spot (for 'SOARE', 'AROSE', 'AEROS')
# =============================================================================

#Starting the clock to calculate how long the algorithm takes to compute
start_time = datetime.now()

#Counter to find the complexity of algorithm by counting no. of times in the loop
no_of_step_in_loop = 0

print("Calculating probability that letter is revealed in correct spot for each word ")

#Dictionary to store probability that letters are revealed in correct spot for each word
prob_letter_correct_spot = {}

#Counter for checking which word is currently being processed on cmd line
i = 0

#Main loop to find the probability that letters are revealed in correct spot

for check_word in ['SOARE', 'AROSE', 'AEROS']:
    i = i+1
    sys.stdout.write("\rProcessing %s" % check_word)
    sys.stdout.flush()
    correct_spot_list = []
    for compare_word in words['list']:
        no_of_letters_revealed = 0
        j = 0
        for j in range (0,4):
            if(check_word[j] == compare_word[j]):
                no_of_letters_revealed += 1
        correct_spot_list.append(no_of_letters_revealed)
        j+=1
    numpy_correct_spot_list = np.array(correct_spot_list)    
    prob_letter_correct_spot[check_word] = len(numpy_correct_spot_list[numpy_correct_spot_list>0])/len(words['list'])
                    
end_time = datetime.now()
loop_time = end_time - start_time
print("Loop Time: "+ str(loop_time.total_seconds()) + " seconds")                    
print("No. of steps spent in the loop: " + str(no_of_step_in_loop))


#Converting to dataframe
prob_letter_correct_spot_df = pd.DataFrame(prob_letter_correct_spot, index=[0])

#Transforming data to required format
prob_letter_correct_spot_df = prob_letter_correct_spot_df.T.reset_index()

#Renaming columns to understandable format
prob_letter_correct_spot_df.columns = ['word', 'prob']

#Print Probabilities of each word
print("Probability of each word that it reveals at least 1 word")
print(prob_letter_correct_spot_df.sort_values('prob', ascending=False))

#Exporting to Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:    
    prob_letter_correct_spot_df.to_excel(writer, sheet_name = 'Prob_AtLeast1_CorrectSpot', index=False)

# =============================================================================
# From Analysis 4, we find the best word is SOARE

# Analysis 5 - Finding probability that SOARE reveals at least 1, 2, 3, 4, 5
# or no letters
# =============================================================================

no_of_step_in_loop = 0

prob_letter = {}

start_time = datetime.now()
#Loop to find probability that SOARE reveals at least n letters
for check_word in ['SOARE']:
    
    #Array to store no. of letters revealed for each check word
    revealed_list = []
    
    #Second loop that runs through each word list and we label this as 'compare_word'
    for compare_word in words['list']:
        no_of_letters_revealed = 0
        
        #Loop to go through each letter in the check word
        for letter_check_word in check_word:
            no_of_step_in_loop+=1
            
            #Checking if the letter of check word is in compare word
            if(letter_check_word in compare_word):
                no_of_letters_revealed += 1
                #Removing the letter from compare word so that check word doesn't check for same letter again
                compare_word = compare_word.replace(letter_check_word, "", 1)
        
        #Adding no. of letters revealed to our array
        revealed_list.append(no_of_letters_revealed)
        
    #Converting revealed list array to numpy array
    numpy_revealed_list = np.array(revealed_list)
    
    for i in range (0, 5):
        #Finding probability that at no letters are revealed
        prob_letter[check_word + '_' + str(i)] = len(numpy_revealed_list[numpy_revealed_list>i])/len(words['list'])
end_time = datetime.now()
loop_time = end_time - start_time
print("\nLoop Time: "+ str(loop_time.total_seconds()) + " seconds")                    
print("No. of steps spent in the loop: " + str(no_of_step_in_loop))

#Converting to dataframe
SOARE_prob_df = pd.DataFrame(prob_letter, index=[0])

#Transforming data to required format
SOARE_prob_df = SOARE_prob_df.T.reset_index()

#Renaming columns to understandable format
SOARE_prob_df.columns = ['n', 'prob_at_least_n_letters']

#Print Probabilities of each word
print("Probability that SOARE reveals n letters")
print(SOARE_prob_df.sort_values('prob_at_least_n_letters'))
print("")

#Exporting to Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:
    prob_letter_correct_spot_df.to_excel(writer, sheet_name = 'Prob_SOARE', index=False)


# =============================================================================
# From Analysis 5, we find that there is still 4% chance that SOARE reveals 0
# words. We need run further analysis to find another word without SOARE
#
# Analysis 6 - Finding words that reveal the most letters on average
# without the letters 'S', 'O', 'A', 'R', 'E'
# =============================================================================

#In case the WORDLE of the day doesn't contain 'S', 'O', 'A', 'R', 'E', finding a word without these letters that reveal the most letters
avg_letter_without_soare = average_letter_revealed_df[~average_letter_revealed_df['letter'].str.contains('S|O|A|E|R')]

print("Top 5 words by avg letters revealed without 'S','O','A','R','E'")
print(avg_letter_without_soare.sort_values('avg_letter_revealed', ascending=False)[0:5])

#Storing data in Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:
    avg_letter_without_soare.to_excel(writer, sheet_name='AvgLetters_without_SOARE', index=False)

# =============================================================================
# From Analysis 5, we find we need to choose between UNTIL vs UNLIT
#
# Analysis 6 - Finding words probability that UNTIL and UNLIT reveals at least
# letter in correct spot
# =============================================================================

#Dictionary to store probability that letters are revealed in correct spot for each word
prob_letter_correct_spot = {}

#Counter for checking which word is currently being processed on cmd line
i = 0

#Main loop to find the probability that letters are revealed in correct spot

for check_word in ['UNLIT', 'UNTIL']:
    i = i+1
    sys.stdout.write("\rProcessing %s" % check_word)
    sys.stdout.flush()
    correct_spot_list = []
    for compare_word in words['list']:
        no_of_letters_revealed = 0
        j = 0
        for j in range (0,4):
            if(check_word[j] == compare_word[j]):
                no_of_letters_revealed += 1
        correct_spot_list.append(no_of_letters_revealed)
        j+=1
    numpy_correct_spot_list = np.array(correct_spot_list)    
    prob_letter_correct_spot[check_word] = len(numpy_correct_spot_list[numpy_correct_spot_list>0])/len(words['list'])
                    
end_time = datetime.now()
loop_time = end_time - start_time
print("Loop Time: "+ str(loop_time.total_seconds()) + " seconds")                    
print("No. of steps spent in the loop: " + str(no_of_step_in_loop))

#Converting to dataframe
prob_letter_correct_spot_df = pd.DataFrame(prob_letter_correct_spot, index=[0])

#Transforming data to required format
prob_letter_correct_spot_df = prob_letter_correct_spot_df.T.reset_index()

#Renaming columns to understandable format
prob_letter_correct_spot_df.columns = ['word', 'prob']

#Print Probabilities of each word
print("Probability of each word that it reveals at least 1 word")
print(prob_letter_correct_spot_df.sort_values('prob', ascending=False))

#Exporting to Excel
with pd.ExcelWriter("Wordle Analysis.xlsx", mode="a", engine="openpyxl") as writer:
    prob_letter_correct_spot_df.to_excel(writer, sheet_name = 'UNLITvsUNTIL_>=1_correctspot', index=False)
