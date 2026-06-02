print('Welcome to Aydens Quiz')
answer=input('Are you ready to play my Quiz ? (yes/no) :')
score=0
total_questions=3
 
if answer.lower()=='yes':
    answer=input('Question 1: What is Aydens favourite YouTube Channel?')
    if answer.lower()=='lanky box':
        score += 1
        print('correct')
    else:
        print('Wrong Answer :(')
 
 
    answer=input('Question 2: What is Aydens favourite sport? ')
    if answer.lower()=='soccer':
        score += 1
        print('correct')
    else:
        print('Wrong Answer :(')
 
    answer=input('Question 3: What is Aydens favourite animal?')
    if answer.lower()=='cheetah':
        score += 1
        print('correct')
    else:
        print('Wrong Answer :(')
 
print('Thankyou for Playing my quiz game, you attempted',score,"questions correctly!")
mark=(score/total_questions)*100
print('Marks obtained:',mark)
print('See ya!')